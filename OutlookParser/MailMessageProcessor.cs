using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using Gentex.ComponentTracker.Plugin;
using Itenso.Rtf;
using Itenso.Rtf.Converter.Html;
using Itenso.Rtf.Converter.Text;
using Itenso.Rtf.Interpreter;
using Itenso.Rtf.Parser;
using Itenso.Rtf.Support;
using Microsoft.VisualBasic;
using OutlookParser.Mime;

namespace OutlookParser
{
  public class MailMessageProcessor : IEmailParser, IEmailRenderer
  {
    private static string Normalize(string value, CompareOptions options = CompareOptions.IgnoreCase | CompareOptions.IgnoreWidth)
    {
      var returnVal = value;
      if ((options & CompareOptions.IgnoreCase) == CompareOptions.IgnoreCase)
      {
        returnVal = CultureInfo.CurrentCulture.TextInfo.ToLower(returnVal);
      }

      // Full-width forms seem to be the same characters regardless of culture.  Therefore, use the Japanese culture.
      if ((options & CompareOptions.IgnoreWidth) == CompareOptions.IgnoreWidth &&
          System.Text.Encoding.Unicode.GetByteCount(returnVal) >= (new StringInfo(returnVal).LengthInTextElements * 2))
      {
        returnVal = Strings.StrConv(returnVal, VbStrConv.Narrow, new CultureInfo("ja").LCID);
      }

      return returnVal;
    }

    public MailMessage Parse(string path, string defaultDomain)
    {
      return Parse(path, new AdEmailResolver(defaultDomain));
    }
    public MailMessage Parse(string path, IEmailResolver resolver)
    {
      if (Path.GetExtension(path).ToLowerInvariant() == ".msg")
      {
        return Parse(path, MailMessageFormat.msg, resolver);
      }
      else
      {
        return Parse(path, MailMessageFormat.eml, resolver);
      }
    }
    public MailMessage Parse(string path, MailMessageFormat format, IEmailResolver resolver)
    {
      using (Stream reader = File.Open(path, FileMode.Open, FileAccess.Read))
      {
        return Parse(reader, format, resolver);
      }
    }
    public MailMessage Parse(Stream stream, MailMessageFormat format, string defaultDomain)
    {
      return Parse(stream, format, new AdEmailResolver(defaultDomain));
    }
    public MailMessage Parse(Stream stream, MailMessageFormat format, IEmailResolver resolver)
    {
      switch (format)
      {
        case MailMessageFormat.msg:
          using (var msg = new OutlookStorage.Message(stream))
          {
            return Parse(msg, resolver);
          }
        default:
          return Parse(Mime.Mime.Parse(stream));
      }
    }
    private MailMessage Parse(OutlookStorage.Message msg, IEmailResolver resolver)
    {
      var result = new MailMessageData();
      MessageBuilder.ParseFillMailHeader(msg.Headers, result.Headers);

      foreach (var attach in msg.Attachments)
      {
        var newAttach = AttachmentHelper.CreateAttachment(new MemoryStream(attach.Data), attach.Filename, TransferEncoding.Base64);
        newAttach.ContentId = attach.ContentId;
        result.Attachments.Add(newAttach);
      }
      // Convert attached appointments to ical attachments and e-mails to eml
      MemoryStream stream;
      foreach (var attach in msg.Messages)
      {
        Attachment newAttach;
        switch (attach.Class)
        {
          case "IPM.Appointment":
            var iCal = new DDay.iCal.iCalendar();
            iCal.Version = "2.0";
            iCal.AddTimeZone(TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time"));
            var ev = new DDay.iCal.Event();
            ev.Organizer = new DDay.iCal.Organizer();
            ev.Organizer.CommonName = attach.From;
            if (resolver == null)
            {
              ev.Organizer.Value = new Uri("MAILTO:" + attach.FromEmail);
            }
            else
            {
              ev.Organizer.Value = new Uri("MAILTO:" + resolver.ProcessAddress(attach.From, attach.FromEmail).Address);
            }
            ev.UID = attach.MessageId;
            if (attach.BodyText == null)
            {
              // parse the rtf structure
              RtfParserListenerStructureBuilder structureBuilder = new RtfParserListenerStructureBuilder();
              RtfParser parser = new RtfParser(structureBuilder);
              parser.IgnoreContentAfterRootGroup = true; // support WordPad documents
              parser.Parse(new RtfSource(attach.BodyRtf));

              ev.Description = RenderRtfAsText(structureBuilder.StructureRoot);
            }
            else
            {
              ev.Description = attach.BodyText;
            }
            ev.Summary = attach.Subject;
            ev.Start = new DDay.iCal.iCalDateTime(attach.AppointmentStart, iCal.TimeZones[0].TZID);
            ev.End = new DDay.iCal.iCalDateTime(attach.AppointmentEnd, iCal.TimeZones[0].TZID);
            ev.Class = "PUBLIC";
            ev.Priority = 5;
            ev.DTStamp = new DDay.iCal.iCalDateTime(attach.SentTime, iCal.TimeZones[0].TZID);
            ev.Transparency = DDay.iCal.TransparencyType.Opaque;
            ev.Location = attach.AppointmentLocation;
            iCal.Events.Add(ev);
            var serializer = new DDay.iCal.Serialization.iCalendar.iCalendarSerializer();
            stream = new MemoryStream();
            serializer.Serialize(iCal, stream, Encoding.UTF8);
            stream.Position = 0;
            newAttach = AttachmentHelper.CreateAttachment(stream, attach.Subject + ".ics", TransferEncoding.Base64);
            result.Attachments.Add(newAttach);
            break;
          case "IPM.Note":
            var attachMsg = Parse(attach, resolver);
            stream = new MemoryStream();
            Render(attachMsg, stream);
            stream.Position = 0;
            newAttach = AttachmentHelper.CreateAttachment(stream, attachMsg.Subject + ".eml", TransferEncoding.Base64);
            result.Attachments.Add(newAttach);
            break;
        }
      }
      MailAddress addr;
      foreach (var recip in msg.Recipients)
      {
        if (resolver == null)
        {
          addr = new MailAddress(GetEmailFromRecip(recip), recip.DisplayName);
        }
        else
        {
          addr = resolver.ProcessAddress(recip.DisplayName, GetEmailFromRecip(recip));
        }
        switch (recip.Type)
        {
          case OutlookStorage.RecipientType.CC:
            result.CC.Add(addr);
            break;
          default:
            result.To.Add(addr);
            break;
        }
      }
      addr = null;
      if (resolver == null)
      {
        if (msg.FromEmail != null)
        {
          addr = new MailAddress(msg.FromEmail ?? "?@?", msg.From ?? "");
        }
      }
      else
      {
        addr = resolver.ProcessAddress(msg.From, msg.FromEmail);
      }
      if (addr != null) result.From = addr;

      var receivedTime = msg.ReceivedTime;
      if (receivedTime != null) result.ReceivedTime = receivedTime.Value;
      result.SentTime = msg.SentTime;
      result.Subject = msg.Subject;

      if (String.IsNullOrEmpty(msg.BodyHtml))
      {
        result.Body = msg.BodyText;
      }
      else
      {
        result.Body = msg.BodyHtml;
        result.IsBodyHtml = true;
        result.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(msg.BodyText, new ContentType("text/plain")));
      }
      if (!String.IsNullOrEmpty(msg.BodyRtf))
      {
        result.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(msg.BodyRtf, new ContentType("application/rtf")));
      }
      result.MessageId = msg.MessageId;
      return result;
    }
    private static MailMessage Parse(Mime.Mime eml)
    {
      var result = new MailMessageData();
      MessageBuilder.ParseFillMailHeader(eml.MainEntity.HeaderString, result.Headers);
      if (eml.MainEntity.ChildEntities != null)
      {
        foreach (var ent in eml.MainEntity.ChildEntities)
        {
          if (ent.ContentDisposition == ContentDisposition_enum.Attachment)
          {
            var newAttach = AttachmentHelper.CreateAttachment(new MemoryStream(ent.Data), ent.ContentDisposition_FileName, TransferEncoding.Base64);
            newAttach.ContentId = ent.ContentID;
            result.Attachments.Add(newAttach);
          }
        }
      }

      if (eml.MainEntity.Bcc != null)
      {
        foreach (var recip in eml.MainEntity.Bcc)
        {
          var maddr = recip as Mime.MailboxAddress;
          if (maddr != null)
          {
            result.Bcc.Add(new MailAddress(maddr.EmailAddress, maddr.DisplayName));
          }
        }
      }
      if (eml.MainEntity.Cc != null)
      {
        foreach (var recip in eml.MainEntity.Cc)
        {
          var maddr = recip as Mime.MailboxAddress;
          if (maddr != null)
          {
            result.CC.Add(new MailAddress(maddr.EmailAddress, maddr.DisplayName));
          }
        }
      }
      if (eml.MainEntity.From != null)
      {
        foreach (var recip in eml.MainEntity.From)
        {
          var maddr = recip as Mime.MailboxAddress;
          if (maddr != null)
          {
            result.From = new MailAddress(maddr.EmailAddress, maddr.DisplayName);
          }
        }
      }
      if (!String.IsNullOrEmpty(eml.MainEntity.Received))
      {
        result.ReceivedTime = MimeUtils.ParseDate(eml.MainEntity.Received.Substring(eml.MainEntity.Received.IndexOf(";") + 1).Trim());
      }
      result.SentTime = eml.MainEntity.Date;
      result.Subject = eml.MainEntity.Subject;
      if (eml.MainEntity.ReplyTo != null)
      {
        foreach (var recip in eml.MainEntity.ReplyTo)
        {
          var maddr = recip as Mime.MailboxAddress;
          if (maddr != null)
          {
            result.ReplyTo = new MailAddress(maddr.EmailAddress, maddr.DisplayName);
          }
        }
      }

      if (eml.MainEntity.To != null)
      {
        foreach (var recip in eml.MainEntity.To)
        {
          var maddr = recip as Mime.MailboxAddress;
          if (maddr != null)
          {
            result.To.Add(new MailAddress(maddr.EmailAddress, maddr.DisplayName));
          }
        }
      }

      if (String.IsNullOrEmpty(eml.BodyHtml))
      {
        result.Body = eml.BodyText;
      }
      else
      {
        result.Body = eml.BodyHtml;
        result.IsBodyHtml = true;
        result.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(eml.BodyText, new ContentType("text/plain")));
      }
      return result;
    }
    
    private static String GetEmailFromRecip(OutlookStorage.Recipient recip)
    {
      String result = recip.Email;
      if (string.IsNullOrEmpty(result))
      {
        var match = Regex.Match(recip.DisplayName, @"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b", RegexOptions.IgnoreCase);
        // Here we check the Match instance.
        if (match.Success)
        {
          // Finally, we get the Group value and display it.
          result = match.Groups[0].Value;
        }
      }
      if (string.IsNullOrEmpty(result))
      {
        result = recip.DisplayName + "@?";
      }
      else if (result.IndexOf('@') < 0)
      {
        result += "@?";
      }
      return result;
    }

    public string OutlookRtfToHtml(Stream stream)
    {
      // Parse the rtf content
      if (stream.CanSeek) stream.Position = 0;
      var content = ParseRtf(stream);

      // Try to extract encoded html from within the rtf (Outlook likes to do this)
      var builder = new StringBuilder();
      BuildHtmlContent(content, builder);
      if (builder.Length > 0 && string.Compare(builder.ToString(0, 5), "<html", true) == 0)
      {
        return builder.ToString();
      }
      else
      {
        return RenderRtfAsHtml(content);
      }
    }

    public void Render(MailMessage message, String path)
    {
      using (Stream writer = File.Open(path, FileMode.Create, FileAccess.Write))
      {
        Render(message, writer);
      }
    }
    public void Render(MailMessage message, Stream stream)
    {
      var m = new Mime.Mime();
      var main = m.MainEntity;
      string[] values;

      main.Header.Clear();
      for (var i = 0; i < message.Headers.Count; i++)
      {
        values = message.Headers.GetValues(i);
        foreach (var value in values)
        {
          main.Header.Add(message.Headers.GetKey(i), value);
        }
      }

      if (message.From != null)
      {
        main.From = new AddressList();
        main.From.Add(new Mime.MailboxAddress(message.From.DisplayName, message.From.Address));
      }

      if (message.To.Count > 0)
      {
        main.To = new AddressList();
        foreach (var addr in message.To)
        {
          main.To.Add(new Mime.MailboxAddress(addr.DisplayName, addr.Address));
        }
      }
      if (message.CC.Count > 0)
      {
        main.Cc = new AddressList();
        foreach (var addr in message.CC)
        {
          main.Cc.Add(new Mime.MailboxAddress(addr.DisplayName, addr.Address));
        }
      }
      if (message.Bcc.Count > 0)
      {
        main.Bcc = new AddressList();
        foreach (var addr in message.Bcc)
        {
          main.Bcc.Add(new Mime.MailboxAddress(addr.DisplayName, addr.Address));
        }
      }
      if (message.ReplyTo != null)
      {
        main.ReplyTo = new AddressList();
        main.ReplyTo.Add(new Mime.MailboxAddress(message.ReplyTo.DisplayName, message.ReplyTo.Address));
      }
      var data = message as MailMessageData;
      if (data == null)
      {
        main.Header.Add("X-Unsent", "1");
      }
      else
      {
        main.Date = data.SentTime;
        if (!string.IsNullOrEmpty(data.MessageId)) main.MessageID = data.MessageId;
      }
      main.Subject = message.Subject;
      main.ContentType = MediaType_enum.Multipart_mixed;

      var bodyEntity = main.ChildEntities.Add();
      bodyEntity.ContentTransferEncoding = ContentTransferEncoding_enum.Binary;
      bodyEntity.ContentType = MediaType_enum.Multipart_alternative;

      var otherEntity = bodyEntity.ChildEntities.Add();
      otherEntity.ContentTransferEncoding = ContentTransferEncoding_enum.Binary;
      if (message.IsBodyHtml)
      {
        otherEntity.ContentType = MediaType_enum.Text_html;
      }
      else
      {
        otherEntity.ContentType = MediaType_enum.Text_plain;
      }
      otherEntity.DataText = message.Body;

      foreach (var alt in message.AlternateViews)
      {
        otherEntity = bodyEntity.ChildEntities.Add();
        otherEntity.ContentTransferEncoding = ContentTransferEncoding_enum.Binary;
        switch (alt.ContentType.MediaType) {
          case "text/html":
            otherEntity.ContentType = MediaType_enum.Text_html;
            alt.ContentStream.Position = 0;
            otherEntity.DataFromStream(alt.ContentStream);
            break;
          case "text/rtf":
          case "application/rtf":
            otherEntity.ContentType = MediaType_enum.Text_html;

            // Copy the memory stream
            var newStream = new MemoryStream();
            CopyStream(alt.ContentStream, newStream);
            alt.ContentStream.Position = 0;
            newStream.Position = 0;

            // Parse the rtf content
            var content = ParseRtf(newStream);

            // Try to extract encoded html from within the rtf (Outlook likes to do this)
            var builder = new StringBuilder();
            BuildHtmlContent(content, builder);
            if (builder.Length > 0 && string.Compare(builder.ToString(0, 5), "<html", true) == 0)
            {
              otherEntity.DataText = builder.ToString();
            }
            else
            {
              otherEntity.DataText = RenderRtfAsHtml(content);
            }
            break;
          default:
            otherEntity.ContentType = MediaType_enum.Text_plain;
            alt.ContentStream.Position = 0;
            otherEntity.DataFromStream(alt.ContentStream);
            break;
        }
      }

      foreach (var attach in message.Attachments)
      {
        otherEntity = main.ChildEntities.Add();
        otherEntity.ContentID = attach.ContentId;
        otherEntity.ContentType = MediaType_enum.Application_octet_stream;
        otherEntity.ContentDisposition = ContentDisposition_enum.Attachment;
        otherEntity.ContentTransferEncoding = ContentTransferEncoding_enum.Base64;
        otherEntity.ContentDisposition_FileName = attach.ContentDisposition.FileName;
        otherEntity.DataFromStream(attach.ContentStream);
      }

      m.ToStream(stream);
    }

    private static void BuildHtmlContent(IRtfGroup content, StringBuilder builder)
    {
      bool doRender = false;
      foreach (IRtfElement elem in content.Contents)
      {
        switch (elem.Kind)
        {
          case RtfElementKind.Group:
            BuildHtmlContent((IRtfGroup)elem, builder);
            break;
          case RtfElementKind.Text:
            if (doRender) builder.Append(((IRtfText)elem).Text);
            break;
          case RtfElementKind.Tag:
            switch (((IRtfTag)elem).Name) {
              case "htmltag":
              case "htmlrtf":
                doRender = !doRender;
                break;
              case "par":
                if (doRender) builder.AppendLine();
                break;
            }
            break;
        }
      }
    }

    private static void CopyStream(Stream input, Stream output)
    {
      byte[] buffer = new byte[32768];
      int read;
      while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
      {
        output.Write(buffer, 0, read);
      }
    }

    private static IRtfGroup ParseRtf(Stream stream)
    {
      IRtfGroup rtfStructure;
      try
      {
        // parse the rtf structure
        RtfParserListenerStructureBuilder structureBuilder = new RtfParserListenerStructureBuilder();
        RtfParser parser = new RtfParser(structureBuilder);
        parser.IgnoreContentAfterRootGroup = true; // support WordPad documents
        parser.Parse(new RtfSource(stream));
        rtfStructure = structureBuilder.StructureRoot;
      }
      catch
      {
        return null;
      }

      return rtfStructure;
    } // ParseRtf

    private static string RenderRtfAsText(IRtfGroup contents)
    {
      var textConvertSettings = new RtfTextConvertSettings();
      textConvertSettings.BulletText = "-";
      var settings = new RtfInterpreterSettings() { IgnoreDuplicatedFonts = true, IgnoreUnknownFonts = true };
      var converter = new RtfTextConverter(textConvertSettings);
      RtfInterpreterTool.Interpret(contents, settings, converter);
      return converter.PlainText;
    }
    private static string RenderRtfAsText(IRtfSource source)
    {
      var textConvertSettings = new RtfTextConvertSettings();
      textConvertSettings.BulletText = "-";
      var settings = new RtfInterpreterSettings() { IgnoreDuplicatedFonts = true, IgnoreUnknownFonts = true };
      var converter = new RtfTextConverter(textConvertSettings);
      RtfInterpreterTool.Interpret(source, settings, converter);
      return converter.PlainText;
    }

    private static string RenderRtfAsHtml(IRtfGroup contents)
    {
      var settings = new RtfInterpreterSettings() { IgnoreDuplicatedFonts = true, IgnoreUnknownFonts = true };
      var rtfDocument = RtfInterpreterTool.BuildDoc(contents, settings);
      var htmlConvertSettings = new RtfHtmlConvertSettings();
      htmlConvertSettings.IsShowHiddenText = false;
      htmlConvertSettings.UseNonBreakingSpaces = false;
      htmlConvertSettings.ConvertScope = RtfHtmlConvertScope.All;
      
      RtfHtmlConverter htmlConverter = new RtfHtmlConverter(rtfDocument, htmlConvertSettings);
      return htmlConverter.Convert();
    }

    private class AttachmentHelper
    {
      public static Attachment CreateAttachment(string fileName, string displayName, TransferEncoding transferEncoding)
      {
        displayName = Normalize(displayName);
        fileName = Normalize(fileName);
        return CreateAttachment(new Attachment(fileName), displayName, transferEncoding);
      }
      public static Attachment CreateAttachment(Stream data, string displayName, TransferEncoding transferEncoding)
      {
        displayName = Normalize(displayName);
        return CreateAttachment(new Attachment(data, displayName), displayName, transferEncoding);
      }
      private static Attachment CreateAttachment(Attachment attachment, string displayName, TransferEncoding transferEncoding)
      {
        attachment.TransferEncoding = transferEncoding;
        attachment.ContentDisposition.FileName = displayName;
        attachment.Name = displayName;
        return attachment;
      }

      private static string SplitEncodedAttachmentName(string encodingtoken, string softbreak, int maxChunkLength, string encoded)
      {
        int splitLength = maxChunkLength - encodingtoken.Length - (softbreak.Length * 2);
        var parts = SplitByLength(encoded, splitLength);

        string encodedAttachmentName = encodingtoken;

        foreach (var part in parts)
          encodedAttachmentName += part + softbreak + encodingtoken;

        encodedAttachmentName = encodedAttachmentName.Remove(encodedAttachmentName.Length - encodingtoken.Length, encodingtoken.Length);
        return encodedAttachmentName;
      }

      private static IEnumerable<string> SplitByLength(string stringToSplit, int length)
      {
        while (stringToSplit.Length > length)
        {
          yield return stringToSplit.Substring(0, length);
          stringToSplit = stringToSplit.Substring(length);
        }

        if (stringToSplit.Length > 0) yield return stringToSplit;
      }
    }
  }
}
