using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;
using System.IO;

namespace OutlookParser
{
  public static class MailMessageExtensions
  {
    public static MailMessage GenerateReply(this MailMessage msg, MailAddress currentAddress, Action<TextReader, TextWriter, string> transform)
    {
      var toList = (from a in msg.To
                    where !a.Equals(currentAddress)
                    select a.ToString());

      var result = new MailMessage(currentAddress.ToString(), (msg.ReplyTo == null ? msg.From.ToString() : msg.ReplyTo.ToString()));
      result.MessageId("<" + Guid.NewGuid().ToString("N").ToUpperInvariant() + "@ct.gentex.com>");

      AlternateView newView;
      if (!string.IsNullOrEmpty(msg.Body))
      {
        var contentType = (msg.IsBodyHtml ? MediaTypeNames.Text.Html : MediaTypeNames.Text.Plain);
        if (transform == null)
        {
          newView = AlternateView.CreateAlternateViewFromString(msg.Body, new ContentType(contentType));
        } 
        else 
        {
          var stream = new MemoryStream(msg.Body.Length);
          using (var reader = new StringReader(msg.Body))
          {
            using (var writer = new StreamWriter(stream))
            {
              transform.Invoke(reader, writer, contentType);
            }
          }
          newView = new AlternateView(stream, new ContentType(contentType));
        }
        
        newView.TransferEncoding = System.Net.Mime.TransferEncoding.SevenBit;
        msg.AlternateViews.Add(newView);
      }

      foreach (var view in msg.AlternateViews)
      {
        view.ContentStream.Position = 0;
        if (transform == null)
        {
          newView = new AlternateView(view.ContentStream, view.ContentType);
        }
        else 
        {
          var stream = new MemoryStream();
          using (var reader = new StreamReader(view.ContentStream))
          {
            using (var writer = new StreamWriter(stream))
            {
              transform.Invoke(reader, writer, view.ContentType.MediaType);
            }
          }
          newView = new AlternateView(stream, view.ContentType);
        }
        newView.BaseUri = view.BaseUri;
        newView.ContentId = view.ContentId;
        newView.TransferEncoding = view.TransferEncoding;
        result.AlternateViews.Add(newView);
      }

      foreach (var addr in msg.CC)
      {
        result.CC.Add(addr);
      }
      result.Priority = msg.Priority;
      if (msg.Subject.StartsWith("RE: ", StringComparison.InvariantCultureIgnoreCase))
      {
        result.Subject = msg.Subject;
      }
      else
      {
        result.Subject = "RE: " + msg.Subject;
      }
      result.Headers.Set("In-Reply-To", msg.MessageId());
      var references = result.Headers.Get("References");
      result.Headers.Set("References", references + (string.IsNullOrEmpty(references) ? "" : Environment.NewLine) + msg.MessageId());

      return result;
    }

    public static string MessageId(this MailMessage msg)
    {
      var values = msg.Headers.GetValues("Message-ID:");
      if (values == null || values.Length < 1) values = msg.Headers.GetValues("Message-ID");
      if (values == null || values.Length < 1) return string.Empty;
      return values[0];
    }
    public static void MessageId(this MailMessage msg, string value)
    {
      msg.Headers["Message-ID"] = value;
    }
  }
}
