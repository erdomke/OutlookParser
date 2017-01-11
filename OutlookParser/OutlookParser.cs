using MimeKit;
using MimeKit.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  public class OutlookParser
  {
    private OutlookStorage.Message _msg;

    public OutlookParser(Stream stream)
    {
      _msg = new OutlookStorage.Message(stream);
    }
    
    public Email ParseMessage()
    {
      //var result = new Email();
      HeaderList headers;
      using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(_msg.Headers)))
      {
        headers = HeaderList.Load(stream);
      }

      var bodyParts = new List<MimeEntity>();
      if (!string.IsNullOrEmpty(_msg.BodyText))
        bodyParts.Add(new TextPart("plain")
        {
          Text = _msg.BodyText
        });
      if (!string.IsNullOrEmpty(_msg.BodyRtf))
        bodyParts.Add(new TextPart("rtf")
        {
          Text = _msg.BodyRtf
        });
      if (!string.IsNullOrEmpty(_msg.BodyHtml))
        bodyParts.Add(new TextPart("html")
        {
          Text = _msg.BodyHtml
        });

      MimeEntity root = null;
      if (bodyParts.Count <= 0)
        throw new InvalidOperationException("No message body found.");
      if (bodyParts.Count == 1)
      {
        root = bodyParts[0];
      }
      else
      {
        var alt = new Multipart("alternative");
        foreach (var body in bodyParts)
        {
          alt.Add(body);
        }
        root = alt;
      }


      if (_msg.Attachments.Any() || _msg.Messages.Any())
      {
        var mixed = new Multipart("mixed");
        mixed.Add(root);

        foreach (var attach in _msg.Attachments)
        {
          var mimeType = (string.IsNullOrEmpty(attach.MimeTag) ? "application/octet-stream" : attach.MimeTag).Split('/');
          var mimeAttach = new MimePart(mimeType[0], mimeType[1])
          {
            ContentObject = new ContentObject(new MemoryStream(attach.Data)),
            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
            ContentTransferEncoding = ContentEncoding.Base64,
            ContentId = attach.ContentId,
            FileName = attach.Filename,
          };
          mixed.Add(mimeAttach);
        }
        
        foreach (var msg in _msg.Messages)
        {
          switch (msg.Type)
          {
            case MessageType.Email:
            case MessageType.EmailSms:
            case MessageType.EmailNonDeliveryReport:
            case MessageType.EmailDeliveryReport:
            case MessageType.EmailDelayedDeliveryReport:
            case MessageType.EmailReadReceipt:
            case MessageType.EmailNonReadReceipt:
            case MessageType.EmailEncryptedAndMaybeSigned:
            case MessageType.EmailEncryptedAndMaybeSignedNonDelivery:
            case MessageType.EmailEncryptedAndMaybeSignedDelivery:
            case MessageType.EmailClearSignedReadReceipt:
            case MessageType.EmailClearSignedNonDelivery:
            case MessageType.EmailClearSignedDelivery:
            case MessageType.EmailBmaStub:
            case MessageType.CiscoUnityVoiceMessage:
            case MessageType.EmailClearSigned:

          }
          var mimeAttach = new MimePart("application", "octet-stream")
          {
            ContentObject = new ContentObject(new MemoryStream(attach.Data)),
            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
            ContentTransferEncoding = ContentEncoding.Base64,
            ContentId = attach.ContentId,
            FileName = attach.Filename,
          };
          mixed.Add(mimeAttach);
        }


        root = mixed;
      }

      var result = new MimeMessage(headers, root);

      //result.LoadHeaders(headers);

      return null;
    }
  }
}
