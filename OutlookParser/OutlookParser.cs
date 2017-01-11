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
    private OutlookMessage _msg;

    public OutlookParser(Stream stream)
    {
      _msg = new OutlookMessage(stream);
    }
    
    public MimeMessage ParseMessage()
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
            FileName = attach.FileName,
          };
          mixed.Add(mimeAttach);
        }
        
        foreach (var message in _msg.Messages)
        {
          var stream = new MemoryStream();
          message.WriteTo(stream);
          stream.Position = 0;

          var mimeAttach = new MimePart("application", "vnd.ms-outlook")
          {
            ContentObject = new ContentObject(stream),
            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
            ContentTransferEncoding = ContentEncoding.Base64,
            ContentId = message.MessageId,
            FileName = (message.Subject ?? "") + ".msg"
          };
          mixed.Add(mimeAttach);
        }

        root = mixed;
      }

      return new MimeMessage(headers, root);
    }
  }
}
