using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  public class OutlookMessage : OutlookStorage
  {
    private OutlookNameProp namePropStorage;

    #region Property(s)

    /// <summary>
    /// Gets the list of recipients in the outlook message.
    /// </summary>
    /// <value>The list of recipients in the outlook message.</value>
    public IEnumerable<OutlookRecipient> Recipients
    {
      get { return this.recipients; }
    }
    private List<OutlookRecipient> recipients = new List<OutlookRecipient>();

    /// <summary>
    /// Gets the list of attachments in the outlook message.
    /// </summary>
    /// <value>The list of attachments in the outlook message.</value>
    public List<OutlookAttachment> Attachments
    {
      get { return this.attachments; }
    }
    private List<OutlookAttachment> attachments = new List<OutlookAttachment>();

    public String MessageId
    {
      get { return this.GetMapiPropertyString(MapiTags.PR_INTERNET_MESSAGE_ID); }
    }
    public string Headers
    {
      get { return this.GetMapiPropertyString(MapiTags.PR_TRANSPORT_MESSAGE_HEADERS); }
    }

    /// <summary>
    /// Gets the list of sub messages in the outlook message.
    /// </summary>
    /// <value>The list of sub messages in the outlook message.</value>
    public List<OutlookMessage> Messages
    {
      get { return this.messages; }
    }
    private List<OutlookMessage> messages = new List<OutlookMessage>();

    /// <summary>
    /// Gets the display value of the contact that sent the email.
    /// </summary>
    /// <value>The display value of the contact that sent the email.</value>
    public string From
    {
      get { return this.GetMapiPropertyString(MapiTags.PR_SENDER_NAME); }
    }

    /// <summary>
    /// Gets the display value of the contact that sent the email.
    /// </summary>
    /// <value>The display value of the contact that sent the email.</value>
    public string FromEmail
    {
      get
      {
        var result = this.GetMapiPropertyString(MapiTags.PR_SENDER_EMAIL_ADDRESS);
        if (string.IsNullOrEmpty(result)) result = this.GetMapiPropertyString(MapiTags.PR_PRIMARY_SEND_ACCT);
        if (string.IsNullOrEmpty(result)) result = this.GetMapiPropertyString(MapiTags.PR_NEXT_SEND_ACCT);
        return result;
      }
    }

    public DateTime? ReceivedTime
    {
      get
      {
        return (DateTime?)this.GetMapiProperty(MapiTags.PR_MESSAGE_DELIVERY_TIME);
      }
    }

    public string AppointmentLocation
    {
      get
      {
        return this.GetMapiPropertyString(GetNamedPropId(MapiTags.Location));
      }
    }
    public DateTime AppointmentStart
    {
      get
      {
        return (DateTime)this.GetMapiProperty(GetNamedPropId(MapiTags.AppointmentStartWhole));
      }
    }
    public DateTime AppointmentEnd
    {
      get
      {
        return (DateTime)this.GetMapiProperty(GetNamedPropId(MapiTags.AppointmentEndWhole));
      }
    }
    public long AppointmentDuration
    {
      get
      {
        return (long)this.GetMapiProperty(GetNamedPropId(MapiTags.Duration));
      }
    }

    private string GetNamedPropId(string numericName)
    {
      var namePropRef = namePropStorage;
      var parent = this.parentMessage as OutlookMessage;
      while (namePropRef == null && parent != null)
      {
        namePropRef = parent.namePropStorage;
        parent = parent.parentMessage as OutlookMessage;
      }
      if (namePropRef == null)
      {
        return null;
      }
      else
      {
        return namePropRef.PropIdFromName(numericName);
      }
    }

    public DateTime SentTime
    {
      get
      {
        return (DateTime)this.GetMapiProperty(MapiTags.PR_CLIENT_SUBMIT_TIME);
      }
    }

    /// <summary>
    /// Gets the subject of the outlook message.
    /// </summary>
    /// <value>The subject of the outlook message.</value>
    public String Subject
    {
      get { return this.GetMapiPropertyString(MapiTags.PR_SUBJECT); }
    }

    /// <summary>
    /// Gets the body of the outlook message in plain text format.
    /// </summary>
    /// <value>The body of the outlook message in plain text format.</value>
    public String BodyText
    {
      get { return this.GetMapiPropertyString(MapiTags.PR_BODY); }
    }

    /// <summary>
    /// Gets the body of the outlook message in html.
    /// </summary>
    /// <value>The body of the outlook message in html.</value>
    public String BodyHtml
    {
      get { return this.GetMapiPropertyString(MapiTags.PR_BODY_HTML); }
    }

    public String Class
    {
      get { return this.GetMapiPropertyString(MapiTags.PR_MESSAGE_CLASS); }
    }

    /// <summary>
    /// Gets the body of the outlook message in RTF format.
    /// </summary>
    /// <value>The body of the outlook message in RTF format.</value>
    public String BodyRtf
    {
      get
      {
        //get value for the RTF compressed MAPI property
        byte[] rtfBytes = this.GetMapiPropertyBytes(MapiTags.PR_RTF_COMPRESSED);

        //return null if no property value exists
        if (rtfBytes == null || rtfBytes.Length == 0)
        {
          return null;
        }

        //decompress the rtf value
        rtfBytes = CLZF.decompressRTF(rtfBytes);

        //encode the rtf value as an ascii string and return
        return Encoding.ASCII.GetString(rtfBytes);
      }
    }

    #endregion

    #region Constructor(s)

    /// <summary>
    /// Initializes a new instance of the <see cref="OutlookMessage"/> class from a msg file.
    /// </summary>
    /// <param name="filename">The msg file to load.</param>
    public OutlookMessage(string msgfile) : base(msgfile) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="OutlookMessage"/> class from a <see cref="Stream"/> containing an IStorage.
    /// </summary>
    /// <param name="storageStream">The <see cref="Stream"/> containing an IStorage.</param>
    public OutlookMessage(Stream storageStream) : base(storageStream) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="OutlookMessage"/> class on the specified <see cref="NativeMethods.IStorage"/>.
    /// </summary>
    /// <param name="storage">The storage to create the <see cref="OutlookMessage"/> on.</param>
    private OutlookMessage(CompoundFile file, CFStorage storage)
      : base(file, storage)
    {
      this._propHeaderSize = MapiTags.PropertiesStreamHeaderTop;
    }

    #endregion

    #region Methods(LoadStorage)

    /// <summary>
    /// Processes sub storages on the specified storage to capture attachment and recipient data.
    /// </summary>
    /// <param name="storage">The storage to check for attachment and recipient data.</param>
    protected override void LoadStorage(CFStorage storage)
    {
      base.LoadStorage(storage);

      _storage.VisitEntries(i =>
      {
        var st = i as CFStorage;
        if (st != null)
        {
          if (i.Name.StartsWith(MapiTags.RecipStoragePrefix))
          {
            OutlookRecipient recipient = new OutlookRecipient(this, st);
            this.recipients.Add(recipient);
          }
          else if (i.Name.StartsWith(MapiTags.AttachStoragePrefix))
          {
            this.LoadAttachmentStorage(st);
          }
          else if (i.Name.StartsWith(MapiTags.NameIdStorage))
          {
            namePropStorage = new OutlookNameProp(this, st);
          }
        }
      }, false);
    }

    /// <summary>
    /// Loads the attachment data out of the specified storage.
    /// </summary>
    /// <param name="storage">The attachment storage.</param>
    private void LoadAttachmentStorage(CFStorage storage)
    {
      //create attachment from attachment storage
      var attachment = new OutlookAttachment(this, storage);

      //if attachment is an embedded msg handle differently than a normal attachment
      var attachMethod = attachment.GetMapiPropertyInt32(MapiTags.PR_ATTACH_METHOD) ?? 0;
      if (attachMethod == MapiTags.ATTACH_EMBEDDED_MSG)
      {
        //create new Message and set parent and header size
        OutlookMessage subMsg = new OutlookMessage(_file, attachment.GetMapiProperty(MapiTags.PR_ATTACH_DATA_BIN) as CFStorage);
        subMsg.parentMessage = this;
        subMsg._propHeaderSize = MapiTags.PropertiesStreamHeaderEmbeded;

        //add to messages list
        this.messages.Add(subMsg);
      }
      else
      {
        //add attachment to attachment list
        this.attachments.Add(attachment);
      }
    }

    #endregion

    #region Methods(Disposing)

    protected override void Disposing()
    {

      //dispose sub storages
      foreach (OutlookStorage subMsg in this.messages)
      {
        subMsg.Dispose();
      }

      //dispose sub storages
      foreach (OutlookStorage recip in this.recipients)
      {
        recip.Dispose();
      }

      //dispose sub storages
      foreach (OutlookStorage attach in this.attachments)
      {
        attach.Dispose();
      }
    }

    #endregion
  }
}
