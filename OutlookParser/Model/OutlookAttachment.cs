using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  public class OutlookAttachment : OutlookStorage
  {
    #region Fields
    /// <summary>
    /// Containts the data of the attachment as an byte array
    /// </summary>
    private byte[] _data;
    #endregion

    #region Properties
    /// <summary>
    /// Returns the filename of the attachment
    /// </summary>
    public string FileName { get; private set; }

    /// <summary>
    /// Retuns the data
    /// </summary>
    public byte[] Data
    {
      get { return _data ?? GetMapiPropertyBytes(MapiTags.PR_ATTACH_DATA_BIN); }
    }

    /// <summary>
    /// Returns the content id or null when not available
    /// </summary>
    public string ContentId { get; private set; }

    /// <summary>
    /// Returns the rendering position or -1 when unknown
    /// </summary>
    public int RenderingPosition { get; private set; }

    /// <summary>
    /// True when the attachment is inline
    /// </summary>
    public bool IsInline { get; private set; }

    /// <summary>
    /// True when the attachment is a contact photo. This can only be true
    /// when the <see cref="Storage.Message"/> object is an 
    /// <see cref="Storage.Message.MessageType.Contact"/> object.
    /// </summary>
    public bool IsContactPhoto { get; private set; }

    /// <summary>
    /// Returns the date and time when the attachment was created or null
    /// when not available
    /// </summary>
    public DateTime? CreationTime { get; private set; }

    /// <summary>
    /// Returns the date and time when the attachment was last modified or null
    /// when not available
    /// </summary>
    public DateTime? LastModificationTime { get; private set; }

    /// <summary>
    /// Returns the Mime Type tag
    /// </summary>
    public string MimeTag { get; private set; }

    /// <summary>
    /// Returns <c>true</c> when the attachment is an OLE attachment
    /// </summary>
    public bool OleAttachment { get; private set; }
    #endregion

    #region Constructor(s)

    /// <summary>
    /// Initializes a new instance of the <see cref="OutlookAttachment"/> class.
    /// </summary>
    /// <param name="message">The message.</param>
    internal OutlookAttachment(OutlookStorage parent, CFStorage message)
      : base(parent, message)
    {
      GC.SuppressFinalize(message);
      _propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;

      CreationTime = GetMapiPropertyDateTime(MapiTags.PR_CREATION_TIME);
      LastModificationTime = GetMapiPropertyDateTime(MapiTags.PR_LAST_MODIFICATION_TIME);

      ContentId = GetMapiPropertyString(MapiTags.PR_ATTACH_CONTENTID);
      IsInline = ContentId != null;

      var isContactPhoto = GetMapiPropertyBool(MapiTags.PR_ATTACHMENT_CONTACTPHOTO);
      if (isContactPhoto == null)
        IsContactPhoto = false;
      else
        IsContactPhoto = (bool)isContactPhoto;

      var renderingPosition = GetMapiPropertyInt32(MapiTags.PR_RENDERING_POSITION);
      if (renderingPosition == null)
        RenderingPosition = -1;
      else
        RenderingPosition = (int)renderingPosition;

      var fileName = GetMapiPropertyString(MapiTags.PR_ATTACH_LONG_FILENAME);

      if (string.IsNullOrEmpty(fileName))
        fileName = GetMapiPropertyString(MapiTags.PR_ATTACH_FILENAME);

      if (string.IsNullOrEmpty(fileName))
        fileName = GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME);

      FileName = fileName != null
          ? FileManager.RemoveInvalidFileNameChars(fileName)
          : "Nameless";

      MimeTag = this.GetMapiPropertyString(MapiTags.PR_ATTACH_MIME_TAG);

      var attachmentMethod = GetMapiPropertyInt32(MapiTags.PR_ATTACH_METHOD);
      switch (attachmentMethod)
      {
        case MapiTags.ATTACH_BY_REFERENCE:
        case MapiTags.ATTACH_BY_REF_RESOLVE:
        case MapiTags.ATTACH_BY_REF_ONLY:
          ResolveAttachment();
          break;

        case MapiTags.ATTACH_OLE:
          var storage = GetMapiProperty(MapiTags.PR_ATTACH_DATA_BIN) as CFStorage;
          //var attachmentOle = new OutlookAttachment(this, );
          _data = storage.GetStream("CONTENTS").GetData();
          var fileTypeInfo = FileTypeSelector.GetFileTypeFileInfo(Data);

          if (string.IsNullOrEmpty(FileName))
            FileName = fileTypeInfo.Description;

          FileName += "." + fileTypeInfo.Extension.ToLower();
          IsInline = true;
          break;
      }
    }

    #endregion

    #region ResolveAttachment
    /// <summary>
    /// Tries to resolve an attachment when the <see cref="MapiTags.PR_ATTACH_METHOD"/> is of the type
    /// <see cref="MapiTags.ATTACH_BY_REFERENCE"/>, <see cref="MapiTags.ATTACH_BY_REF_RESOLVE"/> or
    /// <see cref="MapiTags.ATTACH_BY_REF_ONLY"/>
    /// </summary>
    private void ResolveAttachment()
    {
      //The PR_ATTACH_PATHNAME or PR_ATTACH_LONG_PATHNAME property contains a fully qualified path identifying the attachment
      var attachPathName = GetMapiPropertyString(MapiTags.PR_ATTACH_PATHNAME);
      var attachLongPathName = GetMapiPropertyString(MapiTags.PR_ATTACH_LONG_PATHNAME);

      // Because we are not sure we can access the files we put everything in a try catch
      try
      {
        if (attachLongPathName != null)
        {
          _data = File.ReadAllBytes(attachLongPathName);
          return;
        }

        if (attachPathName == null) return;
        _data = File.ReadAllBytes(attachPathName);
      }
      // ReSharper disable once EmptyGeneralCatchClause
      catch { }
    }
    #endregion
  }
}
