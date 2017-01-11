using OpenMcdf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OutlookParser
{
  public class OutlookRecipient : OutlookStorage
  {
    #region Property(s)

    /// <summary>
    /// Gets the display name.
    /// </summary>
    /// <value>The display name.</value>
    public string DisplayName
    {
      get { return this.GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME); }
    }

    /// <summary>
    /// Gets the recipient email.
    /// </summary>
    /// <value>The recipient email.</value>
    public string Email
    {
      get
      {
        string email = this.GetMapiPropertyString(MapiTags.PR_SMTP_ADDRESS);
        // try EMAIL_2 if EMAIL is blank
        if (String.IsNullOrEmpty(email)) email = this.GetMapiPropertyString(MapiTags.PR_ORGEMAILADDR);
        if (String.IsNullOrEmpty(email)) email = this.GetMapiPropertyString(MapiTags.PR_EMAIL_ADDRESS);
        // try DISPLAY_NAME if EMAIL is still blank, and DISPLAY_NAME is a valid E-mail address
        if (String.IsNullOrEmpty(email) && IsValidEmail(this.GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME)))
        {
          email = this.GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME);
        }
        return email;
      }
    }

    /// <summary>
    /// Gets the recipient type.
    /// </summary>
    /// <value>The recipient type.</value>
    public RecipientType Type
    {
      get
      {
        return (RecipientType)this.GetMapiPropertyInt32(MapiTags.PR_RECIPIENT_TYPE);
      }
    }

    #endregion

    #region Constructor(s)

    /// <summary>
    /// Initializes a new instance of the <see cref="OutlookRecipient"/> class.
    /// </summary>
    /// <param name="message">The message.</param>
    public OutlookRecipient(OutlookStorage parent, CFStorage storage)
      : base(parent, storage)
    {
      this._propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;
    }

    #endregion

    bool invalid = false;
    public bool IsValidEmail(string strIn)
    {
      invalid = false;
      if (String.IsNullOrEmpty(strIn))
        return false;

      // Use IdnMapping class to convert Unicode domain names.
      strIn = Regex.Replace(strIn, @"(@)(.+)$", this.DomainMapper);
      if (invalid)
        return false;

      // Return true if strIn is in valid e-mail format. 
      return Regex.IsMatch(strIn,
             @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
             @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
             RegexOptions.IgnoreCase);
    }

    private string DomainMapper(Match match)
    {
      // IdnMapping class with default property values.
      var idn = new IdnMapping();

      string domainName = match.Groups[2].Value;
      try
      {
        domainName = idn.GetAscii(domainName);
      }
      catch (ArgumentException)
      {
        invalid = true;
      }
      return match.Groups[1].Value + domainName;
    }
  }
}
