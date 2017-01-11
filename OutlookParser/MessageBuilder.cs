using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookParser
{
  /// <summary>
  /// A helper class for reading mail message data and building a MailMessage instance out of it.
  /// </summary>
  public static class MessageBuilder
  {
    /// <summary>
    /// Creates a new empty instance of the MailMessage class from a string containing a raw mail
    /// message header.
    /// </summary>
    /// <param name="text">The mail header to create the MailMessage instance from.</param>
    /// <returns>A MailMessage instance with initialized Header fields but without any
    /// content.</returns>
    public static MailMessage FromHeader(string text)
    {
      //NameValueCollection header = ParseMailHeader(text);
      MailMessage m = new MailMessage();
      ParseFillMailHeader(text, m.Headers);
      Match ma = Regex.Match(m.Headers["Subject"] ?? "", @"=\?([A-Za-z0-9\-_]+)");
      if (ma.Success)
      {
        // encoded-word subject. A subject must not contain any encoded newline
        // characters, so if we find any, we strip them off.
        m.SubjectEncoding = Util.GetEncoding(ma.Groups[1].Value);
        try
        {
          m.Subject = Util.DecodeWords(m.Headers["Subject"]).
            Replace("\n", "").Replace("\r", "");
        }
        catch
        {
          // If, for any reason, decoding fails, set the subject to the
          // original, unaltered string.
          m.Subject = m.Headers["Subject"];
        }
      }
      else
      {
        m.SubjectEncoding = Encoding.ASCII;
        m.Subject = m.Headers["Subject"];
      }
      m.Priority = ParsePriority(m.Headers["Priority"]);
      SetAddressFields(m, m.Headers);
      return m;
    }

    internal static void ParseFillMailHeader(string header, NameValueCollection coll)
    {
      if (string.IsNullOrEmpty(header)) return;

      using (var reader = new StringReader(header))
      {
        string line;
        string fieldname = null;
        string fieldvalue = null;

        var exclude = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase) {
          "Subject", "Comments", "Content-disposition", "User-Agent" };

        while ((line = reader.ReadLine()) != null)
        {
          if (line == String.Empty) continue;

          // Values may stretch over several lines.
          if (line[0] == ' ' || line[0] == '\t')
          {
            if (fieldname != null) fieldvalue += line.TrimEnd();
          }
          else
          {
            if (fieldname != null)
            {
              // Strip comments from RFC822 and MIME fields unless they are unstructured fields.
              if (!exclude.Contains(fieldname)) fieldvalue = StripComments(fieldvalue);

              try
              {
                if (!string.IsNullOrEmpty(fieldvalue)) coll.Add(fieldname, fieldvalue);
              }
              catch
              {
                // HeaderCollection throws an exception if adding an empty string as value, which can
                // happen, if reading a mail message with an empty subject.
                // Also spammers often forge headers, so just fall through and ignore.
              }
            }

            // The mail header consists of field:value pairs.
            int delimiter = line.IndexOf(':');
            if (delimiter < 0) continue;
            fieldname = line.Substring(0, delimiter).Trim();
            fieldvalue = line.Substring(delimiter + 1).Trim();
          }
        }

        if (fieldname != null)
        {
          // Strip comments from RFC822 and MIME fields unless they are unstructured fields.
          if (!exclude.Contains(fieldname)) fieldvalue = StripComments(fieldvalue);

          try
          {
            coll.Add(fieldname, Util.DecodeWords(fieldvalue));
          }
          catch
          {
            // HeaderCollection throws an exception if adding an empty string as value, which can
            // happen, if reading a mail message with an empty subject.
            // Also spammers often forge headers, so just fall through and ignore.
          }
        }
      }
    }

    /// <summary>
    /// Parses the mail header of a mail message and returns it as a NameValueCollection.
    /// </summary>
    /// <param name="header">The mail header to parse.</param>
    /// <returns>A NameValueCollection containing the header fields as keys with their respective
    /// values as values.</returns>
    internal static NameValueCollection ParseMailHeader(string header)
    {
      NameValueCollection coll = new NameValueCollection();
      ParseFillMailHeader(header, coll);
      return coll;
    }

    /// <summary>
    /// Strips RFC822/MIME comments from the specified string.
    /// </summary>
    /// <param name="s">The string to strip comments from.</param>
    /// <returns>A new string stripped of any comments.</returns>
    internal static string StripComments(string s)
    {
      if (String.IsNullOrEmpty(s))
        return s;
      bool inQuotes = false, escape = false;
      char last = ' ';
      int depth = 0;
      StringBuilder builder = new StringBuilder();
      for (int i = 0; i < s.Length; i++)
      {
        char c = s[i];
        if (c == '\\' && !escape)
        {
          escape = true;
          continue;
        }
        if (c == '"' && !escape)
          inQuotes = !inQuotes;
        last = c;
        if (!inQuotes && !escape && c == '(')
          depth++;
        else if (!inQuotes && !escape && c == ')' && depth > 0)
          depth--;
        else if (depth <= 0)
          builder.Append(c);
        escape = false;
      }
      return builder.ToString().Trim();
    }

    /// <summary>
    /// Parses a MIME header field which can contain multiple 'parameter = value'
    /// pairs (such as Content-Type: text/html; charset=iso-8859-1).
    /// </summary>
    /// <param name="field">The header field to parse.</param>
    /// <returns>A NameValueCollection containing the parameter names as keys with the respective
    /// parameter values as values.</returns>
    /// <remarks>The value of the actual field disregarding the 'parameter = value' pairs is stored
    /// in the collection under the key "value" (in the above example of Content-Type, this would
    /// be "text/html").</remarks>
    static NameValueCollection ParseMIMEField(string field)
    {
      NameValueCollection coll = new NameValueCollection();
      var fixup = new HashSet<string>();
      try
      {
        // This accounts for MIME Parameter Value Extensions (RFC2231).
        MatchCollection matches = Regex.Matches(field,
          @"([\w\-]+)(?:\*\d{1,3})?(\*?)?\s*=\s*([^;]+)");
        foreach (Match m in matches)
        {
          string pname = m.Groups[1].Value.Trim(), pval = m.Groups[3].Value.Trim('"');
          coll[pname] = coll[pname] + pval;
          if (m.Groups[2].Value == "*")
            fixup.Add(pname);
        }
        foreach (var pname in fixup)
        {
          try
          {
            coll[pname] = Util.Rfc2231Decode(coll[pname]);
          }
          catch (FormatException)
          {
            // If decoding fails, we should at least return the un-altered value.
          }
        }
        Match mvalue = Regex.Match(field, @"^\s*([^;]+)");
        coll.Add("value", mvalue.Success ? mvalue.Groups[1].Value.Trim() : "");
      }
      catch
      {
        // We don't want this to blow up on the user with weird mails.
        coll.Add("value", String.Empty);
      }
      return coll;
    }

    /// <summary>
    /// Parses a mail header address-list field such as To, Cc and Bcc which can contain multiple
    /// email addresses.
    /// </summary>
    /// <param name="list">The address-list field to parse</param>
    /// <returns>An array of MailAddress objects representing the parsed mail addresses.</returns>
    internal static MailAddress[] ParseAddressList(string list)
    {
      List<MailAddress> mails = new List<MailAddress>();
      if (String.IsNullOrEmpty(list))
        return mails.ToArray();
      foreach (string part in SplitAddressList(list))
      {
        MailAddressCollection mcol = new MailAddressCollection();
        try
        {
          // .NET won't accept address-lists ending with a ';' or a ',' character, see #68.
          mcol.Add(part.TrimEnd(';', ','));
          foreach (MailAddress m in mcol)
          {
            // We might still need to decode the display name if it is Q-encoded.
            string displayName = Util.DecodeWords(m.DisplayName);
            mails.Add(new MailAddress(m.Address, displayName));
          }
        }
        catch
        {
          // We don't want this to throw any exceptions even if the entry is malformed.
        }
      }
      return mails.ToArray();
    }

    /// <summary>
    /// Splits the specified address-list into individual parts consisting of a mail address and
    /// optionally a display-name.
    /// </summary>
    /// <param name="list">The address-list to split into parts.</param>
    /// <returns>An enumerable collection of parts.</returns>
    internal static IEnumerable<string> SplitAddressList(string list)
    {
      IList<string> parts = new List<string>();
      StringBuilder builder = new StringBuilder();
      bool inQuotes = false;
      char last = '.';
      for (int i = 0; i < list.Length; i++)
      {
        if (list[i] == '"' && last != '\\')
          inQuotes = !inQuotes;
        if (list[i] == ',' && !inQuotes)
        {
          parts.Add(builder.ToString().Trim());
          builder.Length = 0;
        }
        else
        {
          builder.Append(list[i]);
        }
        if (i == list.Length - 1)
          parts.Add(builder.ToString().Trim());
      }
      return parts;
    }

    /// <summary>
    /// Parses the priority of a mail message which can be specified as part of the header
    /// information.
    /// </summary>
    /// <param name="priority">The mail header priority value. The value can be null in which case
    /// a "normal priority" is returned.</param>
    /// <returns>A value from the MailPriority enumeration corresponding to the specified mail
    /// priority. If the passed priority value is null or invalid, a normal priority is assumed and
    /// MailPriority.Normal is returned.</returns>
    static MailPriority ParsePriority(string priority)
    {
      Dictionary<string, MailPriority> Map =
        new Dictionary<string, MailPriority>(StringComparer.OrdinalIgnoreCase) {
            { "non-urgent", MailPriority.Low },
            { "normal",	MailPriority.Normal },
            { "urgent",	MailPriority.High }
        };
      try
      {
        return Map[priority];
      }
      catch
      {
        return MailPriority.Normal;
      }
    }

    /// <summary>
    /// Sets the address fields (From, To, CC, etc.) of a MailMessage object using the specified
    /// mail message header information.
    /// </summary>
    /// <param name="m">The MailMessage instance to operate on.</param>
    /// <param name="header">A collection of mail and MIME headers.</param>
    static void SetAddressFields(MailMessage m, NameValueCollection header)
    {
      MailAddress[] addr;
      if (header["To"] != null)
      {
        addr = ParseAddressList(header["To"]);
        foreach (MailAddress a in addr)
          m.To.Add(a);
      }
      if (header["Cc"] != null)
      {
        addr = ParseAddressList(header["Cc"]);
        foreach (MailAddress a in addr)
          m.CC.Add(a);
      }
      if (header["Bcc"] != null)
      {
        addr = ParseAddressList(header["Bcc"]);
        foreach (MailAddress a in addr)
          m.Bcc.Add(a);
      }
      if (header["From"] != null)
      {
        addr = ParseAddressList(header["From"]);
        if (addr.Length > 0)
          m.From = addr[0];
      }
      if (header["Sender"] != null)
      {
        addr = ParseAddressList(header["Sender"]);
        if (addr.Length > 0)
          m.Sender = addr[0];
      }
      if (header["Reply-to"] != null)
      {
        addr = ParseAddressList(header["Reply-to"]);
        if (addr.Length > 0) m.ReplyTo = addr[0];
      }
    }
  }
}
