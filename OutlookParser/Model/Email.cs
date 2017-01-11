using MimeKit;
using MimeKit.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  public class Email
  {
    private Dictionary<string, InternetAddress[]> _addresses 
      = new Dictionary<string, InternetAddress[]>(StringComparer.OrdinalIgnoreCase);
    private DateTimeOffset _date;
    private Importance _importance;
    private List<KeyValuePair<string, string>> _headers;
    private string _inReplyTo;
    private Version _mimeVersion;
    private Priority _priority;
    private List<string> _references;
    private string _resentMessageId;
    private string _messageId;
    private MailboxAddress _resentSender;
    private MailboxAddress _sender;
    private string _subject;
    private DateTimeOffset _resentDate;
    private XPriority _xpriority;

    public IEnumerable<InternetAddress> Bcc { get { return GetAddresses("Bcc"); } }
    public IEnumerable<InternetAddress> Cc { get { return GetAddresses("Cc"); } }
    public DateTimeOffset Date { get { return _date; } }
    public IEnumerable<InternetAddress> From { get { return GetAddresses("From"); } }
    public IEnumerable<KeyValuePair<string, string>> Headers { get { return _headers; } }
    public Importance Importance { get { return _importance; } }
    public string InReplyTo { get { return _inReplyTo; } }
    public Version MimeVersion { get { return _mimeVersion; } }
    public Priority Priority { get { return _priority; } }
    public IEnumerable<string> References { get { return _references; } }
    public IEnumerable<InternetAddress> ReplyTo { get { return GetAddresses("Reply-To"); } }
    public IEnumerable<InternetAddress> ResentBcc { get { return GetAddresses("Resent-Bcc"); } }
    public IEnumerable<InternetAddress> ResentCc { get { return GetAddresses("Resent-Cc"); } }
    public DateTimeOffset ResentDate { get { return _resentDate; } }
    public IEnumerable<InternetAddress> ResentFrom { get { return GetAddresses("Resent-From"); } }
    public string ResentMessageId { get { return _resentMessageId; } }
    public IEnumerable<InternetAddress> ResentReplyTo { get { return GetAddresses("Resent-Reply-To"); } }
    public MailboxAddress ResentSender { get { return _resentSender; } }
    public IEnumerable<InternetAddress> ResentTo { get { return GetAddresses("Resent-To"); } }
    public MailboxAddress Sender { get { return _sender; } }
    public string Subject { get { return _subject; } }
    public IEnumerable<InternetAddress> To { get { return GetAddresses("To"); } }
    public XPriority XPriority { get { return _xpriority; } }


    internal void LoadHeaders(HeaderList headers)
    {
      _headers = new List<KeyValuePair<string, string>>();
      var options = new ParserOptions();
      MimeKit.MailboxAddress address;
      MimeKit.InternetAddressList addresses;
      foreach (var header in headers)
      {
        int index = 0;
        int number = 0;

        _headers.Add(new KeyValuePair<string, string>(header.Field, header.Value));
        var rawValue = header.RawValue;
        switch (header.Id)
        {
          case HeaderId.MimeVersion:
            MimeUtils.TryParse(rawValue, 0, rawValue.Length, out _mimeVersion);
            break;
          case HeaderId.References:
            _references = new List<string>();
            foreach (var msgId in MimeUtils.EnumerateReferences(rawValue, 0, rawValue.Length))
            {
              _references.Add(msgId);
            }
            break;
          case HeaderId.InReplyTo:
            _inReplyTo = MimeUtils.EnumerateReferences(rawValue, 0, rawValue.Length).FirstOrDefault();
            break;
          case HeaderId.ResentMessageId:
            _resentMessageId = MimeUtils.ParseMessageId(rawValue, 0, rawValue.Length);
            break;
          case HeaderId.MessageId:
            _messageId = MimeUtils.ParseMessageId(rawValue, 0, rawValue.Length);
            break;
          case HeaderId.ResentSender:
            if (MimeKit.MailboxAddress.TryParse(options, rawValue, 0, rawValue.Length, out address))
              _resentSender = new MailboxAddress(address);
            break;
          case HeaderId.Sender:
            if (MimeKit.MailboxAddress.TryParse(options, rawValue, 0, rawValue.Length, out address))
              _sender = new MailboxAddress(address);
            break;
          case HeaderId.ResentDate:
            DateUtils.TryParse(rawValue, 0, rawValue.Length, out _resentDate);
            break;
          case HeaderId.Importance:
            switch (header.Value.ToLowerInvariant().Trim())
            {
              case "high": _importance = Importance.High; break;
              case "low": _importance = Importance.Low; break;
              default: _importance = Importance.Normal; break;
            }
            break;
          case HeaderId.Priority:
            switch (header.Value.ToLowerInvariant().Trim())
            {
              case "non-urgent": _priority = Priority.NonUrgent; break;
              case "urgent": _priority = Priority.Urgent; break;
              default: _priority = Priority.Normal; break;
            }
            break;
          case HeaderId.XPriority:
            SkipWhiteSpace(rawValue, ref index, rawValue.Length);

            if (TryParseInt32(rawValue, ref index, rawValue.Length, out number))
            {
              _xpriority = (XPriority)Math.Min(Math.Max(number, 1), 5);
            }
            else
            {
              _xpriority = XPriority.Normal;
            }
            break;
          case HeaderId.Date:
            DateUtils.TryParse(rawValue, 0, rawValue.Length, out _date);
            break;
          case HeaderId.Subject:
            _subject = header.Value;
            break;
          default:
            if (_standardAddressHeaders.Contains(header.Field))
            {
              if (MimeKit.InternetAddressList.TryParse(options, rawValue, 0, rawValue.Length, out addresses))
              {
                _addresses[header.Field] = InternetAddress.ToAddresses(addresses);
              }
            }
            break;
        }
      }
    }

    private IEnumerable<InternetAddress> GetAddresses(string name)
    {
      InternetAddress[] result;
      if (_addresses.TryGetValue(name, out result))
        return result;
      return Enumerable.Empty<InternetAddress>();
    }

    private static HashSet<string> _standardAddressHeaders = new HashSet<string>(new string[]
    {
      "Resent-From", "Resent-Reply-To", "Resent-To", "Resent-Cc", "Resent-Bcc",
      "From", "Reply-To", "To", "Cc", "Bcc"
    }, StringComparer.OrdinalIgnoreCase);

    private static bool SkipWhiteSpace(byte[] text, ref int index, int endIndex)
    {
      int startIndex = index;

      while (index < endIndex && char.IsWhiteSpace((char)text[index]))
        index++;

      return index > startIndex;
    }

    private static bool TryParseInt32(byte[] text, ref int index, int endIndex, out int value)
    {
      int startIndex = index;

      value = 0;

      while (index < endIndex && text[index] >= (byte)'0' && text[index] <= (byte)'9')
      {
        int digit = text[index] - (byte)'0';

        if (value > int.MaxValue / 10)
        {
          // integer overflow
          return false;
        }

        if (value == int.MaxValue / 10 && digit > int.MaxValue % 10)
        {
          // integer overflow
          return false;
        }

        value = (value * 10) + digit;
        index++;
      }

      return index > startIndex;
    }
  }
}
