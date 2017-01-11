using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  public abstract class InternetAddress
  {
    public string Name { get; set; }

    internal static InternetAddress[] ToAddresses(InternetAddressList list)
    {
      return list.Select(m =>
      {
        if (m is MimeKit.GroupAddress)
          return (InternetAddress)new GroupAddress((MimeKit.GroupAddress)m);
        if (m is MimeKit.MailboxAddress)
          return new MailboxAddress((MimeKit.MailboxAddress)m);
        return null;
      })
      .Where(m => m != null)
      .ToArray();
    }
  }
}
