using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  [DebuggerDisplay("{Address,nq}")]
  public class MailboxAddress : InternetAddress
  {
    internal MailboxAddress(MimeKit.MailboxAddress source)
    {
      this.Address = source.Address;
      this.Name = source.Name;
      this.Route = source.Route.ToArray();
    }

    public string Address { get; set; }
    public IEnumerable<string> Route { get; set; }
  }
}
