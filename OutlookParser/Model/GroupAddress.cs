using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  public class GroupAddress : InternetAddress
  {
    internal GroupAddress(MimeKit.GroupAddress source)
    {
      this.Name = source.Name;
      this.Members = ToAddresses(source.Members);
    }

    public IEnumerable<InternetAddress> Members { get; set; }
  }
}
