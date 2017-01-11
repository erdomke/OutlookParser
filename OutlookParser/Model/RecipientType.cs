using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  public enum RecipientType
  {
    Unknown = 0,
    To = 1,
    Cc = 2,
    Bcc = 3,
    Resource = 4,
    Room = 7
  }
}
