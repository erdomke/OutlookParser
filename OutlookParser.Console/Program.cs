using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser.Console
{
  class Program
  {
    static void Main(string[] args)
    {
      using (var file = new FileStream(@"C:\Users\eric.domke\Desktop\BoM Review Approvals.msg", FileMode.Open))
      {
        var parser = new OutlookParser(file);
        var email = parser.ParseMessage();
        email.WriteTo(@"C:\Users\eric.domke\Desktop\Test.eml");
      }
    }
  }
}
