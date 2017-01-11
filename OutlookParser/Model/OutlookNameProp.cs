using OpenMcdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  internal class OutlookNameProp : OutlookStorage
  {
    public OutlookNameProp(OutlookStorage parent, CFStorage storage) 
      : base(parent, storage) { }

    public string PropIdFromName(string propId)
    {
      byte[] data = this.GetMapiPropertyBytes("0003");
      for (int i = 0; i < data.Length; i += 8)
      {
        if (BitConverter.ToInt32(data, i).ToString("X4") == propId)
        {
          return (BitConverter.ToInt16(data, i + 6) + 0x8000).ToString("X4");
        }
      }
      return null;
    }
  }
}
