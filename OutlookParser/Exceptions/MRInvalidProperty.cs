using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  /// <summary>
  ///     Raised when a property is invalid
  /// </summary>
  public class MRInvalidProperty : Exception
  {
    internal MRInvalidProperty()
    {
    }

    internal MRInvalidProperty(string message) : base(message)
    {
    }

    internal MRInvalidProperty(string message, Exception inner) : base(message, inner)
    {
    }
  }
}
