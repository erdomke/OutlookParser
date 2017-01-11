using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  /// <summary>
  /// An enumeration of message importance values.
  /// </summary>
  /// <remarks>
  /// Indicates the importance of a message.
  /// </remarks>
  public enum Importance
  {
    /// <summary>
    /// The message is of low importance.
    /// </summary>
    Low,

    /// <summary>
    /// The message is of normal importance.
    /// </summary>
    Normal,

    /// <summary>
    /// The message is of high importance.
    /// </summary>
    High
  }
}
