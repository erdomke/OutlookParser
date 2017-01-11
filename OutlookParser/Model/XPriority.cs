using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  /// <summary>
  /// An enumeration of X-Priority header values.
  /// </summary>
  /// <remarks>
  /// Indicates the priority of a message.
  /// </remarks>
  public enum XPriority
  {
    /// <summary>
    /// The message is of the highest priority.
    /// </summary>
    Highest = 1,

    /// <summary>
    /// The message is high priority.
    /// </summary>
    High = 2,

    /// <summary>
    /// The message is of normal priority.
    /// </summary>
    Normal = 3,

    /// <summary>
    /// The message is of low priority.
    /// </summary>
    Low = 4,

    /// <summary>
    /// The message is of lowest priority.
    /// </summary>
    Lowest = 5
  }
}
