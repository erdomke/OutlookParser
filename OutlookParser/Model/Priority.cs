using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  /// <summary>
  /// An enumeration of message priority values.
  /// </summary>
  /// <remarks>
  /// Indicates the priority of a message.
  /// </remarks>
  public enum Priority
  {
    /// <summary>
    /// The message has non-urgent priority.
    /// </summary>
    NonUrgent,

    /// <summary>
    /// The message has normal priority.
    /// </summary>
    Normal,

    /// <summary>
    /// The message has urgent priority.
    /// </summary>
    Urgent
  }
}
