using OpenMcdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookParser
{
  /// <summary>
  /// Class used to contain all the appointment information of a <see cref="Storage.Message"/>.
  /// </summary>
  public class OutlookAppointment : OutlookStorage
  {
    #region Public enum AppointmentRecurrenceType
    /// <summary>
    /// The recurrence type of an appointment
    /// </summary>
    public enum AppointmentRecurrenceType
    {
      /// <summary>
      /// There is no reccurence
      /// </summary>
      None = -1,

      /// <summary>
      /// The appointment is daily
      /// </summary>
      Daily = 0,

      /// <summary>
      /// The appointment is weekly
      /// </summary>
      Weekly = 1,

      /// <summary>
      /// The appointment is monthly
      /// </summary>
      Montly = 2,

      /// <summary>
      /// The appointment is yearly
      /// </summary>
      Yearly = 3
    }
    #endregion

    #region Public enum AppointmentClientIntent
    /// <summary>
    /// The intent of an appointment
    /// </summary>
    public enum AppointmentClientIntent
    {
      /// <summary>
      /// The user is the owner of the Meeting object's
      /// </summary>
      Manager = 1,

      /// <summary>
      /// The user is a delegate acting on a Meeting object in a delegator's Calendar folder. If this bit is set, the ciManager bit SHOULD NOT be set
      /// </summary>
      Delegate = 2,

      /// <summary>
      /// The user deleted the Meeting object with no response sent to the organizer
      /// </summary>
      DeletedWithNoResponse = 4,

      /// <summary>
      /// The user deleted an exception to a recurring series with no response sent to the organizer
      /// </summary>
      DeletedExceptionWithNoResponse = 8,

      /// <summary>
      /// Appointment accepted as tentative
      /// </summary>
      RespondedTentative = 16,

      /// <summary>
      /// Appointment accepted
      /// </summary>
      RespondedAccept = 32,

      /// <summary>
      /// Appointment declined
      /// </summary>
      RespondedDecline = 64,

      /// <summary>
      /// The user modified the start time
      /// </summary>
      ModifiedStartTime = 128,

      /// <summary>
      /// The user modified the end time
      /// </summary>
      ModifiedEndTime = 256,

      /// <summary>
      /// The user changed the location of the meeting
      /// </summary>
      ModifiedLocation = 512,

      /// <summary>
      /// The user declined an exception to a recurring series
      /// </summary>
      RespondedExceptionDecline = 1024,

      /// <summary>
      /// The user declined an exception to a recurring series
      /// </summary>
      Canceled = 2048,

      /// <summary>
      /// The user canceled an exception to a recurring serie
      /// </summary>
      ExceptionCanceled = 4096
    }
    #endregion

    #region Properties
    /// <summary>
    /// Returns the location for the appointment, null when not available
    /// </summary>
    public string Location { get; private set; }

    /// <summary>
    /// Returns the start time for the appointment, null when not available
    /// </summary>
    public DateTime? Start { get; private set; }

    /// <summary>
    /// Returns the end time for the appointment, null when not available
    /// </summary>
    public DateTime? End { get; private set; }

    /// <summary>
    /// Returns a string with all the attendees (To and CC), if you also want their E-mail addresses then
    /// get the <see cref="Storage.Message.Recipients"/> from the message, null when not available
    /// </summary>
    public string AllAttendees { get; private set; }

    /// <summary>
    /// Returns a string with all the TO (mandatory) attendees. If you also want their E-mail addresses then
    /// get the <see cref="Storage.Message.Recipients"/> from the <see cref="Storage.Message"/> and filter this 
    /// one on <see cref="Storage.Recipient.RecipientType.To"/>. Null when not available
    /// </summary>
    public string ToAttendees { get; private set; }

    /// <summary>
    /// Returns a string with all the CC (optional) attendees. If you also want their E-mail addresses then
    /// get the <see cref="Storage.Message.Recipients"/> from the <see cref="Storage.Message"/> and filter this 
    /// one on <see cref="Storage.Recipient.RecipientType.Cc"/>. Null when not available
    /// </summary>
    public string CcAttendees { get; private set; }

    /// <summary>
    /// Returns A value of <c>true</c> for the PidLidAppointmentNotAllowPropose property ([MS-OXPROPS] section 2.17) 
    /// indicates that attendees are not allowed to propose a new date and/or time for the meeting. A value of 
    /// <c>false</c> or the absence of this property indicates that the attendees are allowed to propose a new date 
    /// and/or time. This property is meaningful only on Meeting objects, Meeting Request objects, and Meeting 
    /// Update objects. Null when not available
    /// </summary>
    public bool? NotAllowPropose { get; private set; }

    /// <summary>
    /// Returns a <see cref="UnsendableRecipients"/> object with all the unsendable attendees. Null when not available
    /// </summary>
    public UnsendableRecipients UnsendableRecipients { get; private set; }

    /// <summary>
    /// Returns the reccurence type (daily, weekly, monthly or yearly) for the <see cref="Storage.Appointment"/>
    /// </summary>
    public AppointmentRecurrenceType ReccurrenceType { get; private set; }

    /// <summary>
    /// Returns the reccurence patern for the <see cref="Storage.Appointment"/>, null when not available
    /// </summary>
    public string RecurrencePatern { get; private set; }

    /// <summary>
    /// The clients intention for the the <see cref="Storage.Appointment"/> as a list,
    /// null when not available
    /// of <see cref="AppointmentClientIntent"/>
    /// </summary>
    public IEnumerable<AppointmentClientIntent> ClientIntent { get; private set; }
    #endregion

    #region Constructor
    /// <summary>
    /// Initializes a new instance of the <see cref="Storage.Task" /> class.
    /// </summary>
    /// <param name="message"> The message. </param>
    internal OutlookAppointment(OutlookStorage parent, CFStorage message)
      : base(parent, message)
    {
      //GC.SuppressFinalize(message);
      _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

      Location = GetMapiPropertyString(MapiTags.Location);
      Start = GetMapiPropertyDateTime(MapiTags.AppointmentStartWhole);
      End = GetMapiPropertyDateTime(MapiTags.AppointmentEndWhole);
      AllAttendees = GetMapiPropertyString(MapiTags.AppointmentAllAttendees);
      ToAttendees = GetMapiPropertyString(MapiTags.AppointmentToAttendees);
      CcAttendees = GetMapiPropertyString(MapiTags.AppointmentCCAttendees);
      NotAllowPropose = GetMapiPropertyBool(MapiTags.AppointmentNotAllowPropose);
      UnsendableRecipients = GetUnsendableRecipients(MapiTags.AppointmentUnsendableRecipients);

      #region Recurrence
      var recurrenceType = GetMapiPropertyInt32(MapiTags.ReccurrenceType);
      if (recurrenceType == null)
      {
        ReccurrenceType = AppointmentRecurrenceType.None;
      }
      else
      {
        switch (recurrenceType)
        {
          case 1:
            ReccurrenceType = AppointmentRecurrenceType.Daily;
            break;

          case 2:
            ReccurrenceType = AppointmentRecurrenceType.Weekly;
            break;

          case 3:
          case 4:
            ReccurrenceType = AppointmentRecurrenceType.Montly;
            break;

          case 5:
          case 6:
            ReccurrenceType = AppointmentRecurrenceType.Yearly;
            break;

          default:
            ReccurrenceType = AppointmentRecurrenceType.None;
            break;
        }
      }

      RecurrencePatern = GetMapiPropertyString(MapiTags.ReccurrencePattern);
      #endregion

      #region ClientIntent
      var clientIntentList = new List<AppointmentClientIntent>();
      var clientIntent = GetMapiPropertyInt32(MapiTags.PidLidClientIntent);

      if (clientIntent == null)
        ClientIntent = null;
      else
      {
        var bitwiseValue = (int)clientIntent;

        if ((bitwiseValue & 1) == 1)
        {
          clientIntentList.Add(AppointmentClientIntent.Manager);
        }

        if ((bitwiseValue & 2) == 2)
        {
          clientIntentList.Add(AppointmentClientIntent.Delegate);
        }

        if ((bitwiseValue & 4) == 4)
        {
          clientIntentList.Add(AppointmentClientIntent.DeletedWithNoResponse);
        }

        if ((bitwiseValue & 8) == 8)
        {
          clientIntentList.Add(AppointmentClientIntent.DeletedExceptionWithNoResponse);
        }

        if ((bitwiseValue & 16) == 16)
        {
          clientIntentList.Add(AppointmentClientIntent.RespondedTentative);
        }

        if ((bitwiseValue & 32) == 32)
        {
          clientIntentList.Add(AppointmentClientIntent.RespondedAccept);
        }

        if ((bitwiseValue & 64) == 64)
        {
          clientIntentList.Add(AppointmentClientIntent.RespondedDecline);
        }
        if ((bitwiseValue & 128) == 128)
        {
          clientIntentList.Add(AppointmentClientIntent.ModifiedStartTime);
        }

        if ((bitwiseValue & 256) == 256)
        {
          clientIntentList.Add(AppointmentClientIntent.ModifiedEndTime);
        }

        if ((bitwiseValue & 512) == 512)
        {
          clientIntentList.Add(AppointmentClientIntent.ModifiedLocation);
        }

        if ((bitwiseValue & 1024) == 1024)
        {
          clientIntentList.Add(AppointmentClientIntent.RespondedExceptionDecline);
        }

        if ((bitwiseValue & 2048) == 2048)
        {
          clientIntentList.Add(AppointmentClientIntent.Canceled);
        }

        if ((bitwiseValue & 4096) == 4096)
        {
          clientIntentList.Add(AppointmentClientIntent.ExceptionCanceled);
        }

        ClientIntent = clientIntentList.AsReadOnly();
      }
      #endregion
    }
    #endregion
  }
}
