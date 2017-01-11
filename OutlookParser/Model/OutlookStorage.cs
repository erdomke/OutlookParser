using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenMcdf;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;

namespace OutlookParser
{
  public class OutlookStorage : IDisposable
  {
    #region CLZF (This Region Has A Seperate Licence)

    /*
         * Copyright (c) 2005 Oren J. Maurice <oymaurice@hazorea.org.il>
         * 
         * Redistribution and use in source and binary forms, with or without modifica-
         * tion, are permitted provided that the following conditions are met:
         * 
         *   1.  Redistributions of source code must retain the above copyright notice,
         *       this list of conditions and the following disclaimer.
         * 
         *   2.  Redistributions in binary form must reproduce the above copyright
         *       notice, this list of conditions and the following disclaimer in the
         *       documentation and/or other materials provided with the distribution.
         * 
         *   3.  The name of the author may not be used to endorse or promote products
         *       derived from this software without specific prior written permission.
         * 
         * THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED
         * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MER-
         * CHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO
         * EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPE-
         * CIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
         * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
         * OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
         * WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTH-
         * ERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
         * OF THE POSSIBILITY OF SUCH DAMAGE.
         *
         * Alternatively, the contents of this file may be used under the terms of
         * the GNU General Public License version 2 (the "GPL"), in which case the
         * provisions of the GPL are applicable instead of the above. If you wish to
         * allow the use of your version of this file only under the terms of the
         * GPL and not to allow others to use your version of this file under the
         * BSD license, indicate your decision by deleting the provisions above and
         * replace them with the notice and other provisions required by the GPL. If
         * you do not delete the provisions above, a recipient may use your version
         * of this file under either the BSD or the GPL.
         */

    /// <summary>
    /// Summary description for CLZF.
    /// </summary>
    public class CLZF
    {
      /*
       This program is free software; you can redistribute it and/or modify
       it under the terms of the GNU General Public License as published by
       the Free Software Foundation; either version 2 of the License, or
       (at your option) any later version.

       You should have received a copy of the GNU General Public License
       along with this program; if not, write to the Free Software Foundation,
       Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA
      */

      /*
       * Prebuffered bytes used in RTF-compressed format (found them in RTFLIB32.LIB)
      */
      static byte[] COMPRESSED_RTF_PREBUF;
      static string prebuf = "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}" +
          "{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript " +
          "\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier" +
          "{\\colortbl\\red0\\green0\\blue0\n\r\\par " +
          "\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx";

      /* The lookup table used in the CRC32 calculation */
      static uint[] CRC32_TABLE =
          {
            0x00000000, 0x77073096, 0xEE0E612C, 0x990951BA, 0x076DC419,
            0x706AF48F, 0xE963A535, 0x9E6495A3, 0x0EDB8832, 0x79DCB8A4,
            0xE0D5E91E, 0x97D2D988, 0x09B64C2B, 0x7EB17CBD, 0xE7B82D07,
            0x90BF1D91, 0x1DB71064, 0x6AB020F2, 0xF3B97148, 0x84BE41DE,
            0x1ADAD47D, 0x6DDDE4EB, 0xF4D4B551, 0x83D385C7, 0x136C9856,
            0x646BA8C0, 0xFD62F97A, 0x8A65C9EC, 0x14015C4F, 0x63066CD9,
            0xFA0F3D63, 0x8D080DF5, 0x3B6E20C8, 0x4C69105E, 0xD56041E4,
            0xA2677172, 0x3C03E4D1, 0x4B04D447, 0xD20D85FD, 0xA50AB56B,
            0x35B5A8FA, 0x42B2986C, 0xDBBBC9D6, 0xACBCF940, 0x32D86CE3,
            0x45DF5C75, 0xDCD60DCF, 0xABD13D59, 0x26D930AC, 0x51DE003A,
            0xC8D75180, 0xBFD06116, 0x21B4F4B5, 0x56B3C423, 0xCFBA9599,
            0xB8BDA50F, 0x2802B89E, 0x5F058808, 0xC60CD9B2, 0xB10BE924,
            0x2F6F7C87, 0x58684C11, 0xC1611DAB, 0xB6662D3D, 0x76DC4190,
            0x01DB7106, 0x98D220BC, 0xEFD5102A, 0x71B18589, 0x06B6B51F,
            0x9FBFE4A5, 0xE8B8D433, 0x7807C9A2, 0x0F00F934, 0x9609A88E,
            0xE10E9818, 0x7F6A0DBB, 0x086D3D2D, 0x91646C97, 0xE6635C01,
            0x6B6B51F4, 0x1C6C6162, 0x856530D8, 0xF262004E, 0x6C0695ED,
            0x1B01A57B, 0x8208F4C1, 0xF50FC457, 0x65B0D9C6, 0x12B7E950,
            0x8BBEB8EA, 0xFCB9887C, 0x62DD1DDF, 0x15DA2D49, 0x8CD37CF3,
            0xFBD44C65, 0x4DB26158, 0x3AB551CE, 0xA3BC0074, 0xD4BB30E2,
            0x4ADFA541, 0x3DD895D7, 0xA4D1C46D, 0xD3D6F4FB, 0x4369E96A,
            0x346ED9FC, 0xAD678846, 0xDA60B8D0, 0x44042D73, 0x33031DE5,
            0xAA0A4C5F, 0xDD0D7CC9, 0x5005713C, 0x270241AA, 0xBE0B1010,
            0xC90C2086, 0x5768B525, 0x206F85B3, 0xB966D409, 0xCE61E49F,
            0x5EDEF90E, 0x29D9C998, 0xB0D09822, 0xC7D7A8B4, 0x59B33D17,
            0x2EB40D81, 0xB7BD5C3B, 0xC0BA6CAD, 0xEDB88320, 0x9ABFB3B6,
            0x03B6E20C, 0x74B1D29A, 0xEAD54739, 0x9DD277AF, 0x04DB2615,
            0x73DC1683, 0xE3630B12, 0x94643B84, 0x0D6D6A3E, 0x7A6A5AA8,
            0xE40ECF0B, 0x9309FF9D, 0x0A00AE27, 0x7D079EB1, 0xF00F9344,
            0x8708A3D2, 0x1E01F268, 0x6906C2FE, 0xF762575D, 0x806567CB,
            0x196C3671, 0x6E6B06E7, 0xFED41B76, 0x89D32BE0, 0x10DA7A5A,
            0x67DD4ACC, 0xF9B9DF6F, 0x8EBEEFF9, 0x17B7BE43, 0x60B08ED5,
            0xD6D6A3E8, 0xA1D1937E, 0x38D8C2C4, 0x4FDFF252, 0xD1BB67F1,
            0xA6BC5767, 0x3FB506DD, 0x48B2364B, 0xD80D2BDA, 0xAF0A1B4C,
            0x36034AF6, 0x41047A60, 0xDF60EFC3, 0xA867DF55, 0x316E8EEF,
            0x4669BE79, 0xCB61B38C, 0xBC66831A, 0x256FD2A0, 0x5268E236,
            0xCC0C7795, 0xBB0B4703, 0x220216B9, 0x5505262F, 0xC5BA3BBE,
            0xB2BD0B28, 0x2BB45A92, 0x5CB36A04, 0xC2D7FFA7, 0xB5D0CF31,
            0x2CD99E8B, 0x5BDEAE1D, 0x9B64C2B0, 0xEC63F226, 0x756AA39C,
            0x026D930A, 0x9C0906A9, 0xEB0E363F, 0x72076785, 0x05005713,
            0x95BF4A82, 0xE2B87A14, 0x7BB12BAE, 0x0CB61B38, 0x92D28E9B,
            0xE5D5BE0D, 0x7CDCEFB7, 0x0BDBDF21, 0x86D3D2D4, 0xF1D4E242,
            0x68DDB3F8, 0x1FDA836E, 0x81BE16CD, 0xF6B9265B, 0x6FB077E1,
            0x18B74777, 0x88085AE6, 0xFF0F6A70, 0x66063BCA, 0x11010B5C,
            0x8F659EFF, 0xF862AE69, 0x616BFFD3, 0x166CCF45, 0xA00AE278,
            0xD70DD2EE, 0x4E048354, 0x3903B3C2, 0xA7672661, 0xD06016F7,
            0x4969474D, 0x3E6E77DB, 0xAED16A4A, 0xD9D65ADC, 0x40DF0B66,
            0x37D83BF0, 0xA9BCAE53, 0xDEBB9EC5, 0x47B2CF7F, 0x30B5FFE9,
            0xBDBDF21C, 0xCABAC28A, 0x53B39330, 0x24B4A3A6, 0xBAD03605,
            0xCDD70693, 0x54DE5729, 0x23D967BF, 0xB3667A2E, 0xC4614AB8,
            0x5D681B02, 0x2A6F2B94, 0xB40BBE37, 0xC30C8EA1, 0x5A05DF1B,
            0x2D02EF8D
          };

      /*
       * Calculates the CRC32 of the given bytes.
       * The CRC32 calculation is similar to the standard one as demonstrated
       * in RFC 1952, but with the inversion (before and after the calculation)
       * ommited.
       * 
       * @param buf the byte array to calculate CRC32 on
       * @param off the offset within buf at which the CRC32 calculation will start
       * @param len the number of bytes on which to calculate the CRC32
       * @return the CRC32 value.
       */
      static public int calculateCRC32(byte[] buf, int off, int len)
      {
        uint c = 0;
        int end = off + len;
        for (int i = off; i < end; i++)
        {
          //!!!!        c = CRC32_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
          c = CRC32_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >> 8);
        }
        return (int)c;
      }

      /*
           * Returns an unsigned 32-bit value from little-endian ordered bytes.
           *
           * @param   buf a byte array from which byte values are taken
           * @param   offset the offset within buf from which byte values are taken
           * @return  an unsigned 32-bit value as a long.
      */
      public static long getU32(byte[] buf, int offset)
      {
        return ((buf[offset] & 0xFF) | ((buf[offset + 1] & 0xFF) << 8) | ((buf[offset + 2] & 0xFF) << 16) | ((buf[offset + 3] & 0xFF) << 24)) & 0x00000000FFFFFFFFL;
      }

      /*
       * Returns an unsigned 8-bit value from a byte array.
       *
       * @param   buf a byte array from which byte value is taken
       * @param   offset the offset within buf from which byte value is taken
       * @return  an unsigned 8-bit value as an int.
       */
      public static int getU8(byte[] buf, int offset)
      {
        return buf[offset] & 0xFF;
      }

      /*
        * Decompresses compressed-RTF data.
        *
        * @param   src the compressed-RTF data bytes
        * @return  an array containing the decompressed bytes.
        * @throws  IllegalArgumentException if src does not contain valid                                                                                                                                            *          compressed-RTF bytes.
      */
      public static byte[] decompressRTF(byte[] src)
      {
        byte[] dst; // destination for uncompressed bytes
        int inPos = 0; // current position in src array
        int outPos = 0; // current position in dst array

        COMPRESSED_RTF_PREBUF = System.Text.Encoding.ASCII.GetBytes(prebuf);

        // get header fields (as defined in RTFLIB.H)
        if (src == null || src.Length < 16)
          throw new Exception("Invalid compressed-RTF header");

        int compressedSize = (int)getU32(src, inPos);
        inPos += 4;
        int uncompressedSize = (int)getU32(src, inPos);
        inPos += 4;
        int magic = (int)getU32(src, inPos);
        inPos += 4;
        
        if (compressedSize != src.Length - 4) // check size excluding the size field itself
          throw new Exception("compressed-RTF data size mismatch");

        // process the data
        if (magic == 0x414c454d)
        { // magic number that identifies the stream as a uncompressed stream
          dst = new byte[uncompressedSize - inPos];
          Array.Copy(src, inPos + 4, dst, outPos, uncompressedSize - inPos); // just copy it as it is
        }
        else if (magic == 0x75465a4c)
        { // magic number that identifies the stream as a compressed stream
          int crc32 = (int)getU32(src, inPos);
          inPos += 4;

          if (crc32 != calculateCRC32(src, 16, src.Length - 16))
            throw new Exception("compressed-RTF CRC32 failed");

          dst = new byte[COMPRESSED_RTF_PREBUF.Length + uncompressedSize];
          Array.Copy(COMPRESSED_RTF_PREBUF, 0, dst, 0, COMPRESSED_RTF_PREBUF.Length);
          outPos = COMPRESSED_RTF_PREBUF.Length;
          int flagCount = 0;
          int flags = 0;
          while (outPos < dst.Length)
          {
            // each flag byte flags 8 literals/references, 1 per bit
            flags = (flagCount++ % 8 == 0) ? getU8(src, inPos++) : flags >> 1;
            if ((flags & 1) == 1)
            { // each flag bit is 1 for reference, 0 for literal
              int offset = getU8(src, inPos++);
              int length = getU8(src, inPos++);
              //!!!!!!!!!            offset = (offset << 4) | (length >>> 4); // the offset relative to block start
              offset = (offset << 4) | (length >> 4); // the offset relative to block start
              length = (length & 0xF) + 2; // the number of bytes to copy
              // the decompression buffer is supposed to wrap around back
              // to the beginning when the end is reached. we save the
              // need for such a buffer by pointing straight into the data
              // buffer, and simulating this behaviour by modifying the
              // pointers appropriately.
              offset = (outPos / 4096) * 4096 + offset;
              if (offset >= outPos) // take from previous block
                offset -= 4096;
              // note: can't use System.arraycopy, because the referenced
              // bytes can cross through the current out position.
              int end = offset + length;
              while (offset < end)
                dst[outPos++] = dst[offset++];
            }
            else
            { // literal
              dst[outPos++] = src[inPos++];
            }
          }
          // copy it back without the prebuffered data
          src = dst;
          dst = new byte[uncompressedSize];
          Array.Copy(src, COMPRESSED_RTF_PREBUF.Length, dst, 0, uncompressedSize);
        }
        else
        { // unknown magic number
          throw new Exception("Unknown compression type (magic number " + magic + ")");
        }

        return dst;
      }
    }

    #endregion

    /// <summary>
    /// Header size of the property stream in the IStorage associated with this instance.
    /// </summary>
    protected int _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

    /// <summary>
    /// A reference to the parent message that this message may belong to.
    /// </summary>
    protected OutlookStorage parentMessage = null;

    /// <summary>
    /// Indicates wether this instance has been disposed.
    /// </summary>
    protected bool disposed = false;

    protected CompoundFile _file;
    protected CFStorage _storage;
    protected List<string> _keys;


    /// <summary>
    /// Contains the <see cref="MessageType"/> of this Message
    /// </summary>
    private MessageType _type = MessageType.Unknown;
    public MessageType Type
    {
      get
      {
        if (_type != MessageType.Unknown)
          return _type;

        var type = GetMapiPropertyString(MapiTags.PR_MESSAGE_CLASS);
        if (type == null)
          return MessageType.Unknown;

        switch (type.ToUpperInvariant())
        {
          case "IPM.NOTE":
            _type = MessageType.Email;
            break;

          case "IPM.NOTE.MOBILE.SMS":
            _type = MessageType.EmailSms;
            break;

          case "REPORT.IPM.NOTE.NDR":
            _type = MessageType.EmailNonDeliveryReport;
            break;

          case "REPORT.IPM.NOTE.DR":
            _type = MessageType.EmailDeliveryReport;
            break;

          case "REPORT.IPM.NOTE.DELAYED":
            _type = MessageType.EmailDelayedDeliveryReport;
            break;

          case "REPORT.IPM.NOTE.IPNRN":
            _type = MessageType.EmailReadReceipt;
            break;

          case "REPORT.IPM.NOTE.IPNNRN":
            _type = MessageType.EmailNonReadReceipt;
            break;

          case "IPM.NOTE.SMIME":
            _type = MessageType.EmailEncryptedAndMaybeSigned;
            break;

          case "REPORT.IPM.NOTE.SMIME.NDR":
            _type = MessageType.EmailEncryptedAndMaybeSignedNonDelivery;
            break;

          case "REPORT.IPM.NOTE.SMIME.DR":
            _type = MessageType.EmailEncryptedAndMaybeSignedDelivery;
            break;

          case "IPM.NOTE.SMIME.MULTIPARTSIGNED":
            _type = MessageType.EmailClearSigned;
            break;

          case "IPM.NOTE.RECEIPT.SMIME.MULTIPARTSIGNED":
            _type = MessageType.EmailClearSigned;
            break;

          case "REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.NDR":
            _type = MessageType.EmailClearSignedNonDelivery;
            break;

          case "REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.DR":
            _type = MessageType.EmailClearSignedDelivery;
            break;

          case "IPM.NOTE.BMA.STUB":
            _type = MessageType.EmailBmaStub;
            break;

          case "IPM.APPOINTMENT":
            _type = MessageType.Appointment;
            break;

          case "IPM.SCHEDULE.MEETING":
            _type = MessageType.AppointmentSchedule;
            break;

          case "IPM.NOTIFICATION.MEETING":
            _type = MessageType.AppointmentNotification;
            break;

          case "IPM.SCHEDULE.MEETING.REQUEST":
            _type = MessageType.AppointmentRequest;
            break;

          case "IPM.SCHEDULE.MEETING.REQUEST.NDR":
            _type = MessageType.AppointmentRequestNonDelivery;
            break;

          case "IPM.SCHEDULE.MEETING.CANCELED":
            _type = MessageType.AppointmentResponseCanceled;
            break;

          case "IPM.SCHEDULE.MEETING.CANCELED.NDR":
            _type = MessageType.AppointmentResponseCanceledNonDelivery;
            break;

          case "IPM.SCHEDULE.MEETING.RESPONSE":
            _type = MessageType.AppointmentResponse;
            break;

          case "IPM.SCHEDULE.MEETING.RESP.POS":
            _type = MessageType.AppointmentResponsePositive;
            break;

          case "IPM.SCHEDULE.MEETING.RESP.POS.NDR":
            _type = MessageType.AppointmentResponsePositiveNonDelivery;
            break;

          case "IPM.SCHEDULE.MEETING.RESP.NEG":
            _type = MessageType.AppointmentResponseNegative;
            break;

          case "IPM.SCHEDULE.MEETING.RESP.NEG.NDR":
            _type = MessageType.AppointmentResponseNegativeNonDelivery;
            break;

          case "IPM.SCHEDULE.MEETING.RESP.TENT":
            _type = MessageType.AppointmentResponseTentative;
            break;

          case "IPM.SCHEDULE.MEETING.RESP.TENT.NDR":
            _type = MessageType.AppointmentResponseTentativeNonDelivery;
            break;

          case "IPM.CONTACT":
            _type = MessageType.Contact;
            break;

          case "IPM.TASK":
            _type = MessageType.Task;
            break;

          case "IPM.TASKREQUEST.ACCEPT":
            _type = MessageType.TaskRequestAccept;
            break;

          case "IPM.TASKREQUEST.DECLINE":
            _type = MessageType.TaskRequestDecline;
            break;

          case "IPM.TASKREQUEST.UPDATE":
            _type = MessageType.TaskRequestUpdate;
            break;

          case "IPM.STICKYNOTE":
            _type = MessageType.StickyNote;
            break;

          case "IPM.NOTE.CUSTOM.CISCO.UNITY.VOICE":
            _type = MessageType.CiscoUnityVoiceMessage;
            break;
        }

        return _type;
      }
    }

    /// <summary>
    /// Gets a value indicating whether this instance is the top level outlook message.
    /// </summary>
    /// <value>
    /// 	<c>true</c> if this instance is the top level outlook message; otherwise, <c>false</c>.
    /// </value>
    private bool IsTopParent
    {
      get
      {
        if (this.parentMessage != null)
        {
          return false;
        }
        return true;
      }
    }

    /// <summary>
    /// Gets the top level outlook message from a sub message at any level.
    /// </summary>
    /// <value>The top level outlook message.</value>
    private OutlookStorage TopParent
    {
      get
      {
        if (this.parentMessage != null)
        {
          return this.parentMessage.TopParent;
        }
        return this;
      }
    }

    public OutlookStorage(string path)
    {
      _file = new CompoundFile(path);
      LoadStorage(_file.RootStorage);
    }
    public OutlookStorage(Stream stream)
    {
      _file = new CompoundFile(stream);
      LoadStorage(_file.RootStorage);
    }
    internal OutlookStorage(OutlookStorage parent, CFStorage storage)
      : this(parent._file, storage) { }
    internal OutlookStorage(CompoundFile file, CFStorage storage)
    {
      _file = file;
      LoadStorage(storage);
    }

    protected virtual void LoadStorage(CFStorage storage)
    {
      _storage = storage;
      _keys = new List<string>();
      _storage.VisitEntries(i => _keys.Add(i.Name), false);
    }

    public void WriteTo(Stream stream)
    {
      var output = new CompoundFile();
      var names = _file.RootStorage.TryGetStorage(MapiTags.NameIdStorage);
      if (names != null)
        CopyStorage(names, output.RootStorage.AddStorage(MapiTags.NameIdStorage));
      CopyStorage(_storage, output.RootStorage);
      output.Save(stream);
    }

    private void CopyStorage(CFStorage source, CFStorage dest)
    {
      source.VisitEntries(i =>
      {
        if (i.IsStorage)
        {
          CopyStorage((CFStorage)i, dest.AddStorage(i.Name));
        }
        else if (i.IsStream)
        {
          var newStream = dest.AddStream(i.Name);
          var currStream = (CFStream)i;
          newStream.SetData(currStream.GetData());
        }
      }, false);
    }

    #region Methods(GetMapiProperty)
    /// <summary>
    /// Gets the raw value of the MAPI property.
    /// </summary>
    /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
    /// <returns>The raw value of the MAPI property.</returns>
    public object GetMapiProperty(string propIdentifier)
    {
      //try get prop value from stream or storage
      object propValue = this.GetMapiPropertyFromStreamOrStorage(propIdentifier);

      //if not found in stream or storage try get prop value from property stream
      if (propValue == null)
      {
        propValue = this.GetMapiPropertyFromPropertyStream(propIdentifier);
      }

      return propValue;
    }

    /// <summary>
    /// Gets the MAPI property value from a stream or storage in this storage.
    /// </summary>
    /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
    /// <returns>The value of the MAPI property or null if not found.</returns>
    private object GetMapiPropertyFromStreamOrStorage(string propIdentifier)
    {
      //determine if the property identifier is in a stream or sub storage
      string propTag = null;
      var propType = PropertyType.PT_UNSPECIFIED;
      foreach (string propKey in _keys)
      {
        if (propKey.StartsWith(MapiTags.SubStgVersion1 + propIdentifier))
        {
          propTag = propKey.Substring(12, 8);
          propType = (PropertyType)ushort.Parse(propKey.Substring(16, 4), System.Globalization.NumberStyles.HexNumber);
          break;
        }
      }

      //depending on prop type use method to get property value
      string containerName = MapiTags.SubStgVersion1 + propTag;
      switch (propType)
      {
        case PropertyType.PT_UNSPECIFIED:
          return null;

        case PropertyType.PT_STRING8:
          return this.GetStreamAsString(containerName, Encoding.Default);

        case PropertyType.PT_UNICODE:
          return this.GetStreamAsString(containerName, Encoding.Unicode);

        case PropertyType.PT_BINARY:
          return _storage.GetStream(containerName).GetData();

        case PropertyType.PT_OBJECT:
          return _storage.GetStorage(containerName);

        case PropertyType.PT_MV_STRING8:
        case PropertyType.PT_MV_UNICODE:

          // If the property is a unicode multiview item we need to read all the properties
          // again and filter out all the multivalue names, they end with -00000000, -00000001, etc..
          var multiValueContainerNames = _keys.Where(propKey => propKey.StartsWith(containerName + "-")).ToList();

          var values = new List<string>();
          foreach (var multiValueContainerName in multiValueContainerNames)
          {
            var value = GetStreamAsString(multiValueContainerName,
                propType == PropertyType.PT_MV_STRING8 ? Encoding.Default : Encoding.Unicode);

            // Multi values always end with a null char so we need to strip that one off
            if (value.EndsWith("/0"))
              value = value.Substring(0, value.Length - 1);

            values.Add(value);
          }

          return values;
        default:
          throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
      }
    }

    /// <summary>
    /// Gets the MAPI property value from the property stream in this storage.
    /// </summary>
    /// <param name="propIdentifier">The 4 char hexadecimal prop identifier.</param>
    /// <returns>The value of the MAPI property or null if not found.</returns>
    private object GetMapiPropertyFromPropertyStream(string propIdentifier)
    {
      var propStream = _storage.GetStream(MapiTags.PropertiesStream);
      
      //if no property stream return null
      if (propStream == null) return null;

      //get the raw bytes for the property stream
      byte[] propBytes = propStream.GetData();

      //iterate over property stream in 16 byte chunks starting from end of header
      for (int i = this._propHeaderSize; i < propBytes.Length; i = i + 16)
      {
        //get property type located in the 1st and 2nd bytes as a unsigned short value
        var propType = (PropertyType)BitConverter.ToUInt16(propBytes, i);

        //get property identifer located in 3nd and 4th bytes as a hexdecimal string
        byte[] propIdent = new byte[] { propBytes[i + 3], propBytes[i + 2] };
        string propIdentString = BitConverter.ToString(propIdent).Replace("-", "");

        //if this is not the property being gotten continue to next property
        if (propIdentString != propIdentifier)
        {
          continue;
        }

        //depending on prop type use method to get property value
        switch (propType)
        {
          case PropertyType.PT_SHORT:
            return BitConverter.ToInt16(propBytes, i + 8);

          case PropertyType.PT_LONG:
            return BitConverter.ToInt32(propBytes, i + 8);

          case PropertyType.PT_DOUBLE:
            return BitConverter.ToDouble(propBytes, i + 8);

          case PropertyType.PT_SYSTIME:
            var fileTime = BitConverter.ToInt64(propBytes, i + 8);
            return DateTime.FromFileTime(fileTime);

          case PropertyType.PT_APPTIME:
            var appTime = BitConverter.ToInt64(propBytes, i + 8);
            return DateTime.FromOADate(appTime);

          case PropertyType.PT_BOOLEAN:
            return BitConverter.ToBoolean(propBytes, i + 8);

          default:
            throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
        }
      }

      //property not found return null
      return null;
    }

    /// <summary>
    /// Gets the data in the specified stream as a string using the specifed encoding to decode the stream data.
    /// </summary>
    /// <param name="streamName">Name of the stream to get string data for.</param>
    /// <param name="streamEncoding">The encoding to decode the stream data with.</param>
    /// <returns>The data in the specified stream as a string.</returns>
    protected string GetStreamAsString(string streamName, Encoding streamEncoding)
    {
      StreamReader streamReader = new StreamReader(new MemoryStream(_storage.GetStream(streamName).GetData()), streamEncoding);
      string streamContent = streamReader.ReadToEnd();
      streamReader.Close();

      return streamContent;
    }

    /// <summary>
    /// Gets the value of the MAPI property as an <see cref="UnsendableRecipients"/>
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a string. </returns>
    protected UnsendableRecipients GetUnsendableRecipients(string propIdentifier)
    {
      var data = GetMapiPropertyBytes(propIdentifier);
      return new UnsendableRecipients(data);
    }

    /// <summary>
    /// Gets the value of the MAPI property as a string.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a string. </returns>
    internal string GetMapiPropertyString(string propIdentifier)
    {
      return GetMapiProperty(propIdentifier) as string;
    }

    /// <summary>
    /// Gets the value of the MAPI property as a list of string.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a list of string. </returns>
    internal IEnumerable<string> GetMapiPropertyStringList(string propIdentifier)
    {
      var list = GetMapiProperty(propIdentifier) as List<string>;
      return list == null ? null : list.AsReadOnly();
    }

    /// <summary>
    /// Gets the value of the MAPI property as a integer.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
    /// <returns> The value of the MAPI property as a integer. </returns>
    internal int? GetMapiPropertyInt32(string propIdentifier)
    {
      var value = GetMapiProperty(propIdentifier);

      if (value != null)
        return (int)value;

      return null;
    }

    /// <summary>
    /// Gets the value of the MAPI property as a double.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a double. </returns>
    internal double? GetMapiPropertyDouble(string propIdentifier)
    {
      var value = GetMapiProperty(propIdentifier);

      if (value != null)
        return (double)value;

      return null;
    }

    /// <summary>
    /// Gets the value of the MAPI property as a datetime.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a datetime or null when not set </returns>
    internal DateTime? GetMapiPropertyDateTime(string propIdentifier)
    {
      var value = GetMapiProperty(propIdentifier);

      if (value != null)
        return (DateTime)value;

      return null;
    }

    /// <summary>
    /// Gets the value of the MAPI property as a bool.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a boolean or null when not set. </returns>
    internal bool? GetMapiPropertyBool(string propIdentifier)
    {
      var value = GetMapiProperty(propIdentifier);

      if (value != null)
        return (bool)value;

      return null;
    }

    /// <summary>
    /// Gets the value of the MAPI property as a byte array.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a byte array. </returns>
    protected byte[] GetMapiPropertyBytes(string propIdentifier)
    {
      return (byte[])GetMapiProperty(propIdentifier);
    }
    #endregion

    #region IDisposable Members

    /// <summary>
    /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    /// </summary>
    public void Dispose()
    {
      if (!this.disposed)
      {
        //ensure only disposed once
        this.disposed = true;

        //call virtual disposing method to let sub classes clean up
        this.Disposing();

        if (_file != null)
        {
          _file.Close();
          _file = null;
        }
      }
    }

    /// <summary>
    /// Gives sub classes the chance to free resources during object disposal.
    /// </summary>
    protected virtual void Disposing() { }

    #endregion

  }
}
