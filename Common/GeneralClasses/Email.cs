using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;

namespace GeneralClasses
{

    /// <summary>
    /// Class Email containing all the functions for handling email
    /// </summary>
    public class Email
    {

        #region Variables

        /// <summary>
        /// List of recipients
        /// </summary>
        List<MapiRecipDesc> m_Recipients = new
            List<MapiRecipDesc>();

        /// <summary>
        /// List of attachments file name
        /// </summary>
        List<string> m_attachments = new List<string>();

        /// <summary>
        /// Integer of the last error
        /// </summary>
        int m_LastError = 0;

        /// <summary>
        /// Address of the default email software
        /// </summary>
        const int MAPI_LOGON_UI = 0x00000001;

        /// <summary>
        /// Address of the dialog email software
        /// </summary>
        const int MAPI_DIALOG = 0x00000008;

        /// <summary>
        /// Number of maximum attachments
        /// </summary>
        const int m_MaxAttachments = 20;

        /// <summary>
        /// Enum of HowTO
        /// </summary>
        enum HowTo { MAPI_ORIG = 0, MAPI_TO, MAPI_CC, MAPI_BCC };

        /// <summary>
        /// Table of possible errors
        /// </summary>
        readonly string[] errors = new string[] {
        "OK [0]", "User abort [1]", "General MAPI failure [2]",
                "MAPI login failure [3]", "Disk full [4]",
                "Insufficient memory [5]", "Access denied [6]",
                "-unknown- [7]", "Too many sessions [8]",
                "Too many files were specified [9]",
                "Too many recipients were specified [10]",
                "A specified attachment was not found [11]",
        "Attachment open failure [12]",
                "Attachment write failure [13]", "Unknown recipient [14]",
                "Bad recipient type [15]", "No messages [16]",
                "Invalid message [17]", "Text too large [18]",
                "Invalid session [19]", "Type not supported [20]",
                "A recipient was specified ambiguously [21]",
                "Message in use [22]", "Network failure [23]",
        "Invalid edit fields [24]", "Invalid recipients [25]",
                "Not supported [26]"
        };

        #endregion

        #region Functions

        /// <summary>
        /// Add recipients to the email 
        /// <param name="_Email">The email</param>
        /// <param name="_HowTo">Which type of recipient</param>
        /// <returns>True if the operation went well, false otherwise</returns>
        /// </summary>
        bool Add_Recipient(string _Email, HowTo _HowTo)
        {
            MapiRecipDesc recipient = new MapiRecipDesc();

            recipient.recipClass = (int)_HowTo;
            recipient.name = _Email;
            m_Recipients.Add(recipient);

            return true;
        }

        /// <summary>
        /// Add recipients to the email type To 
        /// <param name="_Email">The email</param>
        /// <returns>True if the operation went well, false otherwise</returns>
        /// </summary>
        public bool Add_RecipientTo(string _Email)
        {
            return Add_Recipient(_Email, HowTo.MAPI_TO);
        }

        /// <summary>
        /// Ad recipient to the email type CC
        /// <param name="_Email">The Email</param>
        /// <returns>True if the operation went well, false otherwise</returns>
        /// </summary>
        public bool Add_RecipientCC(string _Email)
        {
            return Add_Recipient(_Email, HowTo.MAPI_TO);
        }

        /// <summary>
        /// Add recipients to the email type BCC
        /// <param name="_Email">The Email</param>
        /// <returns>True if the operation went well, false otherwise</returns>
        /// </summary>
        public bool Add_RecipientBCC(string _Email)
        {
            return Add_Recipient(_Email, HowTo.MAPI_TO);
        }

        /// <summary>
        /// Add attachment type File
        /// <param name="_AttachmentFileName">The attachment</param>
        /// </summary>
        public void Add_Attachment(string _AttachmentFileName)
        {
            m_attachments.Add(_AttachmentFileName);
        }

        [DllImport("MAPI32.DLL")]
        static extern int MAPISendMail(IntPtr sess, IntPtr hwnd,
            MapiMessage message, int flg, int rsv);
        /// <summary>
        /// Treat the email (popu or not)
        /// <param name="_Subject">Subject of the email</param>
        /// <param name="_Body">Body of the email</param>
        /// <param name="_How">Way to send the email</param>
        /// <returns>0 if the operation went well, >0 otherwise</returns>
        /// </summary>
        int Send_Mail(string _Subject, string _Body, int _How)
        {
            MapiMessage msg = new MapiMessage();
            msg.subject = _Body;
            msg.noteText = _Subject;

            msg.recips = Get_Recipients(out msg.recipCount);
            msg.files = Get_Attachments(out msg.fileCount);

            m_LastError = MAPISendMail(new IntPtr(0), new IntPtr(0), msg, _How, 0);

            if (m_LastError > 1)
            {
                //Handle the error
                string error = Get_LastError();
                Console.WriteLine(error);
            }
            Cleanup(ref msg);
            return m_LastError;
        }

        /// <summary>
        /// Get the size of recipient
        /// <param name="_RecipCount">Number of recipients</param>
        /// <returns>Size of recipients</returns>
        /// </summary>
        IntPtr Get_Recipients(out int _RecipCount)
        {
            _RecipCount = 0;
            if (m_Recipients.Count == 0)
                return IntPtr.Zero;

            int size = Marshal.SizeOf(typeof(MapiRecipDesc));
            IntPtr intPtr = Marshal.AllocHGlobal(m_Recipients.Count * size);

            int ptr = (int)intPtr;
            foreach (MapiRecipDesc mapiDesc in m_Recipients)
            {
                Marshal.StructureToPtr(mapiDesc, (IntPtr)ptr, false);
                ptr += size;
            }

            _RecipCount = m_Recipients.Count;
            return intPtr;
        }

        /// <summary>
        /// Get the size of attachments
        /// <param name="_FileCount">Number of attachments</param>
        /// <returns>Size of attachments</returns>
        /// </summary>
        IntPtr Get_Attachments(out int _FileCount)
        {
            _FileCount = 0;
            if (m_attachments == null)
                return IntPtr.Zero;

            if ((m_attachments.Count <= 0) || (m_attachments.Count >
                m_MaxAttachments))
                return IntPtr.Zero;

            int size = Marshal.SizeOf(typeof(MapiFileDesc));
            IntPtr intPtr = Marshal.AllocHGlobal(m_attachments.Count * size);

            MapiFileDesc mapiFileDesc = new MapiFileDesc();
            mapiFileDesc.position = -1;
            int ptr = (int)intPtr;

            foreach (string strAttachment in m_attachments)
            {
                mapiFileDesc.name = Path.GetFileName(strAttachment);
                mapiFileDesc.path = strAttachment;
                Marshal.StructureToPtr(mapiFileDesc, (IntPtr)ptr, false);
                ptr += size;
            }

            _FileCount = m_attachments.Count;
            return intPtr;
        }

        /// <summary>
        /// Call the function to open the default emailing software
        /// <param name="_Subject">Subject of the email</param>
        /// <param name="_Body">Body of the email</param>
        /// <returns>0 if the operation went well, >0 otherwise</returns>
        /// </summary>
        public int Send_MailPopup(string _Subject, string _Body)
        {
            return Send_Mail(_Subject, _Body, MAPI_LOGON_UI | MAPI_DIALOG);
        }

        /// <summary>
        /// Call the function to send the email directly
        /// <param name="_Subject">Subject of the email</param>
        /// <param name="_Body">Body of the email</param>
        /// <returns>0 if the operation went well, >0 otherwise</returns>
        /// </summary>
        public int Send_MailDirect(string _Subject, string _Body)
        {
            return Send_Mail(_Subject, _Body, MAPI_LOGON_UI);
        }

        /// <summary>
        /// Call the function to send the email adding the recipients and attachments
        /// <param name="_Recipients">List of recipients</param>
        /// <param name="_Subject">Subject of the email</param>
        /// <param name="_Body">Body of the email</param>
        /// <param name="_Attachments">List of attachments</param>
        /// </summary>
        public void Compose_Mail(string[] _Recipients, string _Subject, string _Body, string[] _Attachments)
        {
            try
            {
                foreach (string recipient in _Recipients)
                    this.Add_RecipientTo(recipient);
                foreach (string attachment in _Attachments)
                    this.Add_Attachment(attachment);
                this.Send_MailPopup(_Body, _Subject);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Clean all memory and variables
        /// <param name="_Message">Mapi message</param>
        /// </summary>
        void Cleanup(ref MapiMessage _Message)
        {
            int size = Marshal.SizeOf(typeof(MapiRecipDesc));
            int ptr = 0;

            if (_Message.recips != IntPtr.Zero)
            {
                ptr = (int)_Message.recips;
                for (int i = 0; i < _Message.recipCount; i++)
                {
                    Marshal.DestroyStructure((IntPtr)ptr,
                        typeof(MapiRecipDesc));
                    ptr += size;
                }
                Marshal.FreeHGlobal(_Message.recips);
            }

            if (_Message.files != IntPtr.Zero)
            {
                size = Marshal.SizeOf(typeof(MapiFileDesc));

                ptr = (int)_Message.files;
                for (int i = 0; i < _Message.fileCount; i++)
                {
                    Marshal.DestroyStructure((IntPtr)ptr,
                        typeof(MapiFileDesc));
                    ptr += size;
                }
                Marshal.FreeHGlobal(_Message.files);
            }

            m_Recipients.Clear();
            m_attachments.Clear();
            m_LastError = 0;
        }

        /// <summary>
        /// Get the last error
        /// </summary>
        public string Get_LastError()
        {
            if (m_LastError <= 26)
                return errors[m_LastError];
            return "MAPI error [" + m_LastError.ToString() + "]";
        }

        #endregion

    }

    #region Other classes

    /// <summary>
    /// Class MapiMessage containing all the members defining a mapi message
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public class MapiMessage
    {
        /// <summary>
        /// reserved
        /// </summary>
        public int reserved;
        /// <summary>
        /// subject
        /// </summary>
        public string subject;
        /// <summary>
        /// noteText
        /// </summary>
        public string noteText;
        /// <summary>
        /// messageType
        /// </summary>
        public string messageType;
        /// <summary>
        /// dateReceived
        /// </summary>
        public string dateReceived;
        /// <summary>
        /// conversationID
        /// </summary>
        public string conversationID;
        /// <summary>
        /// flags
        /// </summary>
        public int flags;
        /// <summary>
        /// originator
        /// </summary>
        public IntPtr originator;
        /// <summary>
        /// recipCount
        /// </summary>
        public int recipCount;
        /// <summary>
        /// recips
        /// </summary>
        public IntPtr recips;
        /// <summary>
        /// fileCount
        /// </summary>
        public int fileCount;
        /// <summary>
        /// files
        /// </summary>
        public IntPtr files;
    }

    /// <summary>
    /// Class MapiMessage containing all the members defining a mapi file
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public class MapiFileDesc
    {
        /// <summary>
        /// reserved
        /// </summary>
        public int reserved;
        /// <summary>
        /// flags
        /// </summary>
        public int flags;
        /// <summary>
        /// position
        /// </summary>
        public int position;
        /// <summary>
        /// path
        /// </summary>
        public string path;
        /// <summary>
        /// name
        /// </summary>
        public string name;
        /// <summary>
        /// type
        /// </summary>
        public IntPtr type;
    }

    /// <summary>
    /// Class MapiMessage containing all the members defining a mapi recipient
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public class MapiRecipDesc
    {
        /// <summary>
        /// reserved
        /// </summary>
        public int reserved;
        /// <summary>
        /// recipClass
        /// </summary>
        public int recipClass;
        /// <summary>
        /// name
        /// </summary>
        public string name;
        /// <summary>
        /// address
        /// </summary>
        public string address;
        /// <summary>
        /// eIDSize
        /// </summary>
        public int eIDSize;
        /// <summary>
        /// entryID
        /// </summary>
        public IntPtr entryID;
    }

    #endregion

}