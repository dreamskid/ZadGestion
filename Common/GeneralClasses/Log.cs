using System;

namespace GeneralClasses
{
    /// <summary>
    /// Class Log containing all the functions for handling error
    /// </summary>
    public class Log
    {

        #region Variables

        /// <summary>
        /// Log file name
        /// </summary>
        private string m_LogFile = "ZadGestion.log";

        #endregion

        #region Getter/Setter

        /// <summary>
        /// Getter/Setter for the error
        /// </summary>
        public string error { get; set; }

        /// <summary>
        /// Getter for the log filename
        /// </summary>
        public string Get_LogFilename()
        {
            return m_LogFile;
        }
        /// <summary>
        /// Setter for the log filename
        /// <param name="_LogFile">Log file name</param>
        /// </summary>
        public void Set_LogFilename(string _LogFileName)
        {
            m_LogFile = _LogFileName;
        }

        #endregion

        #region Functions

        /// <summary>
        /// Write a line in the log file from a primary and a secondary messages
        /// <param name="_Action">Message number 2</param>
        /// </summary>
        public void WriteAction(string _Action)
        {
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter("./ZadGestion.log", true))
            {
                file.WriteLine("[ACTION] " + DateTime.Now.ToString() + " - " + _Action);
            }
        }

        /// <summary>
        /// Write a line in the log file from a primary and a secondary messages
        /// <param name="_MethodName">Name of the method where the message is coming from</param>
        /// <param name="_Message">Message</param>
        /// </summary>
        public void WriteMessage(string _MethodName, string _Message = "")
        {
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter("./ZadGestion.log", true))
            {
                file.WriteLine("[MESSAGE] " + DateTime.Now.ToString() + " - " + _MethodName + "\t" + _Message);
            }
#if DEBUG
            Console.WriteLine(_MethodName + "\t" + _Message);
#endif
        }

        /// <summary>
        /// Write a line in the log file from an exception
        /// <param name="_MethodName">Name of the method where the exception is coming from</param>
        /// <param name="_Exception">Exception</param>
        /// </summary>
        public void WriteException(string _MethodName, Exception _Exception)
        {
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter("./ZadGestion.log", true))
            {
                string exceptionMessage = "[EXCEPTION] " + DateTime.Now.ToString() + " - " + _MethodName + "\t" + _Exception.Message;
                if (_Exception.InnerException != null)
                {
                    exceptionMessage += "\t" + _Exception.InnerException.Message;
                }
                file.WriteLine(exceptionMessage);
            }
#if DEBUG
            Console.WriteLine(_MethodName + "\t" + _Exception.Message + "\t" + _Exception.StackTrace);
#endif
        }

        #endregion

    }
}