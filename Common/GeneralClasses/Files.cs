using System;

namespace GeneralClasses
{

    /// <summary>
    /// Class Files containing all the functions for handling files
    /// </summary>
    public class Files
    {

        #region Variables

        /// <summary>
        /// Error handler 
        /// </summary>
        private static Log m_Error_Handler = new Log();

        #endregion

        #region Functions

        /// <summary>
        /// Test if a file exists
        /// <param name="_FileName">Location of the file</param>
        /// <returns>True if the file exists, false otherwise</returns>
        /// </summary>
        public bool File_Exists(string _FileName)
        {
            if (System.IO.File.Exists(_FileName))
            {
                return true;
            }
            else
            {
                if (_FileName.Contains("file:\\"))
                {
                    _FileName = _FileName.Replace("file:\\", "");
                    return File_Exists(_FileName);
                }
                return false;
            }
        }

        /// <summary>
        /// Delete a file
        /// <param name="_FileName">Location of the file</param>
        /// <returns>True if the operation went well, false otherwise</returns>
        /// </summary>
        public bool Delete_File(string _FileName)
        {
            try
            {
                if (_FileName.Contains("file:\\"))
                {
                    _FileName = _FileName.Replace("file:\\", "");
                }
                System.IO.File.Delete(_FileName);
                return true;
            }
            catch (Exception exception)
            {
                m_Error_Handler.WriteMessage(System.Reflection.MethodBase.GetCurrentMethod().Name, exception.StackTrace);
                return false;
            }
        }

        #endregion

    }
}
