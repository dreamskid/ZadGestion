using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Resources;

namespace GeneralClasses
{

    /// <summary>
    /// Class ResourcesManager containing all the functions for handling files
    /// </summary>
    public class ResourcesManager
    {

        #region Variables

        /// <summary>
        /// Error handler 
        /// </summary>
        private static Log m_Error_Handler = new Log();

        #endregion

        #region Functions

        /// <summary>
        /// Dictionnary containing all the entries in the ressource file
        /// </summary>
        private static Dictionary<string, string> m_ResourceMap = null;

        /// <summary>
        /// Function which fills the Dictionnary
        /// <param name="_Language">Language chosen for the software</param>
        /// <param name="_File_Handler">Global file handler</param>
        /// <returns>True if the filling has been correctly done, false otherwise</returns>
        /// </summary>
        public bool Fill_Resources(string _Language, Files _File_Handler)
        {
            string fileName = @"Language\Language_" + _Language + ".resx";
            if (_File_Handler.File_Exists(fileName) == false)
            {
                Console.WriteLine("File does not exist - " + fileName);
                return false;
            }
            try
            {
                ResXResourceReader resourceReader = new ResXResourceReader(fileName);
                m_ResourceMap = new Dictionary<string, string>();
                foreach (DictionaryEntry dico in resourceReader)
                {
                    m_ResourceMap.Add(dico.Key.ToString(), dico.Value.ToString());
                }
                resourceReader.Close();
                return true;
            }
            catch (Exception e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                return false;
            }
        }

        /// <summary>
        /// Function which get a ressource from the dictionnary
        /// <param name="_ResourceId">Id of the resource to retrieve</param>
        /// <returns>The string corresponding to the given resource Id in the resource map</returns>
        /// </summary>
        public string Get_Resources(string _ResourceId)
        {
            try
            {
                return m_ResourceMap[_ResourceId];
            }
            catch (KeyNotFoundException e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name + ": " + _ResourceId, e);
                return _ResourceId;
            }
            catch (Exception e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                return _ResourceId;
            }
        }

        /// <summary>
        /// Read the version of the software in the Update.slc file
        /// <returns>The version of the software</returns>
        /// </summary>
        public string Read_Version()
        {
            //Base version
            string version = "v1.0.0.0";
            try
            {
                //Get file
                string exePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\";
                string updateFilePath = exePath + "Resources\\Update.slc";
                string updateFile = new Uri(updateFilePath).LocalPath;

                //Read file
                if (File.Exists(updateFile))
                {
                    var lines = File.ReadLines(updateFile);
                    foreach (var line in lines)
                    {
                        version = line;
                        break;
                    }
                }

                //Return the version
                return version;
            }
            catch (Exception e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                return version;
            }
        }

        #endregion

    }
}
