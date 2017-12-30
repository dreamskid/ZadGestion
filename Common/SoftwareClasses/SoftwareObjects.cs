using System.Collections.Generic;

namespace SoftwareClasses
{

    /// <summary>
    /// Class containing all the collection of class for the software
    /// </summary>
    public class SoftwareObjects
    {

        #region Variables

        /// <summary>
        /// List of missions
        /// </summary>
        private static List<Mission> m_MissionsCollection = new List<Mission>();

        /// <summary>
        /// List of shifts
        /// </summary>
        private static List<Shift> m_ShiftsCollection = new List<Shift>();

        /// <summary>
        /// List of host and hostesses
        /// </summary>
        private static List<Hostess> m_HostAndHostesssCollection = new List<Hostess>();

        /// <summary>
        /// List of clients
        /// </summary>
        private static List<Client> m_ClientsCollection = new List<Client>();

        /// <summary>
        /// Settings
        /// </summary>
        private static Settings m_Settings = new Settings();

        #endregion

        #region Getter/Setter

        /// <summary>
        /// Getter/Setter for the collection of missions
        /// </summary>
        public static List<Mission> MissionsCollection
        {
            get { return m_MissionsCollection; }
            set { m_MissionsCollection = value; }
        }

        /// <summary>
        /// Getter/Setter for the collection of shifts
        /// </summary>
        public static List<Shift> ShiftsCollection
        {
            get { return m_ShiftsCollection; }
            set { m_ShiftsCollection = value; }
        }

        /// <summary>
        /// Getter/Setter for the collection of hosts and hostesses
        /// </summary>
        public static List<Hostess> HostsAndHotessesCollection
        {
            get { return m_HostAndHostesssCollection; }
            set { m_HostAndHostesssCollection = value; }
        }

        /// <summary>
        /// Getter/Setter for the collection of clients
        /// </summary>
        public static List<Client> ClientsCollection
        {
            get { return m_ClientsCollection; }
            set { m_ClientsCollection = value; }
        }

        /// <summary>
        /// Getter/Setter for the global settings
        /// </summary>
        public static Settings GlobalSettings
        {
            get { return m_Settings; }
            set { m_Settings = value; }
        }

        #endregion

    }
}
