namespace SoftwareClasses
{

    /// <summary>
    /// Class Shift containing all the variables describing a shift
    /// </summary>
    public class Shift
    {
        /// <summary>
        /// Getter/Setter for the date
        /// </summary>
        public string date
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for mission id
        /// </summary>
        public string id
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the hourly rate
        /// </summary>
        public string hourly_rate
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the list of hosts and hostesses id
        /// </summary>
        public string id_list_hostsandhostesses
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the related mission id
        /// </summary>
        public string id_mission
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the start  of shift time
        /// </summary>
        public string start_time
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the end of shift time
        /// </summary>
        public string end_time
        {
            get;
            set;
        }
    }
}
