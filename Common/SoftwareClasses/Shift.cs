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
        /// Getter/Setter for the time of the end of shift
        /// </summary>
        public string end_time
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
        /// Getter/Setter for the shift id
        /// </summary>
        public string id
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the host or hostess id
        /// </summary>
        public string id_hostorhostess
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
        /// Getter/Setter for the pause
        /// </summary>
        public string pause
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the start time of the shift
        /// </summary>
        public string start_time
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for suit loan
        /// </summary>
        public bool suit
        {
            get;
            set;
        }
    }
}
