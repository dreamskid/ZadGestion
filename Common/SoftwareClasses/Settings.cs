namespace SoftwareClasses
{
    /// <summary>
    /// Class Settings containing all the variables corresponding to the settings choosen by the user
    /// </summary>
    public class Settings
    {

        /// <summary>
        /// Getter/Setter for the database definition
        /// </summary>
        public string database_definition
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the hourly rate definition
        /// </summary>
        public string id
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the photos repository
        /// </summary>
        public string photos_repository
        {
            get;
            set;
        }

        #region Shift

        /// <summary>
        /// Getter/Setter for the hourly rate definition
        /// </summary>
        public string hourly_rate
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the pause duration definition
        /// </summary>
        public string pause_duration
        {
            get;
            set;
        }

        #endregion
    }
}
