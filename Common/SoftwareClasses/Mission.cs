namespace SoftwareClasses
{
    /// <summary>
    /// Mission status
    /// </summary>
    public enum MissionStatus
    {
        NONE = 0,
        CREATED = 1,
        DONE = 2,
        DECLINED = 3,
        BILLED = 4
    };

    /// <summary>
    /// Class Mission containing all the variables describing a mission
    /// </summary>
    public class Mission
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public Mission()
        {
        }

        /// <summary>
        /// Constructor by copy
        /// </summary>
        public Mission(Mission _MissionToCopy)
        {
            address = _MissionToCopy.address;
            city = _MissionToCopy.city;
            client_name = _MissionToCopy.client_name;
            country = _MissionToCopy.country;
            end_date = _MissionToCopy.end_date;
            start_date = _MissionToCopy.start_date;
            state = _MissionToCopy.state;
            zipcode = _MissionToCopy.zipcode;
        }

        /// <summary>
        /// Getter/Setter for client's adress
        /// </summary>
        public string address
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for client's archived mode
        /// </summary>
        public int archived
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the city
        /// </summary>
        public string city
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for client id
        /// </summary>
        public string client_name
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the country
        /// </summary>
        public string country
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the creation date
        /// </summary>
        public string date_creation
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the end date
        /// </summary>
        public string end_date
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
        /// Getter/Setter for the list of shifts id
        /// </summary>
        public string id_list_shifts
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the start date
        /// </summary>
        public string start_date
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the state
        /// </summary>
        public string state
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for the zipcode
        /// </summary>
        public string zipcode
        {
            get;
            set;
        }

        /// <summary>
        /// Functions
        /// Creation of the Id of the mission
        /// </summary>
        public string Create_MissionId()
        {
            string id = "";

            id += SoftwareObjects.MissionsCollection.Count;
            id += "_";
            if (client_name != "")
            {
                id += client_name.Substring(0, 1);
            }
            if (date_creation != "")
            {
                id += date_creation;
            }
            if (city != "")
            {
                id += "_" + city;
            }

            return id;
        }

    }
}
