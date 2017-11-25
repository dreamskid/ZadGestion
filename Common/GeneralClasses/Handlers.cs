namespace GeneralClasses
{

    /// <summary>
    /// Class Handlers containing all the handlers for common classes
    /// </summary>
    public class Handlers
    {

        #region Variables

        /// <summary>
        /// Handler for language
        /// </summary>
        private string m_Language_Handler = "";

        /// <summary>
        /// Handler for files
        /// </summary>
        private Files m_File_Handler = new Files();

        /// <summary>
        /// Handler for error
        /// </summary>
        private Log m_Error_Handler = new Log();

        /// <summary>
        /// Handler for date and times
        /// </summary>
        private DateAndTime m_DateAndTime_Handler = new DateAndTime();

        /// <summary>
        /// Handler for images
        /// </summary>
        private Images m_Image_Handler = new Images();

        /// <summary>
        /// Handler for controls
        /// </summary>
        private Controls m_Controls_Handler = new Controls();

        /// <summary>
        /// Handler for resources
        /// </summary>
        private ResourcesManager m_Resources_Handler = new ResourcesManager();

        /// <summary>
        /// Handler for word generation
        /// </summary>
        private WordGeneration m_Word_Handler = new WordGeneration();

        /// <summary>
        /// Handler for excel generation
        /// </summary>
        private ExcelGeneration m_Excel_Handler = new ExcelGeneration();

        #endregion

        #region Getter/Setter

        /// <summary>
        /// Getter/Setter for language
        /// </summary>
        public string Language_Handler
        {
            get
            {
                return m_Language_Handler;
            }
            set
            {
                m_Language_Handler = value;
            }
        }

        /// <summary>
        /// Getter/Setter for File_Handler
        /// </summary>
        public Files File_Handler
        {
            get
            {
                return m_File_Handler;
            }
            set
            {
                m_File_Handler = value;
            }
        }

        /// <summary>
        /// Getter/Setter for Error_Handler
        /// </summary>
        public Log Error_Handler
        {
            get
            {
                return m_Error_Handler;
            }
            set
            {
                m_Error_Handler = value;
            }
        }

        /// <summary>
        /// Getter/Setter for DateAndTime_Handler
        /// </summary>
        public DateAndTime DateAndTime_Handler
        {
            get
            {
                return m_DateAndTime_Handler;
            }
            set
            {
                m_DateAndTime_Handler = value;
            }
        }

        /// <summary>
        /// Getter/Setter for Image_Handler
        /// </summary>
        public Images Image_Handler
        {
            get
            {
                return m_Image_Handler;
            }
            set
            {
                m_Image_Handler = value;
            }
        }

        /// <summary>
        /// Getter/Setter for Controls_Handler
        /// </summary>
        public Controls Controls_Handler
        {
            get
            {
                return m_Controls_Handler;
            }
            set
            {
                m_Controls_Handler = value;
            }
        }

        /// <summary>
        /// Getter/Setter for Resources_Handler
        /// </summary>
        public ResourcesManager Resources_Handler
        {
            get
            {
                return m_Resources_Handler;
            }
            set
            {
                m_Resources_Handler = value;
            }
        }

        /// <summary>
        /// Getter/Setter for m_Word_Handler
        /// </summary>
        public WordGeneration Word_Handler
        {
            get
            {
                return m_Word_Handler;
            }
            set
            {
                m_Word_Handler = value;
            }
        }

        /// <summary>
        /// Getter/Setter for m_Excel_Handler
        /// </summary>
        public ExcelGeneration Excel_Handler
        {
            get
            {
                return m_Excel_Handler;
            }
            set
            {
                m_Excel_Handler = value;
            }
        }

        #endregion

    }
}