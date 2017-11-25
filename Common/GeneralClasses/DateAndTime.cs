using System;

namespace GeneralClasses
{

    /// <summary>
    /// Class DateTime containing all the functions for handling date and time
    /// </summary>
    public class DateAndTime
    {

        #region Variables

        /// <summary>
        /// Error handler 
        /// </summary>
        private static Log m_Error_Handler = new Log();

        #endregion

        #region Functions

        /// <summary>
        /// Get the present year
        /// <returns>string containing the present year</returns>
        /// </summary>
        public string Get_PresentYear()
        {
            string res = DateTime.Now.Year.ToString();
            return res;
        }

        /// <summary>
        /// Get the present year in the short format
        /// <returns>string containing the present year in the short format</returns>
        /// </summary>
        public string Get_PresentYearShort()
        {
            DateTime date = DateTime.Now;
            string res = date.ToString("yy");
            return res;
        }

        /// <summary>
        /// Get the present month
        /// <returns>string containing the present month</returns>
        /// </summary>
        public string Get_PresentMonth()
        {
            string res = DateTime.Now.Month.ToString();
            if (res.Length == 1)
                res = "0" + res;
            return res;
        }

        /// <summary>
        /// Get the present day
        /// <returns>string containing the present day</returns>
        /// </summary>
        public string Get_PresentDay()
        {
            string res = DateTime.Now.Day.ToString();
            if (res.Length == 1)
                res = "0" + res;
            return res;
        }

        /// <summary>
        /// Get the present hour
        /// <returns>string containing the present hour</returns>
        /// </summary>
        public string Get_PresentHour()
        {
            string res = DateTime.Now.Hour.ToString();
            if (res.Length == 1)
                res = "0" + res;
            return res;
        }

        /// <summary>
        /// Transform the date in function of the culture in order to be shown
        /// <param name="_Date">Date to treat</param>
        /// <param name="_Language">Language</param>
        /// <returns>string containing the treated date</returns>
        /// </summary>
        public string Treat_Date(string _Date, string _Language)
        {
            try
            {
                if (_Date == null)
                {
                    return "";
                }
                else if (_Date.Contains("0000-00-00"))
                {
                    _Date = null;
                    return "";
                }
                else if (_Language == null)
                {
                    Exception e = new Exception("No language");
                    throw e;
                }
                else if (_Date == "" || _Language == "")
                {
                    return "";
                }
                IFormatProvider culture = new System.Globalization.CultureInfo(_Language, true);
                DateTime dateTime = DateTime.Parse(_Date, culture, System.Globalization.DateTimeStyles.AssumeLocal);
                return dateTime.ToShortDateString();
            }
            catch (Exception e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                return "";
            }
        }

        #endregion

    }
}
