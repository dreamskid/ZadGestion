using GeneralClasses;
using System;
using System.Reflection;
using System.Windows;

namespace Software
{
    /// <summary>
    /// Interaction logic for WindowDatePicker.xaml
    /// </summary>
    public partial class WindowDatePicker : Window
    {
        #region Initialization

        #region Variables

        /// <summary>
        /// Global handlers for common objects
        /// </summary>
        public static Handlers m_Global_Handler = null;

        /// <summary>
        /// First date selected
        /// </summary>
        public DateTime m_FirstSelectedDate;

        /// <summary>
        /// End date selected
        /// </summary>
        public DateTime m_EndSelectedDate;

        #endregion

        #region Functions

        /// <summary>
        /// Constructor for the date picker main window
        /// </summary>
        public WindowDatePicker(GeneralClasses.Handlers _GlobalHandler, string _FirstDateExplanation, string _EndDateExplanation)
        {
            try
            {
                //Initialization of all the components
                InitializeComponent();

                //Setting the variables
                m_Global_Handler = _GlobalHandler;

                //No need of a define content function here
                this.Title = m_Global_Handler.Resources_Handler.Get_Resources("PickDate");
                Lbl_FirstDateExplanation.Content = _FirstDateExplanation;
                Lbl_EndDateExplanation.Content = _EndDateExplanation;

            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        #endregion

        #endregion

        #region 

        /// <summary>
        /// Save dates and quit on selected date event in the Dtp_EndDatePicker
        /// </summary>
        private void Dtp_EndDatePicker_CalendarClosed(object sender, RoutedEventArgs e)
        {
            Save_Dates();
        }


        /// <summary>
        /// Save dates and quit on key Enter event in the Dtp_EndDatePicker
        /// </summary>
        private void Dtp_EndDatePicker_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                Save_Dates();
            }
        }

        #endregion

        #region Functions

        private void Save_Dates()
        {
            try
            {
                //Set the selected date
                m_FirstSelectedDate = Convert.ToDateTime(Dtp_FirstDatePicker.SelectedDate);
                m_EndSelectedDate = Convert.ToDateTime(Dtp_EndDatePicker.SelectedDate);

                if (m_FirstSelectedDate != DateTime.MinValue && m_EndSelectedDate != DateTime.MinValue)
                {
                    //Close the window
                    this.DialogResult = true;
                    this.Close();
                }

                return;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                this.DialogResult = false;
                return;
            }
        }

        #endregion
    }
}
