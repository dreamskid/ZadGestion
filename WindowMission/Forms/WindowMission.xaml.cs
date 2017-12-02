using GeneralClasses;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace WindowMission
{

    /// <summary>
    /// Class MainWindow_Services from the namespace Services
    /// </summary>
    public partial class MainWindow_Mission : Window
    {

        #region Initialization

        #region Variables

        /// <summary>
        /// Confirmation of quit the window
        /// </summary>
        private bool m_ConfirmQuit = false;

        /// <summary>
        /// Global handlers for common objects
        /// </summary>
        public static Handlers m_Global_Handler = null;

        /// <summary>
        /// Internet handler
        /// </summary>
        public static Database.Database m_Database_Handler = null;

        /// <summary>
        /// Missions
        /// </summary>
        public Billing m_Mission;

        #endregion

        #region Functions

        /// <summary>
        /// Constructor for the services main window
        /// </summary>
        public MainWindow_Mission(Handlers _Global_Handler, Database.Database _Database_Handler, Billing _Mission)
        {
            try
            {
                //Initialize
                InitializeComponent();
                this.Closing += new System.ComponentModel.CancelEventHandler(Window_Closing);

                //Initialize variables
                m_Global_Handler = _Global_Handler;
                m_Database_Handler = _Database_Handler;
                m_Mission = _Mission;

                //Setting the contents of objects
                if (m_Global_Handler != null)
                {
                    Define_Content();
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Define each content for labels, button, text in the form
        /// </summary>
        private void Define_Content()
        {
            //Header
            this.Title = m_Global_Handler.Resources_Handler.Get_Resources("Mission") + " " + m_Mission.num_Mission;

        }

        #endregion

        #endregion

        #region Event
        
        /// <summary>
        /// Closing event by clicking oon the X
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //Already quit without saving
            if (m_ConfirmQuit == true)
            {
                return;
            }

            //Confirm
            System.Windows.MessageBoxResult result = System.Windows.MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSavingMessage"),
                            m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSavingCaption"),
                            MessageBoxButton.YesNo, MessageBoxImage.Exclamation);

            //Quit
            if (result == MessageBoxResult.Yes)
            {
                this.DialogResult = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

        #endregion
        
    }
}