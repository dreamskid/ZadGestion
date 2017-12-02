using GeneralClasses;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media.Animation;

namespace Launcher
{
    /// <summary>
    /// Class MainWindow from the namespace WindowChoice
    /// </summary>
    public partial class MainWindow : Window
    {

        #region Initialization

        /// <summary>
        /// Panel
        /// </summary>
        public enum Panel
        {
            ZADGESTION = 1,
            HOSTANDHOSTESS = 2,
            SETTINGS = 3
        };

        /// <summary>
        /// Global handlers for common objects
        /// </summary>
        public static Handlers m_Global_Handler = new Handlers();

        /// <summary>
        /// Handler for internet
        /// </summary>
        public static Database.Database m_Database_Handler = new Database.Database(m_Global_Handler);

        /// <summary>
        /// Constructor for the main window
        /// </summary>
        public MainWindow()
        {
            //Initialization of all the components
            InitializeComponent();

            try
            {
                //Get the default language
                System.Windows.Forms.InputLanguage lg = System.Windows.Forms.InputLanguage.CurrentInputLanguage;
                m_Global_Handler.Language_Handler = "fr-FR";
                /*
                if (lg != null)
                {
                    m_Global_Handler.Language_Handler = lg.Culture.Name;
                }
                */

                //Get the resources
                if (m_Global_Handler.Resources_Handler.Fill_Resources(m_Global_Handler.Language_Handler, m_Global_Handler.File_Handler) == false)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("ResourcesNotFound"),
                                m_Global_Handler.Resources_Handler.Get_Resources("ResourcesNotFoundCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                    Close();
                    return;
                }

                //Verification of word and excel
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("WordNotPresentErrorText"), m_Global_Handler.Resources_Handler.Get_Resources("WordNotPresentErrorCaption"));
                    Close();
                    return;
                }
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("ExcelNotPresentErrorText"), m_Global_Handler.Resources_Handler.Get_Resources("ExcelNotPresentErrorCaption"));
                    Close();
                    return;
                }

                //Setting the contents of objects
                Define_Content();

                //Setting the language of the software
                System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(m_Global_Handler.Language_Handler);

                //Show
                this.Show();

                return;
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message,
                               m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                m_Global_Handler.Log_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, exception);
                Close();
                return;
            }
        }

        /// <summary>
        /// Define each content for labels, button, text in the form
        /// </summary>
        private void Define_Content()
        {

            //Buttons
            Btn_Main_Soft.Content = m_Global_Handler.Resources_Handler.Get_Resources("Missions");
            Btn_Main_Settings.Content = m_Global_Handler.Resources_Handler.Get_Resources("Settings");
            Btn_Main_HostAndHostess.Content = m_Global_Handler.Resources_Handler.Get_Resources("HostsAndHostesses");

            //Text box
            string version = m_Global_Handler.Resources_Handler.Read_Version();
            Txt_Version.Text = "v" + version;
        }

        #endregion

        #region Events

        /// <summary>
        /// Click on the button Software
        /// </summary>
        private void Btn_Main_Soft_Click(object sender, RoutedEventArgs e)
        {

            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            //Open the wait window
            windowWait.Start(m_Global_Handler, "OpenSoftPrincipalMessage", "OpenSoftSecondaryMessage");

            //Open the software form
            Software.MainWindow_Software Software = new Software.MainWindow_Software(m_Global_Handler, m_Database_Handler, 1);
            Software.Show();

            //Close the main window form
            Close();

            //Close the wait windows
            windowWait.Stop();
        }

        /// <summary>
        /// Click on the cross button
        /// </summary>
        private void Btn_Main_Close_Click(object sender, RoutedEventArgs e)
        {
            Main_Window_Closing();
        }

        /// <summary>
        /// Click on the button Host and hostess
        /// </summary>
        private void Btn_Main_HostAndHostess_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            //Open the wait window
            windowWait.Start(m_Global_Handler, "OpenSoftPrincipalMessage", "OpenSoftSecondaryMessage");

            //Open the software form
            Software.MainWindow_Software Software = new Software.MainWindow_Software(m_Global_Handler, m_Database_Handler, 2);
            Software.Show();

            //Close the main window form
            Close();

            //Close the wait windows
            windowWait.Stop();
        }

        /// <summary>
        /// Click on the button Settings
        /// </summary>
        private void Btn_Main_Settings_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            //Open the wait window
            windowWait.Start(m_Global_Handler, "OpenSoftPrincipalMessage", "OpenSoftSecondaryMessage");

            //Open the software form
            Software.MainWindow_Software Software = new Software.MainWindow_Software(m_Global_Handler, m_Database_Handler, 3);
            Software.Show();

            //Close the main window form
            Close();

            //Close the wait windows
            windowWait.Stop();
        }

        /// <summary>
        /// Internet connection lost event handler
        /// Need to be better TODO
        /// </summary>
        void NetworkChange_NetworkAvailabilityChanged(object sender, EventArgs e)
        {
            MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("ConnectivityLostText"), m_Global_Handler.Resources_Handler.Get_Resources("ConnectivityLostCaption"));
        }

        #endregion

        #region Animations

        /// <summary>
        /// Boolean indicating if the form is already faded or not
        /// </summary>
        bool AlreadyFaded = false;

        /// <summary>
        /// Animation for the closing of the form
        /// </summary>
        private void Main_Window_Closing()
        {
            if (!AlreadyFaded)
            {
                AlreadyFaded = true;
                DoubleAnimation anim = new DoubleAnimation(0, (Duration)TimeSpan.FromSeconds(0.3));
                anim.Completed += new EventHandler(Main_Window_Closing_Completed);
                this.BeginAnimation(UIElement.OpacityProperty, anim);
            }
        }

        /// <summary>
        /// End of the animation
        /// </summary>
        void Main_Window_Closing_Completed(object sender, EventArgs e)
        {
            Close();
        }

        #endregion

    }
}