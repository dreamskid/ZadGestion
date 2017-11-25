using GeneralClasses;
using MySql.Data.MySqlClient;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Xml;

namespace Software
{

    /// <summary>
    /// Class MainWindow_Software from the namespace Software
    /// </summary>
    public partial class MainWindow_Software : Window
    {

        #region Initialization

        #region Variables

        #region Colors

        /// <summary>
        /// Initialization
        /// Color
        /// Background color for archived mission
        /// </summary>
        Brush m_Color_ArchivedMission = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Background color for hostess
        /// </summary>
        Brush m_Color_HostAndHostess = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Background color for mission
        /// </summary>
        Brush m_Color_Mission = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Color green
        /// </summary>
        Brush m_Color_Button = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Color for main button
        /// </summary>
        public Brush m_Color_MainButton = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Color for button on mouse over event
        /// </summary>
        public Brush m_Color_MouseOver = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Color red
        /// </summary>
        Brush m_Color_Red = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Background color for selected archived mission
        /// </summary>
        Brush m_Color_SelectedArchivedMission = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Background color for selected hostess
        /// </summary>
        Brush m_Color_SelectedHostAndHostess = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Background color for selected mission
        /// </summary>
        Brush m_Color_SelectedMission = null;

        /// <summary>
        /// Initialization
        /// Color
        /// Color for selected main button
        /// </summary>
        Brush m_Color_SelectedMainButton = null;

        #endregion

        #region Datagrids variables

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Class describing a service contained in the missions category datagrid
        /// </summary>
        public class m_Datagrid_Missions_Services
        {
            /// <summary>
            /// Constructor
            /// </summary>
            public m_Datagrid_Missions_Services(string _reference, string _description, string _quantity)
            {
                reference = _reference;
                description = _description;
                quantity = _quantity;
            }

            /// <summary>
            /// Reference
            /// </summary>
            public string reference { set; get; }
            /// <summary>
            /// Description
            /// </summary>
            public string description { set; get; }
            /// <summary>
            /// Amount
            /// </summary>
            public string quantity { set; get; }
        }

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection of services contained in the missions category datagrid
        /// </summary>
        ObservableCollection<m_Datagrid_Missions_Services> m_DataGrid_Missions_ServicesCollection =
            new ObservableCollection<m_Datagrid_Missions_Services>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Class describing a mission contained in the host and hostess category datagrid
        /// </summary>
        public class m_DataGrid_HostAndHostess_Missions
        {
            /// <summary>
            /// Constructor
            /// </summary>
            public m_DataGrid_HostAndHostess_Missions(string _Id, string _Client, string _City, string _Date)
            {
                id = _Id;
                client = _Client;
                city = _City;
                date = _Date;
            }

            /// <summary>
            /// Id
            /// </summary>
            public string id { set; get; }
            /// <summary>
            /// Client
            /// </summary>
            public string client { set; get; }
            /// <summary>
            /// City
            /// </summary>
            public string city { set; get; }
            /// <summary>
            /// Date
            /// </summary>
            public string date { set; get; }
        }

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection of missions done by the selected host or hostess
        /// </summary>
        ObservableCollection<m_DataGrid_HostAndHostess_Missions> m_DataGrid_HostAndHostess_MissionsCollection =
            new ObservableCollection<m_DataGrid_HostAndHostess_Missions>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Class describing a mission contained in the clients category datagrid
        /// </summary>
        public class m_DataGrid_Clients_Missions
        {
            /// <summary>
            /// Constructor
            /// </summary>
            public m_DataGrid_Clients_Missions(string _Id, string _Description, string _City, string _Date)
            {
                id = _Id;
                description = _Description;
                city = _City;
                date = _Date;
            }

            /// <summary>
            /// Id
            /// </summary>
            public string id { set; get; }
            /// <summary>
            /// Name
            /// </summary>
            public string description { set; get; }
            /// <summary>
            /// City
            /// </summary>
            public string city { set; get; }
            /// <summary>
            /// Date
            /// </summary>
            public string date { set; get; }
        }

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection of missions done by the selected host or hostess
        /// </summary>
        ObservableCollection<m_DataGrid_Clients_Missions> m_DataGrid_Clients_MissionsCollection =
            new ObservableCollection<m_DataGrid_Clients_Missions>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection of missions panels contained in the missions grid
        /// </summary>
        List<Billing> m_Grid_Details_Missions_MissionsCollection = new List<Billing>();

        #endregion

        #region Global variables

        /// <summary>
        /// Initialization
        /// Global variable
        /// Xml document containing all the cities in France with zipcode
        /// </summary>
        XmlDocument m_Cities_DocFrance = new XmlDocument();

        /// <summary>
        /// Initialization
        /// Global variable
        /// Global handlers for common objects
        /// </summary>
        public static Handlers m_Global_Handler = null;

        /// <summary>
        /// Initialization
        /// Global variable
        /// Database handler
        /// </summary>
        public static Database.Database m_Database_Handler = null;

        /// <summary>
        /// Initialization
        /// Global variable
        /// Boolean indicating if the initialization of the window is done
        /// </summary>
        public bool m_Is_InitializationDone = false;

        /// <summary>
        /// Initialization
        /// Global variable
        /// Boolean indicating if the invoices should be corrected (rest amount and dates)
        /// </summary>
        public bool m_VerifyInvoices = false;

        #endregion

        #region Selected buttons

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected hostess cellphone button
        /// </summary>
        public Button m_Button_HostAndHostess_SelectedCellPhone = null;

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected hostess button
        /// </summary>
        public Button m_Button_HostAndHostess_SelectedHostAndHostess = null;

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected hostess email button
        /// </summary>
        public Button m_Button_HostAndHostess_SelectedEmail = null;

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected mission email button
        /// </summary>
        public Button m_Button_Mission_SelectedEmail = null;

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected mission button
        /// </summary>
        public Button m_Button_Mission_SelectedMission = null;

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected invoice button
        /// </summary>
        public Button m_Button_Mission_SelectedInvoice = null;

        #endregion

        #region Specific variables

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Missions next number
        /// </summary>
        private string m_MissionNextNumber = "";

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Boolean indicating if the mission category is in archive mode or not
        /// </summary>
        private bool m_Mission_IsArchiveMode = false;

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Selected status for missions
        /// </summary>
        private MissionStatus m_Mission_SelectedStatus = MissionStatus.NONE;

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Selected hostess id
        /// </summary>
        public string m_Id_SelectedHostAndHostess = "-1";

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Boolean indicating the sorting way of the hosts and hostesses
        /// </summary>
        bool m_IsSortHostess_Ascending = true;

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Boolean indicating the sorting way of the missions
        /// </summary>
        bool m_IsSortMission_Ascending = true;

        #endregion

        #endregion

        #region Functions

        /// <summary>
        /// Initialization
        /// Function
        /// Constructor for the software main window
        /// </summary>
        public MainWindow_Software(Handlers _Global_Handler, Database.Database _Database_Handler, int _Panel)
        {
            try
            {
                //Initialization of all the components
                InitializeComponent();

                //Setting the variables
                m_Global_Handler = _Global_Handler;
                m_Database_Handler = _Database_Handler;

                //New session
                m_Global_Handler.Error_Handler.WriteAction("New session");

                //Colors
                BrushConverter converter = new BrushConverter();
                Brush panelColor = (Brush)converter.ConvertFromString("#FFFFFFFF");
                m_Color_MainButton = (Brush)converter.ConvertFromString("#FF3B3839");
                m_Color_SelectedMainButton = (Brush)converter.ConvertFromString("#FF595C61");
                m_Color_Button = (Brush)converter.ConvertFromString("#FF6F0178");
                m_Color_Red = (Brush)converter.ConvertFromString("#FFF03535");
                m_Color_HostAndHostess = panelColor;
                m_Color_SelectedHostAndHostess = (Brush)converter.ConvertFromString("#FF595C61");
                m_Color_ArchivedMission = (Brush)converter.ConvertFromString("#FF595C61");
                m_Color_SelectedArchivedMission = (Brush)converter.ConvertFromString("#FF66A2CD");
                m_Color_Mission = panelColor;
                m_Color_SelectedMission = (Brush)converter.ConvertFromString("#FF595C61");

                //Graphism
                m_Color_MainButton = Btn_Software_Missions.Background;
                if (_Panel == 1)
                {
                    Btn_Software_Missions.Background = m_Color_SelectedMainButton;
                }
                else if (_Panel == 2)
                {
                    //Graphism
                    Btn_Software_Clients.Background = m_Color_MainButton;
                    Btn_Software_HostAndHostess.Background = m_Color_SelectedMainButton;
                    Btn_Software_Missions.Background = m_Color_MainButton;
                    Btn_Software_Settings.Background = m_Color_MainButton;

                    //Visibility
                    Grid_Clients.Visibility = Visibility.Collapsed;
                    Grid_HostAndHostess.Visibility = Visibility.Visible;
                    Grid_Missions.Visibility = Visibility.Collapsed;
                    Grid_Settings.Visibility = Visibility.Collapsed;
                }
                else if (_Panel == 3)
                {
                    //Graphism
                    Btn_Software_Clients.Background = m_Color_MainButton;
                    Btn_Software_HostAndHostess.Background = m_Color_MainButton;
                    Btn_Software_Missions.Background = m_Color_MainButton;
                    Btn_Software_Settings.Background = m_Color_SelectedMainButton;

                    //Visibility
                    Grid_Clients.Visibility = Visibility.Collapsed;
                    Grid_HostAndHostess.Visibility = Visibility.Collapsed;
                    Grid_Missions.Visibility = Visibility.Collapsed;
                    Grid_Settings.Visibility = Visibility.Visible;
                }

                //Fill the cities
                string fileFrance = ".\\Cities\\France.xml";
                if (m_Global_Handler.File_Handler.File_Exists(fileFrance))
                {
                    m_Cities_DocFrance.Load(fileFrance);
                }

                //Actualize settings from the database
                Actualize_SettingsFromDatabase();

                //Actualize the hosts and hostesses collection from the database
                Actualize_GridHostsAndHostessesFromDatabase();

                //Actualize the missions collection
                Actualize_GridMissionsFromDatabase();

                //Setting the contents of objects
                if (m_Global_Handler != null)
                {
                    Define_Content();
                }
            }
            catch (Exception e)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return;
            }
        }

        /// <summary>
        /// Initialization
        /// Function
        /// Define each content for labels, button, text in the window
        /// </summary>
        private void Define_Content()
        {
            try
            {
                //Title
                string version = m_Global_Handler.Resources_Handler.Read_Version();
                this.Title = "Zad Gestion - v" + version;

                //Buttons
                Btn_Software_Clients.Content = m_Global_Handler.Resources_Handler.Get_Resources("Clients");
                TextBlock hostessCustomerService = new TextBlock();
                hostessCustomerService.Text = m_Global_Handler.Resources_Handler.Get_Resources("ContactCustomerService");
                hostessCustomerService.TextAlignment = TextAlignment.Center;
                hostessCustomerService.TextWrapping = TextWrapping.Wrap;
                Btn_Software_CustomerService.Content = hostessCustomerService;
                TextBlock hostandhostesses = new TextBlock();
                hostandhostesses.Text = m_Global_Handler.Resources_Handler.Get_Resources("HostsAndHostesses");
                hostandhostesses.TextAlignment = TextAlignment.Center;
                hostandhostesses.TextWrapping = TextWrapping.Wrap;
                Btn_Software_HostAndHostess.Content = hostandhostesses;
                Btn_Software_Missions.Content = m_Global_Handler.Resources_Handler.Get_Resources("Missions");
                Btn_Software_Settings.Content = m_Global_Handler.Resources_Handler.Get_Resources("Settings");

                #region Host and hostess

                //Buttons
                Btn_HostAndHostess_Create.Content = m_Global_Handler.Resources_Handler.Get_Resources("Create");
                Btn_HostAndHostess_Delete.Content = m_Global_Handler.Resources_Handler.Get_Resources("Delete");
                Btn_HostAndHostess_Edit.Content = m_Global_Handler.Resources_Handler.Get_Resources("Edit");
                Btn_HostAndHostess_GenerateExcelStatement.Content = m_Global_Handler.Resources_Handler.Get_Resources("GenerateExcelStatement");

                //Combobox
                List<string> cmbHostess = new List<string>();
                cmbHostess.Add(m_Global_Handler.Resources_Handler.Get_Resources("LastNameHostess"));
                cmbHostess.Add(m_Global_Handler.Resources_Handler.Get_Resources("City"));
                cmbHostess.Add(m_Global_Handler.Resources_Handler.Get_Resources("Sex"));
                cmbHostess.Add(m_Global_Handler.Resources_Handler.Get_Resources("ZipCode"));
                cmbHostess.Add(m_Global_Handler.Resources_Handler.Get_Resources("CreationDate"));
                cmbHostess.Sort();
                foreach (string cmbHostesstr in cmbHostess)
                {
                    Cmb_HostAndHostess_SortBy.Items.Add(cmbHostesstr);
                }

                //Datagrid
                DataGrid_HostAndHostess_Missions.ItemsSource = m_DataGrid_HostAndHostess_MissionsCollection;

                //Labels
                Lbl_HostAndHostess_Research.Content = m_Global_Handler.Resources_Handler.Get_Resources("ResearchHostOrHostess");
                Lbl_HostAndHostess_Sort.Content = m_Global_Handler.Resources_Handler.Get_Resources("SortBy");

                #endregion

                #region Clients

                //Buttons
                Btn_Clients_Create.Content = m_Global_Handler.Resources_Handler.Get_Resources("Create");
                Btn_Clients_Delete.Content = m_Global_Handler.Resources_Handler.Get_Resources("Delete");
                Btn_Clients_Edit.Content = m_Global_Handler.Resources_Handler.Get_Resources("Edit");
                Btn_Clients_GenerateExcelStatement.Content = m_Global_Handler.Resources_Handler.Get_Resources("GenerateExcelStatement");

                //Combobox
                List<string> cmbClients = new List<string>();
                cmbClients.Add(m_Global_Handler.Resources_Handler.Get_Resources("Name"));
                cmbClients.Add(m_Global_Handler.Resources_Handler.Get_Resources("City"));
                cmbClients.Add(m_Global_Handler.Resources_Handler.Get_Resources("ZipCode"));
                cmbClients.Add(m_Global_Handler.Resources_Handler.Get_Resources("CreationDate"));
                cmbClients.Sort();
                foreach (string cmbClientsStr in cmbClients)
                {
                    Cmb_Clients_SortBy.Items.Add(cmbClientsStr);
                }

                //Datagrid
                DataGrid_Clients_Missions.ItemsSource = m_DataGrid_Clients_MissionsCollection;

                //Labels
                Lbl_Clients_Research.Content = m_Global_Handler.Resources_Handler.Get_Resources("ResearchHostOrHostess");
                Lbl_Clients_Sort.Content = m_Global_Handler.Resources_Handler.Get_Resources("SortBy");

                #endregion

                #region Missions

                //Buttons
                Btn_Missions_Close.Content = m_Global_Handler.Resources_Handler.Get_Resources("CloseMission");
                Btn_Missions_Create.Content = m_Global_Handler.Resources_Handler.Get_Resources("CreateMission");
                Btn_Missions_Delete.Content = m_Global_Handler.Resources_Handler.Get_Resources("DeleteMission");
                Btn_Missions_Duplicate.Content = m_Global_Handler.Resources_Handler.Get_Resources("DuplicateMission");
                Btn_Missions_Edit.Content = m_Global_Handler.Resources_Handler.Get_Resources("EditMission");
                Btn_Missions_ShowClosed.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionsShowArchives");
                Btn_Missions_ShowInProgress.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionsShowInProgress");
                Btn_Missions_GenerateExcelStatement.Content = m_Global_Handler.Resources_Handler.Get_Resources("GenerateExcelStatement");

                //Combo boxes
                List<string> cmbMission = new List<string>();
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("Amount"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("ClientCompanyName"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("City"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("CreationDate"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("Missionsubject"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("MissionNumber"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("LastNameHostess"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("Status"));
                cmbMission.Sort();
                foreach (string cmbMissionstr in cmbMission)
                {
                    Cmb_Missions_SortBy.Items.Add(cmbMissionstr);
                }

                //Datagrid
                DataGrid_Missions_Services.ItemsSource = m_DataGrid_Missions_ServicesCollection;

                //Labels
                Lbl_Missions_CreationDate.Content = m_Global_Handler.Resources_Handler.Get_Resources("CreationDate");
                Lbl_Missions_Research.Content = m_Global_Handler.Resources_Handler.Get_Resources("ResearchMission");
                Lbl_Missions_Sort.Content = m_Global_Handler.Resources_Handler.Get_Resources("SortBy");
                Lbl_Missions_Image_Mission_Done.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionLegendDone");
                Lbl_Missions_Image_Mission_Billed.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionLegendBilled");
                Lbl_Missions_Image_Mission_Created.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionLegendCreated");
                Lbl_Missions_Image_Mission_Declined.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionLegendDeclined");

                #endregion

                #region Settings

                //Tab item
                TbItem_Settings_General.Header = m_Global_Handler.Resources_Handler.Get_Resources("Generalities");

                //Buttons
                Btn_Settings_General_Database_Save.Content = m_Global_Handler.Resources_Handler.Get_Resources("Save");
                Btn_Settings_General_Photos_Choose.Content = m_Global_Handler.Resources_Handler.Get_Resources("Choose");

                //Labels
                Lbl_Settings_General_Database.Content = m_Global_Handler.Resources_Handler.Get_Resources("DatabaseDefinition");
                Lbl_Settings_General_Photos.Content = m_Global_Handler.Resources_Handler.Get_Resources("PhotosDirectory");

                #endregion

                //Initialization done
                m_Is_InitializationDone = true;
            }
            catch (Exception e)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return;
            }
        }

        #endregion

        #endregion

        #region Events

        #region Window

        /// <summary>
        /// Event
        /// Window
        /// Animation for the window closing event
        /// </summary>
        private void Software_Window_Closing_Event(object sender, CancelEventArgs e)
        {
            try
            {
                m_Global_Handler.Error_Handler.WriteAction("End session");
                Closing -= Software_Window_Closing_Event;
                e.Cancel = true;
                var anim = new DoubleAnimation(0, (Duration)TimeSpan.FromSeconds(0.3));
                anim.Completed += (s, _) => this.Close();
                this.BeginAnimation(UIElement.OpacityProperty, anim);
                Environment.Exit(0);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                Close();
            }
        }

        #endregion

        #region Main buttons

        /// <summary>
        /// Event - Main buttons - Click on the Clients button
        /// </summary>
        private void Btn_Software_Clients_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Graphism
                Btn_Software_Clients.Background = m_Color_SelectedMainButton;
                Btn_Software_HostAndHostess.Background = m_Color_MainButton;
                Btn_Software_Missions.Background = m_Color_MainButton;
                Btn_Software_Settings.Background = m_Color_MainButton;

                //Visibility
                Grid_Clients.Visibility = Visibility.Visible;
                Grid_HostAndHostess.Visibility = Visibility.Collapsed;
                Grid_Missions.Visibility = Visibility.Collapsed;
                Grid_Settings.Visibility = Visibility.Collapsed;

            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                Close();
            }
        }

        /// <summary>
        /// Event - Main buttons - Click on the hostess the customer service button
        /// </summary>
        private void Btn_Software_CustomerService_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Action
                m_Global_Handler.Error_Handler.WriteAction("Contact customer service");

                //Open the wait window
                windowWait.Start(m_Global_Handler, "HostessCustomerServicePrincipalMessage", "HostessCustomerServiceSecondaryMessage");

                //Send the email with the log attached
                string exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\";
                exePath = exePath.Replace("file:\\", "");
                string file = exePath + "ZadGestion.log";
                string[] attachments = new string[] { file };
                if (!File.Exists(file))
                {
                    attachments = new string[] { };
                }
                string[] recipients = new string[] { "yohann.tschudi@gmail.com" };
                new Email().Compose_Mail(recipients, "", "", attachments);

                //Close window wait
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close window wait
                windowWait.Stop();

                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event - Main buttons - Click on the missions button
        /// </summary>
        private void Btn_Software_HostAndHostess_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Graphism
                Btn_Software_Clients.Background = m_Color_MainButton;
                Btn_Software_HostAndHostess.Background = m_Color_SelectedMainButton;
                Btn_Software_Missions.Background = m_Color_MainButton;
                Btn_Software_Settings.Background = m_Color_MainButton;

                //Visibility
                Grid_Clients.Visibility = Visibility.Collapsed;
                Grid_HostAndHostess.Visibility = Visibility.Visible;
                Grid_Missions.Visibility = Visibility.Collapsed;
                Grid_Settings.Visibility = Visibility.Collapsed;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                Close();
            }
        }

        /// <summary>
        /// Event - Main buttons - Click on the Hostess button
        /// </summary>
        private void Btn_Software_Missions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Graphism
                Btn_Software_Clients.Background = m_Color_MainButton;
                Btn_Software_HostAndHostess.Background = m_Color_MainButton;
                Btn_Software_Missions.Background = m_Color_SelectedMainButton;
                Btn_Software_Settings.Background = m_Color_MainButton;

                //Visibility
                Grid_Clients.Visibility = Visibility.Collapsed;
                Grid_HostAndHostess.Visibility = Visibility.Collapsed;
                Grid_Missions.Visibility = Visibility.Visible;
                Grid_Settings.Visibility = Visibility.Collapsed;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                Close();
            }
        }

        /// <summary>
        /// Event - Main buttons - Click on the settings button
        /// </summary>
        private void Btn_Software_Settings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Graphism
                Btn_Software_Clients.Background = m_Color_MainButton;
                Btn_Software_HostAndHostess.Background = m_Color_MainButton;
                Btn_Software_Missions.Background = m_Color_MainButton;
                Btn_Software_Settings.Background = m_Color_SelectedMainButton;

                //Visibility
                Grid_Clients.Visibility = Visibility.Collapsed;
                Grid_HostAndHostess.Visibility = Visibility.Collapsed;
                Grid_Missions.Visibility = Visibility.Collapsed;
                Grid_Settings.Visibility = Visibility.Visible;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                Close();
            }
        }

        #endregion

        #region Missions

        /// <summary>
        /// Event
        /// Missions
        /// Click on archive/restore button
        /// </summary>
        private void Btn_Missions_Close_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the mission
                Billing missionSel = Get_SelectedMissionFromButton();
                if (missionSel == null)
                {
                    return;
                }

                //In progress mode
                if (m_Mission_IsArchiveMode == false)
                {
                    //Open the window wait
                    windowWait.Start(m_Global_Handler, "MissionArchivePrincipalMessage", "MissionArchiveSecondaryMessage");

                    //Add to database    
                    string res = m_Database_Handler.Close_Mission(missionSel.id);

                    //Treat the result
                    if (res.Contains("result"))
                    {
                        //Action
                        m_Global_Handler.Error_Handler.WriteAction("Mission " + missionSel.id + " closed");

                        //Actualize grid from collection
                        Actualize_GridMissionsFromDatabase();
                        Filter_GridMissionsFromMissionsCollection(m_Mission_SelectedStatus);

                        //Mission archive date
                        string format = "yyyy-MM-dd";
                        missionSel.date_Mission_archived = DateTime.Today.ToString(format);

                        //Clear the fields
                        Txt_Missions_Research.Text = "";
                        Txt_Missions_FirstName.Text = "";
                        Txt_Missions_LastName.Text = "";
                        Txt_Missions_ClientCompanyName.Text = "";
                        Txt_Missions_CreationDate.Text = "";
                        m_DataGrid_Missions_ServicesCollection.Clear();
                        DataGrid_Missions_Services.Items.Refresh();

                        //Null buttons
                        m_Button_Mission_SelectedMission = null;
                        m_Button_Mission_SelectedEmail = null;
                        m_Button_Mission_SelectedInvoice = null;

                        //Disable the buttons
                        Btn_Missions_Duplicate.IsEnabled = false;
                        Btn_Missions_Edit.IsEnabled = false;
                        Btn_Missions_Delete.IsEnabled = false;
                        Btn_Missions_Close.IsEnabled = false;

                        //Close the wait window
                        windowWait.Stop();

                        return;
                    }
                    else if (res.Contains("error"))
                    {
                        //Close the wait window
                        windowWait.Stop();

                        //Treatment of the error
                        Log ClassError = m_Database_Handler.Deserialize_JSON<Log>(res);
                        string errorText = ClassError.error;
                        MessageBox.Show(this, errorText, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                        return;
                    }
                    else
                    {
                        //Close the wait window
                        windowWait.Stop();

                        //Error connecting to web site
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                        return;
                    }
                }
                //In restore mode
                else
                {
                    //Open the window wait
                    windowWait.Start(m_Global_Handler, "MissionRestorePrincipalMessage", "MissionRestoreSecondaryMessage");

                    //Add to database    
                    string res = m_Database_Handler.Reopen_Mission(missionSel.id);

                    //Treat the result
                    if (res.Contains("result"))
                    {
                        //Action
                        m_Global_Handler.Error_Handler.WriteAction("Mission " + missionSel.id + " restored");

                        //Actualize grid from collection
                        Actualize_GridMissionsFromDatabase();
                        m_Mission_SelectedStatus = MissionStatus.NONE;

                        //Mission archive date
                        missionSel.date_Mission_archived = null;

                        //Clear the fields
                        Txt_Missions_Research.Text = "";
                        Txt_Missions_FirstName.Text = "";
                        Txt_Missions_LastName.Text = "";
                        Txt_Missions_ClientCompanyName.Text = "";
                        Txt_Missions_CreationDate.Text = "";
                        m_DataGrid_Missions_ServicesCollection.Clear();
                        DataGrid_Missions_Services.Items.Refresh();

                        //Null buttons
                        m_Button_Mission_SelectedMission = null;
                        m_Button_Mission_SelectedEmail = null;
                        m_Button_Mission_SelectedInvoice = null;

                        //Disable the buttons
                        Btn_Missions_Duplicate.IsEnabled = false;
                        Btn_Missions_Edit.IsEnabled = false;
                        Btn_Missions_Delete.IsEnabled = false;
                        Btn_Missions_Close.IsEnabled = false;

                        //Close the wait window
                        windowWait.Stop();

                        return;
                    }
                    else if (res.Contains("error"))
                    {
                        //Close the wait window
                        windowWait.Stop();

                        //Treatment of the error
                        Log ClassError = m_Database_Handler.Deserialize_JSON<Log>(res);
                        string errorText = ClassError.error;
                        MessageBox.Show(this, errorText, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                        return;
                    }
                    else
                    {
                        //Close the wait window
                        windowWait.Stop();

                        //Error connecting to web site
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                        return;
                    }
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                Close();
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Click on create an mission button
        /// </summary>
        private void Btn_Missions_Create_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the mission window
                Billing newMission = new Billing();
                newMission.id_HostAndHostess = m_Id_SelectedHostAndHostess;
                WindowMission.MainWindow_Mission missionWindow = new WindowMission.MainWindow_Mission(m_Global_Handler, m_Database_Handler, newMission);
                Nullable<bool> resShow = missionWindow.ShowDialog();

                //Validation
                if (resShow == true)
                {
                    if (missionWindow.m_Mission != null)
                    {
                        //Open the wait window
                        windowWait.Start(m_Global_Handler, "MissionCreationPrincipalMessage", "MissionCreationSecondaryMessage");

                        //Get the created mission
                        newMission = missionWindow.m_Mission;
                        newMission.id_HostAndHostess = m_Id_SelectedHostAndHostess;
                        newMission.date_Mission_creation = DateTime.Now.ToString();

                        //Add to database        
                        string res = m_Database_Handler.Add_MissionToDatabase();

                        //Treat the result
                        if (res.Contains("result"))
                        {
                            //Get id
                            string id = "";
                            MatchCollection matches = Regex.Matches(res, "[0-9]");
                            foreach (Match match in matches)
                            {
                                id += match.Value;
                            }
                            newMission.id = Convert.ToInt32(id);
                            newMission.num_Mission = m_MissionNextNumber;

                            //Action
                            m_Global_Handler.Error_Handler.WriteAction("Mission " + newMission.id + " created");

                            //Add to collection
                            SoftwareObjects.MissionsCollection.Add(newMission);

                            //Select the new mission
                            Select_Mission(newMission);

                            //Close the window wait
                            windowWait.Stop();

                            return;
                        }
                        else if (res.Contains("error"))
                        {
                            //Close the wait window
                            windowWait.Stop();

                            //Treatment of the error
                            Log ClassError = m_Database_Handler.Deserialize_JSON<Log>(res);
                            string errorText = ClassError.error;
                            MessageBox.Show(this, errorText, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                            return;
                        }
                        else
                        {
                            //Close the wait window
                            windowWait.Stop();

                            //Error connecting to web site
                            MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                            return;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                windowWait.Stop();
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }

        }

        /// <summary>
        /// Event
        /// Missions
        /// Click on delete the mission button
        /// </summary>
        private void Btn_Missions_Delete_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the mission
                Billing missionSel = Get_SelectedMissionFromButton();
                if (missionSel == null)
                {
                    return;
                }

                //Verification of an associate invoice
                if (missionSel.num_invoice != null && missionSel.num_invoice != "")
                {
                    string message = "";
                    if (missionSel.date_invoice_archived != null && missionSel.date_invoice_archived != "")
                    {
                        message = m_Global_Handler.Resources_Handler.Get_Resources("MissionAssociateToArchivedInvoice");
                    }
                    else
                    {
                        message = m_Global_Handler.Resources_Handler.Get_Resources("MissionAssociateToInProgressInvoice");
                    }
                    MessageBox.Show(this, message + " (" + missionSel.num_invoice + ")", m_Global_Handler.Resources_Handler.Get_Resources("MissionAssociateToInvoiceCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                //Confirm the delete
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissionConfirmDelete"),
                    m_Global_Handler.Resources_Handler.Get_Resources("MissionConfirmDeleteCaption"),
                    MessageBoxButton.YesNo, MessageBoxImage.Error);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                //Open the window wait
                windowWait.Start(m_Global_Handler, "MissionDeletionPrincipalMessage", "MissionDeletionSecondaryMessage");

                //Edit in database
                string res = m_Database_Handler.Delete_MissionToDatabase(missionSel.id);

                //Treat the result
                if (res.Contains("OK"))
                {
                    //Action
                    m_Global_Handler.Error_Handler.WriteAction("Mission " + missionSel.id + " deleted");

                    //Actualize and filter
                    Actualize_GridMissionsFromDatabase();
                    Filter_GridMissionsFromMissionsCollection(m_Mission_SelectedStatus);

                    //Clear the fields
                    Txt_Missions_ClientCompanyName.Text = "";
                    Txt_Missions_FirstName.Text = "";
                    Txt_Missions_LastName.Text = "";
                    Txt_Missions_Research.Text = "";
                    m_Id_SelectedHostAndHostess = "-1";
                    Txt_Missions_CreationDate.Text = "";
                    Cmb_Missions_SortBy.Text = "";

                    //Clear datagrid
                    m_DataGrid_Missions_ServicesCollection.Clear();
                    DataGrid_Missions_Services.Items.Refresh();

                    //Null buttons
                    m_Button_Mission_SelectedMission = null;
                    m_Button_Mission_SelectedEmail = null;
                    m_Button_Mission_SelectedInvoice = null;

                    //Disable the buttons
                    Btn_Missions_Duplicate.IsEnabled = false;
                    Btn_Missions_Edit.IsEnabled = false;
                    Btn_Missions_Delete.IsEnabled = false;
                    Btn_Missions_Close.IsEnabled = false;

                    //Close the wait window
                    windowWait.Stop();

                    return;
                }
                else if (res.Contains("error"))
                {
                    //Close the wait window
                    windowWait.Stop();

                    //Treatment of the error
                    Log ClassError = m_Database_Handler.Deserialize_JSON<Log>(res);
                    string errorText = ClassError.error;
                    MessageBox.Show(this, errorText, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }
                else
                {
                    //Close the wait window
                    windowWait.Stop();

                    //Error connecting to web site
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }
            }
            catch (Exception exception)
            {
                //Close the wait window
                windowWait.Stop();

                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Click on duplicate the mission button
        /// </summary>
        private void Btn_Missions_Duplicate_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the mission
                Billing missionSel = Get_SelectedMissionFromButton();
                if (missionSel == null)
                {
                    return;
                }

                //Open the wait window
                windowWait.Start(m_Global_Handler, "MissionDuplicationPrincipalMessage", "MissionDuplicationSecondaryMessage");

                //Get the created mission
                Billing newMission = new Billing(missionSel);
                newMission.id_HostAndHostess = m_Id_SelectedHostAndHostess;
                newMission.date_Mission_creation = DateTime.Now.ToString();

                //Add to database
                string res = m_Database_Handler.Add_MissionToDatabase();

                //Treat the result
                if (res.Contains("result"))
                {
                    //Get id
                    string id = "";
                    MatchCollection matches = Regex.Matches(res, "[0-9]");
                    foreach (Match match in matches)
                    {
                        id += match.Value;
                    }
                    newMission.id = Convert.ToInt32(id);
                    newMission.num_Mission = m_MissionNextNumber;

                    //Action
                    m_Global_Handler.Error_Handler.WriteAction("Mission " + missionSel.id + " duplicated to mission " + newMission.id);

                    //Add to collection
                    SoftwareObjects.MissionsCollection.Add(newMission);

                    //Select the new mission
                    Select_Mission(newMission);

                    //Close the window wait
                    windowWait.Stop();

                    return;
                }
                else if (res.Contains("error"))
                {
                    //Close the wait window
                    windowWait.Stop();

                    //Treatment of the error
                    Log ClassError = m_Database_Handler.Deserialize_JSON<Log>(res);
                    string errorText = ClassError.error;
                    MessageBox.Show(this, errorText, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }
                else
                {
                    //Close the wait window
                    windowWait.Stop();

                    //Error connecting to web site
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }

            }
            catch (Exception exception)
            {
                windowWait.Stop();
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Click on edit the mission button
        /// </summary>
        private void Btn_Missions_Edit_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the mission
                Billing missionSel = Get_SelectedMissionFromButton();
                if (missionSel == null)
                {
                    return;
                }

                //Verification of an associate invoice
                if (missionSel.num_invoice != null && missionSel.num_invoice != "")
                {
                    string message = "";
                    if (missionSel.date_invoice_archived != null && missionSel.date_invoice_archived != "")
                    {
                        message = m_Global_Handler.Resources_Handler.Get_Resources("MissionAssociateToArchivedInvoice");
                    }
                    else
                    {
                        message = m_Global_Handler.Resources_Handler.Get_Resources("MissionAssociateToInProgressInvoice");
                    }
                    MessageBox.Show(this, message + " (" + missionSel.num_invoice + ")", m_Global_Handler.Resources_Handler.Get_Resources("MissionAssociateToInvoiceCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                //Save sent, accepted and refused date for comparison
                string acceptedDate = missionSel.date_Mission_accepted;
                if (acceptedDate == null)
                {
                    acceptedDate = "";
                }
                string declinedDate = missionSel.date_Mission_declined;
                if (declinedDate == null)
                {
                    declinedDate = "";
                }
                string sentDate = missionSel.date_Mission_sent;
                if (sentDate == null)
                {
                    sentDate = "";
                }

                //Open the mission window
                WindowMission.MainWindow_Mission missionWindow = new WindowMission.MainWindow_Mission(m_Global_Handler, m_Database_Handler, missionSel);
                Nullable<bool> resShow = missionWindow.ShowDialog();

                //Validation
                if (resShow == true)
                {
                    if (missionWindow.m_Mission == null)
                    {
                        Exception error = new Exception("Exception thrown - Editing mission went wrong");
                        throw error;
                    }

                    //Open the window wait
                    windowWait.Start(m_Global_Handler, "MissionEditionPrincipalMessage", "MissionEditionSecondaryMessage");

                    //Get the edited mission
                    missionSel = missionWindow.m_Mission;

                    //Edit in database
                    string res = m_Database_Handler.Edit_MissionToDatabase(missionSel.id);

                    //Treat the result
                    if (res.Contains("result"))
                    {
                        //Action
                        m_Global_Handler.Error_Handler.WriteAction("Mission " + missionSel.id + " edited");
                    }
                    //Error in the edit
                    else if (res.Contains("error"))
                    {
                        //Close the wait window
                        windowWait.Stop();
                        //Treatment of the error
                        Log ClassError = m_Database_Handler.Deserialize_JSON<Log>(res);
                        string errorText = ClassError.error;
                        MessageBox.Show(this, errorText, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                        return;
                    }
                    else
                    {
                        //Close the wait window
                        windowWait.Stop();
                        //Error connecting to web site
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                        return;
                    }
                }

                //State of mission
                MissionStatus missionstate = MissionStatus.NONE;

                //Treatment of mission button
                StackPanel panel = (StackPanel)m_Button_Mission_SelectedMission.Parent;
                Button buttonMission = (Button)panel.Children[0];
                Button buttonMissionsendEmail = (Button)panel.Children[1];
                Manage_MissionButton(missionSel, buttonMission, buttonMissionsendEmail, false, true);

                if (missionstate == MissionStatus.DONE)
                {
                    Actualize_GridMissionsFromDatabase();
                    m_Mission_SelectedStatus = MissionStatus.DONE;
                    Filter_GridMissionsFromMissionsCollection(m_Mission_SelectedStatus);
                    Select_Mission(missionSel);
                }
                else if (missionstate == MissionStatus.DECLINED)
                {
                    Actualize_GridMissionsFromDatabase();
                    m_Mission_SelectedStatus = MissionStatus.DECLINED;
                    Filter_GridMissionsFromMissionsCollection(m_Mission_SelectedStatus);
                    Select_Mission(missionSel);
                }

                //Close the window
                windowWait.Stop();

                return;
            }
            catch (Exception exception)
            {
                //Close the window
                windowWait.Stop();

                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Click on generate missions statement in excel format button
        /// </summary>
        private void Btn_Missions_GenerateExcelStatement_Click(object sender, RoutedEventArgs e)
        {

            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get dates
                WindowDatePicker datePicker = new WindowDatePicker(m_Global_Handler, m_Global_Handler.Resources_Handler.Get_Resources("FirstDate"), m_Global_Handler.Resources_Handler.Get_Resources("EndDate"));
                Nullable<bool> resFirstDate = datePicker.ShowDialog();
                if (resFirstDate == false)
                {
                    return;
                }
                DateTime firstDatePicker = datePicker.m_FirstSelectedDate;
                DateTime endDatePicker = datePicker.m_EndSelectedDate;

                //Verification
                if (firstDatePicker > endDatePicker)
                {
                    DateTime temp = firstDatePicker;
                    firstDatePicker = endDatePicker;
                    endDatePicker = temp;
                }

                //Date time to string
                string format = "yyyy-MM-dd";
                string firstDatePickerStr = Convert.ToDateTime(firstDatePicker).ToString(format);
                firstDatePickerStr = m_Global_Handler.DateAndTime_Handler.Treat_Date(firstDatePickerStr, m_Global_Handler.Language_Handler);
                firstDatePickerStr = firstDatePickerStr.Replace("/", "-");
                string endDatePickerStr = Convert.ToDateTime(endDatePicker).ToString(format);
                endDatePickerStr = m_Global_Handler.DateAndTime_Handler.Treat_Date(endDatePickerStr, m_Global_Handler.Language_Handler);
                endDatePickerStr = endDatePickerStr.Replace("/", "-");

                //Get the name of the pdf
                string fileNameXLS = m_Global_Handler.Resources_Handler.Get_Resources("Mission") + " - " +
                    m_Global_Handler.Resources_Handler.Get_Resources("StatementBetween") + firstDatePickerStr + " "
                    + m_Global_Handler.Resources_Handler.Get_Resources("and") + " " + endDatePickerStr;
                System.Windows.Forms.SaveFileDialog saveFile = new System.Windows.Forms.SaveFileDialog();
                saveFile.FileName = fileNameXLS;
                saveFile.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx;*.xls";
                if (saveFile.ShowDialog() != System.Windows.Forms.DialogResult.OK || saveFile.FileName.Length == 0)
                {
                    return;
                }
                else
                {
                    fileNameXLS = saveFile.FileName;
                }

                //Open the wait window
                windowWait.Start(m_Global_Handler, "MissionsExcelGenerationPrincipalMessage", "MissionsExcelGenerationSecondaryMessage");

                //Get the list of missions in this two dates
                List<Billing> missionsToExport = new List<Billing>();
                for (int iBill = 0; iBill < SoftwareObjects.MissionsCollection.Count; ++iBill)
                {
                    Billing mission = SoftwareObjects.MissionsCollection[iBill];
                    DateTime dateCreation = Convert.ToDateTime(mission.date_Mission_creation);
                    if (dateCreation >= firstDatePicker && dateCreation <= endDatePicker)
                    {
                        missionsToExport.Add(mission);
                    }
                }

                //Sort the collection
                missionsToExport.Sort(delegate (Billing x, Billing y)
                {
                    if (x.num_Mission == null && y.num_Mission == null) return 0;
                    else if (x.num_Mission == null) return 1;
                    else if (y.num_Mission == null) return -1;
                    else
                    {
                        return x.date_Mission_creation.CompareTo(y.date_Mission_creation);
                    }
                });

                //Write in a file with tab separated
                List<string> lines = new List<string>();
                string columnsHeader = m_Global_Handler.Resources_Handler.Get_Resources("CustomerID") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Customer") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Company") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("MissionNumber") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Missionsubject") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CreationDate") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("BillingDate") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("AmountET") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("AmountVAT") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("AmountIT");
                lines.Add(columnsHeader);
                for (int iMission = 0; iMission < missionsToExport.Count; ++iMission)
                {
                    Billing mission = missionsToExport[iMission];
                    Hostess hostess = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(mission.id_HostAndHostess));
                    string Hostesstr = "";
                    string hostessID = "";
                    string hostessCompany = "";
                    if (hostess != null)
                    {
                        Hostesstr = hostess.firstname + " " + hostess.lastname;
                        hostessID = hostess.id.ToString();
                        hostessCompany = hostess.id_paycheck;
                    }
                    double amountVAT = mission.grand_amount - mission.amount;
                    string line = hostessID + "\t" + Hostesstr + "\t" + hostessCompany + "\t" + mission.num_Mission + "\t" + mission.subject + "\t" +
                        m_Global_Handler.DateAndTime_Handler.Treat_Date(mission.date_Mission_creation, m_Global_Handler.Language_Handler) + "\t" +
                        m_Global_Handler.DateAndTime_Handler.Treat_Date(mission.date_invoice_creation, m_Global_Handler.Language_Handler) + "\t" +
                        mission.amount.ToString("0.00", CultureInfo.GetCultureInfo(m_Global_Handler.Language_Handler)) + "\t" +
                        amountVAT.ToString("0.00", CultureInfo.GetCultureInfo(m_Global_Handler.Language_Handler)) + "\t" +
                        mission.grand_amount.ToString("0.00", CultureInfo.GetCultureInfo(m_Global_Handler.Language_Handler));
                    lines.Add(line);
                }

                //Generate the excel file
                int result = m_Global_Handler.Excel_Handler.Generate_ExcelStatement(fileNameXLS, lines);
                if (result == -1)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissionsStatementExcelGenerationFailed"),
                                m_Global_Handler.Resources_Handler.Get_Resources("MissionsStatementExcelGenerationFailedCaption"),
                                MessageBoxButton.OK, MessageBoxImage.Error);
                    //Close the wait window
                    windowWait.Stop();

                    return;
                }

                //Action
                m_Global_Handler.Error_Handler.WriteAction("Missions exported to excel file " + fileNameXLS);

                //Close the wait window
                windowWait.Stop();

                //Open the file
                Process.Start(fileNameXLS);
            }
            catch (Exception exception)
            {
                //Close the wait window
                windowWait.Stop();

                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Click on show archived missions button
        /// </summary>
        private void Btn_Missions_ShowClosed_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window wait
                windowWait.Start(m_Global_Handler, "MissionshowArchivedPrincipalMessage", "MissionshowArchivedSecondaryMessage");

                //Mode
                m_Mission_IsArchiveMode = true;

                //Actualize
                Actualize_GridMissionsFromDatabase();
                m_Mission_SelectedStatus = MissionStatus.NONE;

                //Clear the fields
                Txt_Missions_ClientCompanyName.Text = "";
                Txt_Missions_FirstName.Text = "";
                Txt_Missions_LastName.Text = "";
                Txt_Missions_Research.Text = "";
                m_Id_SelectedHostAndHostess = "-1";
                Txt_Missions_CreationDate.Text = "";
                Cmb_Missions_SortBy.Text = "";

                //Clear datagrid
                m_DataGrid_Missions_ServicesCollection.Clear();
                DataGrid_Missions_Services.Items.Refresh();

                //Manage buttons
                Btn_Missions_Duplicate.IsEnabled = false;
                Btn_Missions_Edit.IsEnabled = false;
                Btn_Missions_Delete.IsEnabled = false;
                Btn_Missions_Close.IsEnabled = false;
                Btn_Missions_Close.Content = m_Global_Handler.Resources_Handler.Get_Resources("Restore");
                Btn_Missions_Legend_Mission_Done.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Billed.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Created.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Declined.Background = m_Color_Button;

                //Close the wait window
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close the wait window
                windowWait.Stop();

                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Click on show in progress missions button
        /// </summary>
        private void Btn_Missions_ShowInProgress_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window wait
                windowWait.Start(m_Global_Handler, "MissionshowInProgressPrincipalMessage", "MissionshowInProgressSecondaryMessage");

                //Mode
                m_Mission_IsArchiveMode = false;

                //Actualize
                Actualize_GridMissionsFromDatabase();
                m_Mission_SelectedStatus = MissionStatus.NONE;

                //Clear the fields
                Txt_Missions_ClientCompanyName.Text = "";
                Txt_Missions_FirstName.Text = "";
                Txt_Missions_LastName.Text = "";
                Txt_Missions_Research.Text = "";
                m_Id_SelectedHostAndHostess = "-1";
                Txt_Missions_CreationDate.Text = "";
                Cmb_Missions_SortBy.Text = "";

                //Clear datagrid
                m_DataGrid_Missions_ServicesCollection.Clear();
                DataGrid_Missions_Services.Items.Refresh();

                //Manage buttons
                Btn_Missions_Duplicate.IsEnabled = false;
                Btn_Missions_Edit.IsEnabled = false;
                Btn_Missions_Delete.IsEnabled = false;
                Btn_Missions_Close.IsEnabled = false;
                Btn_Missions_Close.Content = m_Global_Handler.Resources_Handler.Get_Resources("Archive");

                Btn_Missions_Legend_Mission_Done.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Billed.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Created.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Declined.Background = m_Color_Button;

                //Close the wait window
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close the wait window
                windowWait.Stop();

                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Select a type of sort in the missions combo box
        /// Sort in a(n) ascending/descending way the missions panels based on the selected item
        /// </summary>
        private void Cmb_Missions_SortBy_SelectionChanged(object sender, EventArgs e)
        {
            //Invert sorting
            m_IsSortMission_Ascending = !m_IsSortMission_Ascending;

            //Save in case of problems
            List<Billing> savedCollection = new List<Billing>();
            for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
            {
                //Get mission
                Billing mission = SoftwareObjects.MissionsCollection[iMission];
                if (mission.num_Mission == "" || mission.num_Mission == null)
                {
                    continue;
                }
                //Archived
                if (m_Mission_IsArchiveMode == true && (mission.date_Mission_archived != null && mission.date_Mission_archived != ""))
                {
                    savedCollection.Add(mission);
                }
                //In progress
                else if (m_Mission_IsArchiveMode == false && (mission.date_Mission_archived == null || mission.date_Mission_archived == ""))
                {
                    savedCollection.Add(mission);
                }
            }
            m_Grid_Details_Missions_MissionsCollection = savedCollection;

            //No sort
            if (Cmb_Missions_SortBy.SelectedIndex == -1)
            {
                //Clear the grid
                Grid_Missions_Details.Children.Clear();
                Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                //Actualize the collection                
                Actualize_GridMissionsFromMissionsCollection();
            }
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("Amount"))
            {
                //Sort by amount
                try
                {
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Billing x, Billing y)
                    {
                        if (m_IsSortMission_Ascending == true)
                            return x.amount.CompareTo(y.amount);
                        else
                            return y.amount.CompareTo(x.amount);
                    });

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    m_Grid_Details_Missions_MissionsCollection = savedCollection;

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();

                    return;
                }
            }
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("LastNameHostess"))
            {
                //Sort by last name
                try
                {
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Billing x, Billing y)
                    {
                        Hostess hostess_x = SoftwareObjects.HostsAndHotessesCollection.Find(z => z.id.Equals(x.id_HostAndHostess));
                        Hostess hostess_y = SoftwareObjects.HostsAndHotessesCollection.Find(z => z.id.Equals(y.id_HostAndHostess));
                        if (hostess_x.lastname == null && hostess_y.lastname == null) return 0;
                        else if (hostess_x.lastname == null) return -1;
                        else if (hostess_y.lastname == null) return 1;
                        else
                        {
                            if (m_IsSortMission_Ascending == false)
                                return hostess_x.lastname.CompareTo(hostess_y.lastname);
                            else
                                return hostess_y.lastname.CompareTo(hostess_x.lastname);
                        }
                    });

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    m_Grid_Details_Missions_MissionsCollection = savedCollection;

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();

                    return;
                }
            }
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("City"))
            {
                //Sort by city
                try
                {
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Billing x, Billing y)
                    {
                        Hostess hostess_x = SoftwareObjects.HostsAndHotessesCollection.Find(z => z.id.Equals(x.id_HostAndHostess));
                        Hostess hostess_y = SoftwareObjects.HostsAndHotessesCollection.Find(z => z.id.Equals(y.id_HostAndHostess));
                        if (hostess_x.city == null && hostess_y.city == null) return 0;
                        else if (hostess_x.city == null) return -1;
                        else if (hostess_y.city == null) return 1;
                        {
                            if (m_IsSortMission_Ascending == false)
                                return hostess_x.city.CompareTo(hostess_y.city);
                            else
                                return hostess_y.city.CompareTo(hostess_x.city);
                        }
                    });

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    m_Grid_Details_Missions_MissionsCollection = savedCollection;

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();

                    return;
                }
            }
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("CreationDate"))
            {
                //Sort by creation date
                try
                {
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Billing x, Billing y)
                    {
                        if (x.date_Mission_creation == null && y.date_Mission_creation == null) return 0;
                        else if (x.date_Mission_creation == null) return 1;
                        else if (y.date_Mission_creation == null) return -1;
                        else
                        {
                            DateTime xDate = Convert.ToDateTime(x.date_Mission_creation);
                            DateTime yDate = Convert.ToDateTime(y.date_Mission_creation);
                            if (m_IsSortMission_Ascending == false)
                                return xDate.CompareTo(yDate);
                            else
                                return yDate.CompareTo(xDate);
                        }
                    });

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    m_Grid_Details_Missions_MissionsCollection = savedCollection;

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();

                    return;
                }
            }
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("MissionNumber"))
            {
                //Sort by subject
                try
                {
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Billing x, Billing y)
                    {
                        if (x.num_Mission == null && y.num_Mission == null) return 0;
                        else if (x.num_Mission == null) return 1;
                        else if (y.num_Mission == null) return -1;
                        else
                        {
                            if (m_IsSortMission_Ascending == false)
                                return x.num_Mission.CompareTo(y.num_Mission);
                            else
                                return y.num_Mission.CompareTo(x.num_Mission);
                        }
                    });

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    m_Grid_Details_Missions_MissionsCollection = savedCollection;

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();

                    return;
                }
            }
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("Missionsubject"))
            {
                //Sort by subject
                try
                {
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Billing x, Billing y)
                    {
                        if (x.subject == null && y.subject == null) return 0;
                        else if (x.subject == null) return 1;
                        else if (y.subject == null) return -1;
                        else
                        {
                            if (m_IsSortMission_Ascending == false)
                                return x.subject.CompareTo(y.subject);
                            else
                                return y.subject.CompareTo(x.subject);
                        }
                    });

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    m_Grid_Details_Missions_MissionsCollection = savedCollection;

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();

                    return;
                }
            }
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("Status"))
            {
                //Sort by status
                Sort_MissionsByStatus(savedCollection);
            }

            //Apply filter
            Filter_GridMissionsFromMissionsCollection(m_Mission_SelectedStatus);
        }

        /// <summary>
        /// Event
        /// Missions
        /// Auto generating columns for the datagrid containg the services associated to the selected mission
        /// </summary>
        private void Datagrid_Missions_Services_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            string columnHeader = e.Column.Header.ToString();
            if (columnHeader == "description")
            {
                e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Description");
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                DataGridTextColumn dataColumn = (DataGridTextColumn)e.Column;
                dataColumn.ElementStyle = new Style(typeof(TextBlock));
                dataColumn.ElementStyle.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap));
            }
            else if (columnHeader == "quantity")
            {
                e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Quantity");
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            else if (columnHeader == "reference")
            {
                e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Reference");
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Mouse down on the main mission button event
        /// Select the mission
        /// </summary>
        private void IsPressed_MissionButton(object sender, RoutedEventArgs ev)
        {
            if (sender != null)
            {
                try
                {
                    //Get the mission's id
                    Button buttonSel = (Button)sender;
                    StackPanel stackSel = (StackPanel)buttonSel.Parent;
                    int idSel = (int)stackSel.Tag;

                    //Get the mission
                    Billing missionSel = SoftwareObjects.MissionsCollection.Find(x => x.id.Equals(idSel));
                    if (missionSel == null)
                    {
                        Exception error = new Exception("Mission not found !");
                        throw error;
                    }

                    //Get the associated hostess
                    Hostess Hostessel = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(missionSel.id_HostAndHostess));
                    if (Hostessel == null)
                    {
                        MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissionHostessNotFound"),
                            m_Global_Handler.Resources_Handler.Get_Resources("MissionHostessNotFoundCaption"),
                            MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        return;
                    }
                    m_Id_SelectedHostAndHostess = Hostessel.id;

                    if (missionSel != null)
                    {
                        //Fill the fields
                        Txt_Missions_FirstName.Text = Hostessel.firstname;
                        Txt_Missions_LastName.Text = Hostessel.lastname;

                        //Treat dates
                        Txt_Missions_CreationDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.date_Mission_creation, m_Global_Handler.Language_Handler);

                        string date = "";
                        string acceptedDate = "";
                        string declinedDate = "";
                        string labelDate = "";
                        if (missionSel.date_Mission_accepted != null)
                        {
                            acceptedDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.date_Mission_accepted, m_Global_Handler.Language_Handler);
                            date = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.date_Mission_accepted, m_Global_Handler.Language_Handler);
                            labelDate = m_Global_Handler.Resources_Handler.Get_Resources("AcceptedDate");
                        }
                        else if (missionSel.date_Mission_declined != null)
                        {
                            declinedDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.date_Mission_declined, m_Global_Handler.Language_Handler);
                            date = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.date_Mission_declined, m_Global_Handler.Language_Handler);
                            labelDate = m_Global_Handler.Resources_Handler.Get_Resources("DeclinedDate");
                        }

                        //Manage the button
                        if (m_Button_Mission_SelectedMission != null)
                        {
                            if (m_Mission_IsArchiveMode == false)
                            {
                                m_Button_Mission_SelectedMission.Background = m_Color_Mission;
                            }
                            else
                            {
                                m_Button_Mission_SelectedMission.Background = m_Color_ArchivedMission;
                            }
                        }
                        if (m_Button_Mission_SelectedEmail != null)
                        {
                            if (m_Mission_IsArchiveMode == false)
                            {
                                m_Button_Mission_SelectedEmail.Background = m_Color_Mission;
                            }
                            else
                            {
                                m_Button_Mission_SelectedEmail.Background = m_Color_ArchivedMission;
                            }

                        }
                        if (m_Button_Mission_SelectedInvoice != null)
                        {
                            if (m_Mission_IsArchiveMode == false)
                            {
                                m_Button_Mission_SelectedInvoice.Background = m_Color_Mission;
                            }
                            else
                            {
                                m_Button_Mission_SelectedInvoice.Background = m_Color_ArchivedMission;
                            }

                        }
                        for (int iChild = 0; iChild < stackSel.Children.Count; ++iChild)
                        {
                            Button childButton = (Button)stackSel.Children[iChild];
                            if (m_Mission_IsArchiveMode == false)
                            {
                                childButton.Background = m_Color_SelectedMission;
                            }
                            else
                            {
                                childButton.Background = m_Color_SelectedArchivedMission;
                            }
                            if (iChild == 0)
                            {
                                m_Button_Mission_SelectedMission = childButton;
                            }
                            else if (childButton.Tag.ToString() == "Email")
                            {
                                m_Button_Mission_SelectedEmail = childButton;
                            }
                            else if (childButton.Tag.ToString() == "Invoice")
                            {
                                m_Button_Mission_SelectedInvoice = childButton;
                            }
                        }

                        if (m_Mission_IsArchiveMode == false)
                        {
                            //Enable the buttons
                            Btn_Missions_Duplicate.IsEnabled = true;
                            Btn_Missions_Edit.IsEnabled = true;
                            Btn_Missions_Delete.IsEnabled = true;
                            Btn_Missions_Close.IsEnabled = true;
                        }
                        else
                        {
                            //Disable the buttons but Restore
                            Btn_Missions_Duplicate.IsEnabled = false;
                            Btn_Missions_Edit.IsEnabled = false;
                            Btn_Missions_Delete.IsEnabled = false;
                            Btn_Missions_Close.IsEnabled = true;
                        }
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Mouse down on the mission's email button
        /// Select the mission, generate it in a PDF format and open a new email to the mission hostess with the PDF attached
        /// </summary>
        private void IsPressed_MissionEmailButton(object sender, RoutedEventArgs ev)
        {
            if (sender != null)
            {
                //No email in archive mode
                if (m_Mission_IsArchiveMode == true)
                {
                    return;
                }

                //Creation of the wait window
                WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

                try
                {
                    //Get the email adress
                    Button emailButton = (Button)sender;

                    //Manage the buttons
                    if (m_Button_Mission_SelectedMission != null)
                    {
                        m_Button_Mission_SelectedMission.Background = m_Color_Mission;
                    }
                    if (m_Button_Mission_SelectedEmail != null)
                    {
                        m_Button_Mission_SelectedEmail.Background = m_Color_Mission;
                    }
                    if (m_Button_Mission_SelectedInvoice != null)
                    {
                        m_Button_Mission_SelectedInvoice.Background = m_Color_Mission;
                    }
                    StackPanel stackSel = (StackPanel)emailButton.Parent;
                    for (int iChild = 0; iChild < stackSel.Children.Count; ++iChild)
                    {
                        Button childButton = (Button)stackSel.Children[iChild];
                        childButton.Background = m_Color_SelectedMission;
                        if (iChild == 0)
                        {
                            m_Button_Mission_SelectedMission = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_Mission_SelectedEmail = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Invoice")
                        {
                            m_Button_Mission_SelectedInvoice = childButton;
                        }
                    }

                    //Get the mission
                    Billing missionSel = Get_SelectedMissionFromButton();
                    if (missionSel == null)
                    {
                        return;
                    }

                    //Get the associated hostess
                    Hostess Hostessel = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(missionSel.id_HostAndHostess));
                    if (Hostessel == null)
                    {
                        Exception error = new Exception("Hostess not found while trying to send the email !");
                        throw error;
                    }

                    //Fill the fields
                    Txt_Missions_FirstName.Text = Hostessel.firstname;
                    Txt_Missions_LastName.Text = Hostessel.lastname;

                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "MissionPDFGenerationPrincipalMessage", "MissionPDFGenerationSecondaryMessage");

                    //Generate the mission
                    int result = m_Global_Handler.Word_Handler.Generate_Bill(m_Global_Handler, missionSel, BillType.Mission, false, "");

                    //Close window wait
                    windowWait.Stop();

                    //Show result of the generation
                    if (result == -1)
                    {
                        MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissionPDFGenerationFailed"),
                                    m_Global_Handler.Resources_Handler.Get_Resources("MissionPDFGenerationFailedCaption"),
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    //Get the email adress
                    string emailAddress = Hostessel.email;
                    string emailSubject = m_Global_Handler.Resources_Handler.Get_Resources("Mission") + " - " + missionSel.subject;
                    string emailBody = m_Global_Handler.Resources_Handler.Get_Resources("MissionMailBody");

                    //Ask for visualizing the mission
                    string generatedPDFFileName = m_Global_Handler.Word_Handler.Get_GeneratedBillFileName();
                    MessageBoxResult resultVisualize = MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("MissionPDFOpenConfirmation"),
                        m_Global_Handler.Resources_Handler.Get_Resources("MissionPDFOpenConfirmationCaption"),
                        MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (resultVisualize == MessageBoxResult.Yes)
                    {
                        Process.Start(generatedPDFFileName);
                    }

                    //Ask for confirmation
                    MessageBoxResult resultConfirmation = MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("MissionPDFSendEmailConfirmation"),
                        m_Global_Handler.Resources_Handler.Get_Resources("MissionPDFSendEmailCaption"),
                        MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (resultConfirmation == MessageBoxResult.No)
                    {
                        return;
                    }

                    //Open draft in default email client
                    var file = generatedPDFFileName.Replace("file:\\", "");
                    string[] attachments = new string[] { file };
                    string[] recipients = new string[] { emailAddress };
                    try
                    {
                        new Email().Compose_Mail(recipients, emailSubject, emailBody, attachments);
                    }
                    catch (Exception exceptionMAPI)
                    {
                        //Impossible to manage attachments
                        m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exceptionMAPI);

                        //WIthout attachment
                        Process p = new Process();
                        string mailto = "mailto:" + emailAddress + "?subject=" + emailSubject + "&body=" + emailBody;
                        ProcessStartInfo ps = new ProcessStartInfo(mailto);
                        ps.CreateNoWindow = false;
                        ps.UseShellExecute = true;
                        p.StartInfo = ps;
                        p.Start();
                        p.WaitForExit();
                    }

                    //Validate the sent date
                    string format = "yyyy-MM-dd";
                    string sentDate = DateTime.Now.ToString(format);
                    missionSel.date_Mission_sent = sentDate;

                }
                catch (Exception exception)
                {
                    //Close window wait
                    windowWait.Stop();

                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        /// <summary>
        /// Boolean tagging the full actualization avoiding to do it multiple time
        /// This variable is only used in the research mission text box text changed event
        /// </summary>
        bool actualizationMissionsDone = true;
        /// <summary>
        /// Event
        /// Missions
        /// Text changed in the research mission text box
        /// Only missions containing the text (at least 2 characters) are shown 
        /// </summary>
        private void Txt_Missions_Research_TextChanged(object sender, TextChangedEventArgs e)
        {
            //Get the researched text
            string researchedText = Txt_Missions_Research.Text;
            researchedText = researchedText.ToLower();

            //Clear the fields
            Txt_Missions_ClientCompanyName.Text = "";
            Txt_Missions_FirstName.Text = "";
            Txt_Missions_LastName.Text = "";
            Txt_Missions_CreationDate.Text = "";
            m_DataGrid_Missions_ServicesCollection.Clear();
            DataGrid_Missions_Services.Items.Refresh();
            m_Button_Mission_SelectedMission = null;
            m_Button_Mission_SelectedEmail = null;
            m_Button_Mission_SelectedInvoice = null;

            //Disable the buttons
            Btn_Missions_Duplicate.IsEnabled = false;
            Btn_Missions_Edit.IsEnabled = false;
            Btn_Missions_Delete.IsEnabled = false;
            Btn_Missions_Close.IsEnabled = false;

            //Verifications - 3 letters minimum
            if (researchedText.Length < 2)
            {
                if (!actualizationMissionsDone)
                {
                    Grid_Missions_Details.Children.Clear();
                    Actualize_GridMissionsFromDatabase();
                    actualizationMissionsDone = true;
                    Filter_GridMissionsFromMissionsCollection(m_Mission_SelectedStatus);
                    return;
                }
            }
            else
            {
                try
                {
                    //Research in each private member of the list of mission
                    List<Billing> foundMissionsList = new List<Billing>();
                    for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
                    {
                        //Get the mission
                        Billing processedMission = SoftwareObjects.MissionsCollection[iMission];

                        //Get the associated hostess
                        Hostess processedHostess = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(processedMission.id_HostAndHostess));

                        //Research
                        if (processedMission.subject.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                        else if (processedHostess.firstname.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                        else if (processedHostess.lastname.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                        else if (processedHostess.zipcode.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                        else if (processedHostess.email.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                        else if (processedHostess.city.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                    }

                    //Displaying the found missions list
                    Grid_Missions_Details.Children.Clear();
                    actualizationMissionsDone = false;
                    if (foundMissionsList.Count > 0)
                    {
                        //Initialize counter for columns and rows of Grid_Missions
                        m_GridMissions_Column = 0;
                        m_GridMissions_Row = 0;
                        for (int iMission = 0; iMission < foundMissionsList.Count; ++iMission)
                        {
                            Add_MissionToGrid(foundMissionsList[iMission]);
                        }

                        //Save the research collection
                        m_Grid_Details_Missions_MissionsCollection = foundMissionsList;
                    }

                    //Apply filter
                    Filter_GridMissionsFromMissionsCollection(m_Mission_SelectedStatus);
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        #endregion

        #region Hosts and hostesses

        /// <summary>
        /// Event
        /// Hostess
        /// Click on add hostess
        /// </summary>
        private void Btn_HostAndHostess_Create_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window
                WindowHostOrHostess.MainWindow_HostOrHostess hostOrHostess_CreateWindow = new WindowHostOrHostess.MainWindow_HostOrHostess(m_Global_Handler,
                    m_Database_Handler, m_Cities_DocFrance, false, null);
                Nullable<bool> resShow = hostOrHostess_CreateWindow.ShowDialog();

                //Validation
                if (resShow == true)
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "HostOrHostessCreationPrincipalMessage", "HostOrHostessCreationSecondaryMessage");

                    //Refresh datagrid
                    m_DataGrid_HostAndHostess_MissionsCollection.Clear();

                    //Add hostess to grid
                    Hostess hostOrHostessToAdd = hostOrHostess_CreateWindow.m_HostOrHostess;
                    Add_HostOrHostessToGrid(hostOrHostess_CreateWindow.m_HostOrHostess);

                    //Action
                    m_Global_Handler.Error_Handler.WriteAction("Host/Hostess " + hostOrHostessToAdd.firstname + " " + hostOrHostessToAdd.lastname + " created");

                    //Select the last hostess
                    if (m_Button_HostAndHostess_SelectedHostAndHostess != null)
                    {
                        m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_HostAndHostess_SelectedCellPhone != null)
                    {
                        m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_HostAndHostess_SelectedEmail != null)
                    {
                        m_Button_HostAndHostess_SelectedEmail.Background = m_Color_HostAndHostess;
                    }
                    StackPanel stack = Grid_HostAndHostess_Details.Children.Cast<StackPanel>().First(f => Grid.GetRow(f) == m_GridHostess_Row && Grid.GetColumn(f) == m_GridHostess_Column - 1);
                    for (int iChild = 0; iChild < stack.Children.Count; ++iChild)
                    {
                        Button childButton = (Button)stack.Children[iChild];
                        childButton.Background = m_Color_SelectedHostAndHostess;
                        if (childButton.Tag.ToString() != "" && childButton.Tag.ToString() != "CellPhone" && childButton.Tag.ToString() != "Phone" && childButton.Tag.ToString() != "Email")
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess = childButton;
                            m_Id_SelectedHostAndHostess = (string)childButton.Tag;
                        }
                        else if (childButton.Tag.ToString() == "CellPhone")
                        {
                            m_Button_HostAndHostess_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_HostAndHostess_SelectedEmail = childButton;
                        }

                        //Display the last hostess
                        ScrollViewer_HostAndHostess_Details.ScrollToBottom();
                    }

                    //Enable the buttons
                    Btn_HostAndHostess_Edit.IsEnabled = true;
                    Btn_HostAndHostess_Delete.IsEnabled = true;

                    //Close the wait window
                    Thread.Sleep(500);
                    windowWait.Stop();

                    return;
                }
            }
            catch (Exception exception)
            {
                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();

                //Write the error into log
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Click on delete hostess
        /// </summary>
        private void Btn_HostAndHostess_Delete_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the hostess
                Hostess hostOrHostess = Get_SelectedHostOrHostessFromButton();
                if (hostOrHostess == null)
                {
                    return;
                }

                //Verification of the presence of missions or invoices
                List<Billing> associatedMissions = new List<Billing>();
                for (int iBill = 0; iBill < SoftwareObjects.MissionsCollection.Count; ++iBill)
                {
                    Billing billSel = SoftwareObjects.MissionsCollection[iBill];
                    if (billSel.id_HostAndHostess == hostOrHostess.id)
                    {
                        associatedMissions.Add(billSel);
                    }
                }
                if (associatedMissions.Count > 0)
                {
                    string message = m_Global_Handler.Resources_Handler.Get_Resources("HostessForbiddenDeleteText") + "\n";
                    message += m_Global_Handler.Resources_Handler.Get_Resources("Mission") + ": ";
                    for (int iMission = 0; iMission < associatedMissions.Count; ++iMission)
                    {
                        message += " " + associatedMissions[iMission].num_Mission;
                    }
                    MessageBox.Show(this, message, m_Global_Handler.Resources_Handler.Get_Resources("HostessForbiddenDeleteCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                //Confirm the delete
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("HostessConfirmDelete"),
                    m_Global_Handler.Resources_Handler.Get_Resources("HostessConfirmDeleteCaption"), MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                //Edit in internet database
                string res = m_Database_Handler.Delete_HostAndHostessToDatabase(hostOrHostess.id);

                //Treat the result
                if (res.Contains("OK"))
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "HostessDeletionPrincipalMessage", "HostessDeletionSecondaryMessage");

                    //Action
                    m_Global_Handler.Error_Handler.WriteAction("Host/Hostess " + hostOrHostess.firstname + " " + hostOrHostess.lastname + " deleted");

                    //Delete the hostess to the collection
                    SoftwareObjects.HostsAndHotessesCollection.Remove(hostOrHostess);
                    Grid_HostAndHostess_Details.Children.Clear();
                    Actualize_GridHostsAndHostessesFromDatabase();

                    //Delete photo repository
                    string repository = SoftwareObjects.GlobalSettings.photos_repository + "\\" + hostOrHostess.id;
                    if (Directory.Exists(repository))
                    {
                        Directory.Delete(repository, true);
                    }

                    //Clear the button
                    m_Button_HostAndHostess_SelectedHostAndHostess = null;
                    m_Button_HostAndHostess_SelectedEmail = null;
                    m_Button_HostAndHostess_SelectedCellPhone = null;

                    //Disable the buttons
                    Btn_HostAndHostess_Edit.IsEnabled = false;
                    Btn_HostAndHostess_Delete.IsEnabled = false;

                    //Close the window wait
                    Thread.Sleep(500);
                    windowWait.Stop();

                    return;
                }
                else if (res.Contains("Error"))
                {
                    //Close the wait window
                    Thread.Sleep(500);
                    windowWait.Stop();

                    //Treatment of the error
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
            catch (Exception exception)
            {
                //Close the wait window
                windowWait.Stop();

                //Write the error into log
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Click on edit hostess
        /// </summary>
        private void Btn_HostAndHostess_Edit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Get the hostess
                Hostess hostOrHostess = Get_SelectedHostOrHostessFromButton();
                if (hostOrHostess == null)
                {
                    return;
                }

                //Open the window
                WindowHostOrHostess.MainWindow_HostOrHostess hostOrHostess_EditWindow = new WindowHostOrHostess.MainWindow_HostOrHostess(m_Global_Handler, m_Database_Handler,
                    m_Cities_DocFrance, true, hostOrHostess);
                Nullable<bool> resShow = hostOrHostess_EditWindow.ShowDialog();

                //Validation
                if (resShow == true)
                {
                    //Modify the hostess
                    hostOrHostess = hostOrHostess_EditWindow.m_HostOrHostess;

                    //Action
                    m_Global_Handler.Error_Handler.WriteAction("Host/Hostess " + hostOrHostess.firstname + " " + hostOrHostess.lastname + " edited");

                    //Treatment of hostess info
                    TextBlock hostessInfo = new TextBlock();
                    Run names = new Run("\n" + hostOrHostess.firstname + " " + hostOrHostess.lastname);
                    names.FontSize = 15;
                    names.FontWeight = FontWeights.Bold;
                    hostessInfo.Inlines.Add(names);
                    hostessInfo.Inlines.Add(new LineBreak());
                    string ageInfo = "";
                    if (hostOrHostess.birth_date.Split(' ').Length == 3)
                    {
                        ageInfo = hostOrHostess.birth_date;
                        int age = Age(hostOrHostess.birth_date);
                        if (age != -1)
                        {
                            ageInfo += ", " + age.ToString() + " " + m_Global_Handler.Resources_Handler.Get_Resources("YearsOld");
                        }
                    }
                    string creationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(hostOrHostess.date_creation, m_Global_Handler.Language_Handler);
                    string info = hostOrHostess.address + "\n" +
                        hostOrHostess.zipcode + ", " + hostOrHostess.city + "\n" +
                        ageInfo + "\n" +
                        m_Global_Handler.Resources_Handler.Get_Resources("CreationDate") + " : " + creationDate + "\n";
                    hostessInfo.Inlines.Add(info);
                    //Create stack button
                    StackPanel buttonStack = new StackPanel();
                    buttonStack.Orientation = Orientation.Horizontal;
                    //Create image
                    Image imgHostess = new Image();
                    string repository = SoftwareObjects.GlobalSettings.photos_repository + "\\" + hostOrHostess.id;
                    if (Directory.Exists(repository))
                    {
                        //Get photos
                        string[] files = Directory.GetFiles(repository);
                        //Display the first photo
                        if (files.Length > 0)
                        {
                            string photoFilename = files[0];
                            System.Drawing.Bitmap photoBmp = m_Global_Handler.Image_Handler.Load_Bitmap(photoFilename);
                            double factor = 100.0 / photoBmp.Height;
                            System.Drawing.Bitmap tmp = (System.Drawing.Bitmap)m_Global_Handler.Image_Handler.ResizeImage(photoBmp, (int)(factor * photoBmp.Width), (int)(factor * photoBmp.Height));
                            imgHostess.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(tmp);
                        }
                    }
                    //Effect
                    System.Windows.Media.Effects.DropShadowEffect dropShadowEffect = new System.Windows.Media.Effects.DropShadowEffect();
                    Color shadowColor = new Color();
                    shadowColor.ScA = 1;
                    shadowColor.ScB = 0;
                    shadowColor.ScG = 0;
                    shadowColor.ScR = 0;
                    dropShadowEffect.Color = shadowColor;
                    dropShadowEffect.Direction = 320;
                    dropShadowEffect.ShadowDepth = 10;
                    dropShadowEffect.BlurRadius = 10;
                    dropShadowEffect.Opacity = 0.5;
                    imgHostess.Effect = dropShadowEffect;
                    //Add image
                    buttonStack.Children.Add(imgHostess);

                    //Add void
                    Label voidLabel = new Label();
                    voidLabel.Content = "\t";
                    buttonStack.Children.Add(voidLabel);

                    //Add label
                    Label buttonLabel = new Label();
                    buttonLabel.Content = hostessInfo;
                    buttonStack.Children.Add(buttonLabel);

                    //Add button
                    m_Button_HostAndHostess_SelectedHostAndHostess.Content = buttonStack;

                    //Get buttons
                    StackPanel panel = (StackPanel)m_Button_HostAndHostess_SelectedHostAndHostess.Parent;
                    Button buttonCellPhone = null;
                    Button buttonEmail = null;
                    for (int iChild = 1; iChild < panel.Children.Count; ++iChild)
                    {
                        Button childButton = (Button)panel.Children[iChild];
                        if (childButton.Tag.ToString() == "CellPhone")
                        {
                            buttonCellPhone = childButton;
                            m_Button_HostAndHostess_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            buttonEmail = childButton;
                            m_Button_HostAndHostess_SelectedEmail = childButton;
                        }

                    }

                    //Modification of the cellphone button
                    buttonCellPhone.Content = null;
                    Image imgCellPhone = new Image();
                    imgCellPhone.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Image_CellPhone);
                    StackPanel stackCellPhonePanel = new StackPanel();
                    stackCellPhonePanel.Orientation = Orientation.Horizontal;
                    stackCellPhonePanel.Margin = new Thickness(5);
                    TextBlock cellPhoneInfo = new TextBlock();
                    string cellPhone = hostOrHostess.cellphone;
                    if (cellPhone.Length == 10)
                    {
                        cellPhone = cellPhone.Insert(8, " ");
                        cellPhone = cellPhone.Insert(6, " ");
                        cellPhone = cellPhone.Insert(4, " ");
                        cellPhone = cellPhone.Insert(2, " ");
                    }
                    string cellPhoneText = "   " + cellPhone;
                    cellPhoneInfo.Inlines.Add(cellPhoneText);
                    stackCellPhonePanel.Children.Add(imgCellPhone);
                    stackCellPhonePanel.Children.Add(cellPhoneInfo);
                    buttonCellPhone.Content = stackCellPhonePanel;
                    buttonCellPhone.Tag = "CellPhone";

                    //Modification of the email button
                    buttonEmail.Content = null;
                    Image imgEmail = new Image();
                    imgEmail.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Image_Email);
                    StackPanel stackEmailPanel = new StackPanel();
                    stackEmailPanel.Orientation = Orientation.Horizontal;
                    stackEmailPanel.Margin = new Thickness(5);
                    TextBlock EmailInfo = new TextBlock();
                    string EmailText = "   " + hostOrHostess.email;
                    EmailInfo.Inlines.Add(EmailText);
                    stackEmailPanel.Children.Add(imgEmail);
                    stackEmailPanel.Children.Add(EmailInfo);
                    buttonEmail.Content = stackEmailPanel;
                    buttonEmail.Tag = "Email";
                }

                return;
            }
            catch (Exception exception)
            {
                //Write the error into log
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Click on generate Hostess statement in excel format button
        /// </summary>
        private void Btn_HostAndHostess_GenerateExcelStatement_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the name of the pdf
                string today = m_Global_Handler.DateAndTime_Handler.Treat_Date(DateTime.Today.ToString("yyyy-MM-dd"), m_Global_Handler.Language_Handler);
                today = today.Replace("/", "-");
                string fileNameXLS = m_Global_Handler.Resources_Handler.Get_Resources("HostessStatement") + " - " + today;
                System.Windows.Forms.SaveFileDialog saveFile = new System.Windows.Forms.SaveFileDialog();
                saveFile.FileName = fileNameXLS;
                saveFile.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx;*.xls";
                if (saveFile.ShowDialog() != System.Windows.Forms.DialogResult.OK || saveFile.FileName.Length == 0)
                {
                    return;
                }
                else
                {
                    fileNameXLS = saveFile.FileName;
                }

                //Open the wait window
                windowWait.Start(m_Global_Handler, "HostessExcelGenerationPrincipalMessage", "HostessExcelGenerationSecondaryMessage");

                //Sort the collection
                SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                {
                    if (x.lastname == null && y.lastname == null) return 0;
                    else if (x.lastname == null) return 1;
                    else if (y.lastname == null) return -1;
                    else
                    {
                        return x.lastname.CompareTo(y.lastname);
                    }
                });

                //Write in a file with tab separated
                List<string> lines = new List<string>();
                string columnsHeader = m_Global_Handler.Resources_Handler.Get_Resources("HostessID") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("LastName") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("FirstName") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("BirthDate") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("IdPaycheck") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Phone") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CellPhone") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Address") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("ZipCode") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("City") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Country") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CreationDate");
                lines.Add(columnsHeader);
                for (int iHostess = 0; iHostess < SoftwareObjects.HostsAndHotessesCollection.Count; ++iHostess)
                {
                    Hostess hostess = SoftwareObjects.HostsAndHotessesCollection[iHostess];
                    string line = hostess.id + "\t" + hostess.lastname + "\t" + hostess.firstname + "\t" + hostess.birth_date + "\t" +
                        hostess.id_paycheck + "\t" + hostess.cellphone + "\t" + hostess.address + "\t" + hostess.zipcode + "\t" + hostess.city +
                        "\t" + hostess.country + "\t" + m_Global_Handler.DateAndTime_Handler.Treat_Date(hostess.date_creation, m_Global_Handler.Language_Handler);
                    lines.Add(line);
                }

                //Generate the excel file
                int result = m_Global_Handler.Excel_Handler.Generate_ExcelStatement(fileNameXLS, lines);
                if (result == -1)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissionsStatementExcelGenerationFailed"),
                                m_Global_Handler.Resources_Handler.Get_Resources("MissionsStatementExcelGenerationFailedCaption"),
                                MessageBoxButton.OK, MessageBoxImage.Error);
                    //Close the wait window
                    windowWait.Stop();

                    return;
                }

                //Action
                m_Global_Handler.Error_Handler.WriteAction("Hosts and hostesses statement generated to Excel file " + fileNameXLS);

                //Close the wait window
                windowWait.Stop();

                //Open the file
                Process.Start(fileNameXLS);
            }
            catch (Exception exception)
            {
                //Close the wait window
                windowWait.Stop();

                //Write the error into log
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Select a type of sort in the hostess combo box
        /// Sort in a(n) ascending/descending way the Hostess panels based on the selected item
        /// </summary>
        private void Cmb_HostAndHostess_SortBy_SelectionChanged(object sender, EventArgs e)
        {
            //Invert sorting
            m_IsSortHostess_Ascending = !m_IsSortHostess_Ascending;

            //Save in case of problems
            List<Hostess> SavedCollection = new List<Hostess>();
            SavedCollection = SoftwareObjects.HostsAndHotessesCollection;

            //No sort
            if (Cmb_HostAndHostess_SortBy.SelectedIndex == -1)
            {
                //Clear the grid
                Grid_HostAndHostess_Details.Children.Clear();
                Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                //Actualize the collection
                Actualize_GridHostsAndHostessesFromCollection();
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("LastNameHostess"))
            {
                //Sort by last name
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.lastname == null && y.lastname == null) return 0;
                        else if (x.lastname == null) return -1;
                        else if (y.lastname == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.lastname.CompareTo(y.lastname);
                            }
                            else
                            {
                                return y.lastname.CompareTo(x.lastname);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("Sex"))
            {
                //Sort by sex
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.sex == null && y.sex == null) return 0;
                        else if (x.sex == null) return -1;
                        else if (y.sex == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.sex.CompareTo(y.sex);
                            }
                            else
                            {
                                return y.sex.CompareTo(x.sex);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("ZipCode"))
            {
                //Sort by last name
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.zipcode == null && y.zipcode == null) return 0;
                        else if (x.zipcode == null) return -1;
                        else if (y.zipcode == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.zipcode.CompareTo(y.zipcode);
                            }
                            else
                            {
                                return y.zipcode.CompareTo(x.zipcode);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("City"))
            {
                //Sort by city
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.city == null && y.city == null) return 0;
                        else if (x.city == null) return -1;
                        else if (y.city == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.city.CompareTo(y.city);
                            }
                            else
                            {
                                return y.city.CompareTo(x.city);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("CreationDate"))
            {
                //Sort by creation date
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.date_creation == null && y.date_creation == null) return 0;
                        else if (x.date_creation == null) return -1;
                        else if (y.date_creation == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.date_creation.CompareTo(y.date_creation);
                            }
                            else
                            {
                                return y.date_creation.CompareTo(x.date_creation);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Auto generating columns for the datagrid containg the missions associated to the selected hostess
        /// </summary>
        private void DataGrid_HostAndHostessMissions_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            try
            {
                if (e.Column.Header.ToString() == "client")
                {
                    e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Customer");
                    e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                }
                else if (e.Column.Header.ToString() == "city")
                {
                    e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("City");
                    e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
                }
                else if (e.Column.Header.ToString() == "date")
                {
                    e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Date");
                    e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
                }
                else if (e.Column.Header.ToString() == "id")
                {
                    e.Column.Visibility = Visibility.Hidden;
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Double click on a row of the datagrid containing the bills
        /// Open the selected bill category
        /// </summary>
        private void DataGrid_HostAndHostessBilling_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (sender != null)
                {
                    DataGrid grid = sender as DataGrid;
                    if (grid != null && grid.SelectedItems != null && grid.SelectedItems.Count == 1)
                    {
                        //Get the hostess
                        m_DataGrid_HostAndHostess_Missions hostessGrid = (m_DataGrid_HostAndHostess_Missions)grid.SelectedItems[0];
                        Hostess Hostessel = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals((int)m_Button_HostAndHostess_SelectedHostAndHostess.Tag));
                        if (Hostessel == null)
                        {
                            return;
                        }

                        //Get selected bill in the datagrid
                        int indexSel = grid.SelectedIndex;
                        if (indexSel < 0 || indexSel > m_DataGrid_HostAndHostess_MissionsCollection.Count - 1)
                        {
                            return;
                        }
                        m_DataGrid_HostAndHostess_Missions billDatagrid = (m_DataGrid_HostAndHostess_Missions)grid.SelectedCells[0].Item;

                        //Get the missions associated to the selected host or hostess
                        List<Billing> missions = new List<Billing>();
                        for (int iBill = 0; iBill < SoftwareObjects.MissionsCollection.Count; ++iBill)
                        {
                            Billing bill = SoftwareObjects.MissionsCollection[iBill];
                            if (bill.id_HostAndHostess == Hostessel.id)
                            {
                                missions.Add(bill);
                            }
                        }
                        m_Grid_Details_Missions_MissionsCollection = missions;

                        //Get selected bill from collection
                        Billing billSel = SoftwareObjects.MissionsCollection.Find(x => x.id.Equals(billDatagrid.id));
                        if (billSel != null)
                        {
                            //Mission
                            if (billSel.num_invoice == "" || billSel.num_invoice == null)
                            {
                                //Graphism
                                Btn_Software_Missions.Background = m_Color_MainButton;
                                Btn_Software_HostAndHostess.Background = m_Color_SelectedMainButton;

                                //Visibility
                                Grid_HostAndHostess.Visibility = Visibility.Collapsed;
                                Grid_Missions.Visibility = Visibility.Visible;

                                //Archived or in progress
                                if (billSel.date_Mission_archived != null && billSel.date_Mission_archived != "")
                                {
                                    m_Mission_IsArchiveMode = true;
                                }

                                //Adding associated missions
                                Grid_Missions_Details.Children.Clear();
                                m_GridMissions_Column = 0;
                                m_GridMissions_Row = 0;
                                for (int iMission = 0; iMission < missions.Count; ++iMission)
                                {
                                    if (m_Mission_IsArchiveMode == true && missions[iMission].date_Mission_archived != null && missions[iMission].date_Mission_archived != "")
                                    {
                                        Add_MissionToGrid(missions[iMission]);
                                    }
                                    else if (m_Mission_IsArchiveMode == false && (missions[iMission].date_Mission_archived == "" || missions[iMission].date_Mission_archived == null))
                                    {
                                        Add_MissionToGrid(missions[iMission]);
                                    }
                                }

                                //Fields
                                Txt_Missions_FirstName.Text = Hostessel.firstname;
                                Txt_Missions_LastName.Text = Hostessel.lastname;

                                //Treat dates
                                Txt_Missions_CreationDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(billSel.date_Mission_creation, m_Global_Handler.Language_Handler);

                                //Select the mission
                                Filter_GridMissionsFromMissionsCollection(MissionStatus.NONE);
                                Select_Mission(billSel);

                                //Enable the buttons
                                Btn_Missions_Duplicate.IsEnabled = true;
                                Btn_Missions_Edit.IsEnabled = true;
                                Btn_Missions_Delete.IsEnabled = true;
                                Btn_Missions_Close.IsEnabled = true;
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Mouse down on the main hostess button event
        /// Select the hostess
        /// </summary>
        private void IsPressed_HostAndHostess_Button(object sender, RoutedEventArgs ev)
        {
            if (sender != null)
            {
                try
                {
                    //Get the hostess's id
                    Button buttonSel = (Button)sender;
                    string idSel = (string)buttonSel.Tag;
                    m_Id_SelectedHostAndHostess = idSel;

                    //Manage the buttons
                    if (m_Button_HostAndHostess_SelectedHostAndHostess != null)
                    {
                        m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_HostAndHostess_SelectedCellPhone != null)
                    {
                        m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_HostAndHostess_SelectedEmail != null)
                    {
                        m_Button_HostAndHostess_SelectedEmail.Background = m_Color_HostAndHostess;
                    }
                    m_Button_HostAndHostess_SelectedHostAndHostess = null;
                    m_Button_HostAndHostess_SelectedCellPhone = null;
                    m_Button_HostAndHostess_SelectedEmail = null;
                    StackPanel stackSel = (StackPanel)buttonSel.Parent;
                    for (int iChild = 0; iChild < stackSel.Children.Count; ++iChild)
                    {
                        Button childButton = (Button)stackSel.Children[iChild];
                        childButton.Background = m_Color_SelectedHostAndHostess;
                        if (iChild == 0)
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess = childButton;
                        }
                        else if (childButton.Tag.ToString() == "CellPhone")
                        {
                            m_Button_HostAndHostess_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_HostAndHostess_SelectedEmail = childButton;
                        }
                    }
                    //Select the hostess
                    Get_SelectedHostOrHostessFromButton();

                    //Enable the buttons
                    Btn_HostAndHostess_Edit.IsEnabled = true;
                    Btn_HostAndHostess_Delete.IsEnabled = true;
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Mouse down on the hostess's email button event
        /// Select hostess and open a new email with the selected hostess email adress
        /// </summary>
        private void IsPressed_HostAndHostess_EmailButton(object sender, RoutedEventArgs ev)
        {
            if (sender != null)
            {
                //Creation of the wait window
                WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

                try
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "HostessEmailPrincipalMessage", "HostessEmailSecondaryMessage");

                    //Get the email adress
                    Button emailButton = (Button)sender;
                    StackPanel emailStackPanel = (StackPanel)emailButton.Content;
                    TextBlock emailTextBox = (TextBlock)emailStackPanel.Children[1];
                    string emailAddress = emailTextBox.Text;

                    //Manage the buttons
                    if (m_Button_HostAndHostess_SelectedHostAndHostess != null)
                    {
                        m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_HostAndHostess_SelectedCellPhone != null)
                    {
                        m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_HostAndHostess_SelectedEmail != null)
                    {
                        m_Button_HostAndHostess_SelectedEmail.Background = m_Color_HostAndHostess;
                    }
                    m_Button_HostAndHostess_SelectedHostAndHostess = null;
                    m_Button_HostAndHostess_SelectedCellPhone = null;
                    m_Button_HostAndHostess_SelectedEmail = null;
                    StackPanel stackSel = (StackPanel)emailButton.Parent;
                    for (int iChildren = 0; iChildren < stackSel.Children.Count; ++iChildren)
                    {
                        Button childButton = (Button)stackSel.Children[iChildren];
                        childButton.Background = m_Color_SelectedHostAndHostess;
                        if (childButton.Tag.ToString() != "" && childButton.Tag.ToString() != "CellPhone" && childButton.Tag.ToString() != "Email")
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess = childButton;
                            m_Id_SelectedHostAndHostess = (string)childButton.Tag;
                        }
                        else if (childButton.Tag.ToString() == "CellPhone")
                        {
                            m_Button_HostAndHostess_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_HostAndHostess_SelectedEmail = childButton;
                        }
                    }

                    //Select the hostess
                    Get_SelectedHostOrHostessFromButton();

                    //Enable the buttons
                    Btn_HostAndHostess_Edit.IsEnabled = true;
                    Btn_HostAndHostess_Delete.IsEnabled = true;

                    //Open the default email sender
                    Process p = new Process();
                    string mailto = "mailto:" + emailAddress;
                    ProcessStartInfo ps = new ProcessStartInfo(mailto);
                    ps.CreateNoWindow = false;
                    ps.UseShellExecute = true;
                    p.StartInfo = ps;
                    p.Start();
                    p.WaitForExit();

                    //Action
                    m_Global_Handler.Error_Handler.WriteAction("Mail sent to " + emailAddress);

                    //Close stop window
                    windowWait.Stop();

                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    windowWait.Stop();
                    return;
                }
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Mouse down on the hostess's phone/cellphone button(s) event
        /// Select the hostess and open skype
        /// </summary>
        private void IsPressed_HostAndHostess_PhoneButton(object sender, RoutedEventArgs ev)
        {
            if (sender != null)
            {
                //Confirm skype call
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("ConfirmSkypeCall"),
                                m_Global_Handler.Resources_Handler.Get_Resources("ConfirmSkypeCallCaption"),
                                MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                //Creation of the wait window
                WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

                try
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "HostessPhonePrincipalMessage", "HostessPhoneSecondaryMessage");

                    //Get the email adress
                    Button phoneButton = (Button)sender;
                    StackPanel phoneStackPanel = (StackPanel)phoneButton.Content;
                    TextBlock phoneTextBox = (TextBlock)phoneStackPanel.Children[1];
                    string phoneNumber = phoneTextBox.Text.Trim();

                    //Manage the buttons
                    if (m_Button_HostAndHostess_SelectedHostAndHostess != null)
                    {
                        m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_HostAndHostess_SelectedCellPhone != null)
                    {
                        m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_HostAndHostess_SelectedEmail != null)
                    {
                        m_Button_HostAndHostess_SelectedEmail.Background = m_Color_HostAndHostess;
                    }
                    m_Button_HostAndHostess_SelectedHostAndHostess = null;
                    m_Button_HostAndHostess_SelectedCellPhone = null;
                    m_Button_HostAndHostess_SelectedEmail = null;
                    StackPanel stackSel = (StackPanel)phoneButton.Parent;
                    for (int iChildren = 0; iChildren < stackSel.Children.Count; ++iChildren)
                    {
                        Button childButton = (Button)stackSel.Children[iChildren];
                        childButton.Background = m_Color_SelectedHostAndHostess;
                        if (childButton.Tag.ToString() != "" && childButton.Tag.ToString() != "CellPhone" && childButton.Tag.ToString() != "Phone" && childButton.Tag.ToString() != "Email")
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess = childButton;
                            m_Id_SelectedHostAndHostess = (string)childButton.Tag;
                        }
                        else if (childButton.Tag.ToString() == "CellPhone")
                        {
                            m_Button_HostAndHostess_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_HostAndHostess_SelectedEmail = childButton;
                        }
                    }

                    //Select the hostess
                    Hostess hostOrHostess = Get_SelectedHostOrHostessFromButton();

                    //Open skype
                    SKYPE4COMLib.Skype skype = new SKYPE4COMLib.Skype();
                    if (skype != null)
                    {
                        if (!skype.Client.IsRunning)
                        {
                            skype.Client.Start(true, true);
                        }
                        //skype.Attach(8, true);
                        //skype.Client.OpenDialpadTab();
                        //SKYPE4COMLib.Call call = skype.PlaceCall(phoneNumber);
                    }

                    //Enable the buttons
                    Btn_HostAndHostess_Edit.IsEnabled = true;
                    Btn_HostAndHostess_Delete.IsEnabled = true;

                    //Action
                    m_Global_Handler.Error_Handler.WriteAction("Skype call made to " + hostOrHostess.firstname + " " + hostOrHostess.lastname);

                    //Close stop window
                    windowWait.Stop();

                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    windowWait.Stop();
                    return;
                }
            }
        }

        /// <summary>
        /// Boolean tagging the full actualization avoiding to do it multiple time
        /// This variable is only used in the research hostess text box text changed event
        /// </summary>
        bool actualizationHostessDone = true;
        /// <summary>
        /// Event
        /// Hostess
        /// Text changed in the research hostess text box
        /// Only Hostess containing the text (at least 2 characters) are shown 
        /// </summary>
        private void Txt_HostAndHostess_Research_TextChanged(object sender, TextChangedEventArgs e)
        {
            //Get the researched text
            string researchedText = Txt_HostAndHostess_Research.Text;
            researchedText = researchedText.ToLower();

            //Clear the fields
            m_Button_HostAndHostess_SelectedHostAndHostess = null;

            //Verifications - 3 letters minimum
            if (researchedText.Length < 2)
            {
                if (!actualizationHostessDone)
                {
                    Grid_HostAndHostess_Details.Children.Clear();
                    Actualize_GridHostsAndHostessesFromDatabase();
                    actualizationHostessDone = true;
                    return;
                }
            }
            else
            {
                try
                {
                    //Research in each private member of the list of hostess
                    List<Hostess> foundHostessList = new List<Hostess>();
                    for (int iHostess = 0; iHostess < SoftwareObjects.HostsAndHotessesCollection.Count; ++iHostess)
                    {
                        Hostess processedHostess = SoftwareObjects.HostsAndHotessesCollection[iHostess];
                        if (processedHostess.address.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.city.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.id_paycheck.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.country.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.email.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.firstname.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.lastname.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.state.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.zipcode.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.date_creation.ToLower().Contains(researchedText))
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                        else if (processedHostess.birth_date.ToLower().Contains(researchedText) && researchedText.Length >= 4)
                        {
                            foundHostessList.Add(processedHostess);
                            continue;
                        }
                    }

                    //Displaying the found Hostess list
                    Grid_HostAndHostess_Details.Children.Clear();
                    actualizationHostessDone = false;
                    if (foundHostessList.Count > 0)
                    {
                        //Clear fields

                        //Initialize counter for columns and rows of Grid_HostAndHostess
                        m_GridHostess_Column = 0;
                        m_GridHostess_Row = 0;
                    }
                    for (int iHostess = 0; iHostess < foundHostessList.Count; ++iHostess)
                    {
                        Add_HostOrHostessToGrid(foundHostessList[iHostess]);
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }

        }

        #endregion

        #region Clients

        /// <summary>
        /// Event
        /// Clients
        /// Click on add client
        /// </summary>
        private void Btn_Clients_Create_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on delete client
        /// </summary>
        private void Btn_Clients_Delete_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on edit client
        /// </summary>
        private void Btn_Clients_Edit_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on generate an excel staement of clients
        /// /// </summary>
        private void Btn_Clients_GenerateExcelStatement_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// Event
        /// Clients
        /// Select a type of sort in the clients combo box
        /// Sort in a(n) ascending/descending way the clients panels based on the selected item
        /// </summary>
        private void Cmb_Clients_SortBy_SelectionChanged(object sender, EventArgs e)
        {
            //Invert sorting
            m_IsSortHostess_Ascending = !m_IsSortHostess_Ascending;

            //Save in case of problems
            List<Hostess> SavedCollection = new List<Hostess>();
            SavedCollection = SoftwareObjects.HostsAndHotessesCollection;

            //No sort
            if (Cmb_HostAndHostess_SortBy.SelectedIndex == -1)
            {
                //Clear the grid
                Grid_HostAndHostess_Details.Children.Clear();
                Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                //Actualize the collection
                Actualize_GridHostsAndHostessesFromCollection();
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("LastNameHostess"))
            {
                //Sort by last name
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.lastname == null && y.lastname == null) return 0;
                        else if (x.lastname == null) return -1;
                        else if (y.lastname == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.lastname.CompareTo(y.lastname);
                            }
                            else
                            {
                                return y.lastname.CompareTo(x.lastname);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("Sex"))
            {
                //Sort by sex
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.sex == null && y.sex == null) return 0;
                        else if (x.sex == null) return -1;
                        else if (y.sex == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.sex.CompareTo(y.sex);
                            }
                            else
                            {
                                return y.sex.CompareTo(x.sex);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("ZipCode"))
            {
                //Sort by last name
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.zipcode == null && y.zipcode == null) return 0;
                        else if (x.zipcode == null) return -1;
                        else if (y.zipcode == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.zipcode.CompareTo(y.zipcode);
                            }
                            else
                            {
                                return y.zipcode.CompareTo(x.zipcode);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("City"))
            {
                //Sort by city
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.city == null && y.city == null) return 0;
                        else if (x.city == null) return -1;
                        else if (y.city == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.city.CompareTo(y.city);
                            }
                            else
                            {
                                return y.city.CompareTo(x.city);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("CreationDate"))
            {
                //Sort by creation date
                try
                {
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.date_creation == null && y.date_creation == null) return 0;
                        else if (x.date_creation == null) return -1;
                        else if (y.date_creation == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.date_creation.CompareTo(y.date_creation);
                            }
                            else
                            {
                                return y.date_creation.CompareTo(x.date_creation);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = SavedCollection;

                    //Clear the grid
                    Grid_HostAndHostess_Details.Children.Clear();
                    Grid_HostAndHostess_Details.RowDefinitions.RemoveRange(0, Grid_HostAndHostess_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridHostsAndHostessesFromCollection();

                    return;
                }
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Auto generating columns for the datagrid containg the missions associated to the selected client
        /// </summary>
        private void DataGrid_ClientsMissions_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            try
            {
                if (e.Column.Header.ToString() == "description")
                {
                    e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Description");
                    e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                }
                else if (e.Column.Header.ToString() == "city")
                {
                    e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("City");
                    e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
                }
                else if (e.Column.Header.ToString() == "date")
                {
                    e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Date");
                    e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
                }
                else if (e.Column.Header.ToString() == "id")
                {
                    e.Column.Visibility = Visibility.Hidden;
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }


        #endregion


        #region Settings

        #region General

        /// <summary>
        /// Event
        /// Settings
        /// General
        /// Choose databse location
        /// </summary>
        private void Btn_Settings_General_Database_Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Confirm the modification
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("DatabaseModificationConfirmationText"),
                    m_Global_Handler.Resources_Handler.Get_Resources("DatabaseModificationConfirmationCaption"),
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                //Test database
                try
                {
                    MySqlConnection conn = new MySqlConnection(Txt_Settings_General_Database.Text);
                    conn.Open();
                }
                catch (ArgumentException Arg_Ex)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, Arg_Ex);
                    MessageBox.Show(Arg_Ex.Message, m_Global_Handler.Resources_Handler.Get_Resources("DatabaseDefinitionModificationErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                    Txt_Settings_General_Database.Text = SoftwareObjects.GlobalSettings.database_definition;
                    return;
                }
                catch (MySqlException SQL_Ex)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, SQL_Ex);
                    MessageBox.Show(SQL_Ex.Message, m_Global_Handler.Resources_Handler.Get_Resources("DatabaseDefinitionModificationErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                    Txt_Settings_General_Database.Text = SoftwareObjects.GlobalSettings.database_definition;
                    return;
                }

                //Save database definition
                SoftwareObjects.GlobalSettings.database_definition = Txt_Settings_General_Database.Text;

                //Save infos
                string settingsFilePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Resources\\Settings.zad";
                string uriPath = settingsFilePath;
                string localPath = new Uri(uriPath).LocalPath;
                using (StreamWriter file = new StreamWriter(localPath, true))
                {
                    file.WriteLine(SoftwareObjects.GlobalSettings.database_definition);
                }

                //Action
                m_Global_Handler.Error_Handler.WriteAction("Database " + SoftwareObjects.GlobalSettings.database_definition + " saved");
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Settings
        /// General
        /// Choose Photos location
        /// </summary>
        private void Btn_Settings_General_Photos_Choose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Open folder browser
                using (var folderBrowser = new System.Windows.Forms.FolderBrowserDialog())
                {
                    System.Windows.Forms.DialogResult result = folderBrowser.ShowDialog();
                    //Get folder
                    if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowser.SelectedPath))
                    {
                        //Save to database
                        if (SoftwareObjects.GlobalSettings.photos_repository != folderBrowser.SelectedPath)
                        {
                            string res = m_Database_Handler.Edit_Settings_General_PhotosDirectory(folderBrowser.SelectedPath);
                            if (res.Contains("OK"))
                            {
                                SoftwareObjects.GlobalSettings.photos_repository = folderBrowser.SelectedPath;
                                Txt_Settings_General_Photos.Text = SoftwareObjects.GlobalSettings.photos_repository;

                                //Action
                                m_Global_Handler.Error_Handler.WriteAction("Phots saved to " + SoftwareObjects.GlobalSettings.photos_repository);
                            }
                            else
                            {
                                MessageBox.Show(res, m_Global_Handler.Resources_Handler.Get_Resources("PhotosDirectoryModificationErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        #endregion

        #endregion

        #endregion

        #region Functions

        #region Common

        /// <summary>
        /// Functions
        /// Common
        /// Actualize the collection of bills from the collection saved in the database
        /// </summary>
        private void Fill_BillsCollectionFromDatabase()
        {
            if (m_Database_Handler != null)
            {
                try
                {
                    //Clear collection
                    SoftwareObjects.MissionsCollection.Clear();

                    //In progress bills
                    string bills = m_Database_Handler.Get_MissionsFromDatabaseDatabase(false);
                    if (bills.Contains("error"))
                    {
                        Exception error = new Exception(bills);
                        throw error;
                    }
                    else if (bills != "")
                    {
                        List<Billing> billsCollection = m_Database_Handler.Deserialize_JSON<List<Billing>>(bills);
                        for (int iBill = 0; iBill < billsCollection.Count; ++iBill)
                        {
                            //Get mission from collection
                            Billing bill = billsCollection[iBill];
                            //Add to the collection
                            SoftwareObjects.MissionsCollection.Add(bill);
                        }
                    }

                    //Archived bills
                    string archivedBills = m_Database_Handler.Get_MissionsFromDatabaseDatabase(true);
                    if (archivedBills.Contains("error"))
                    {
                        Exception error = new Exception(archivedBills);
                        throw error;
                    }
                    else if (archivedBills != "")
                    {
                        List<Billing> archivedBillsCollection = m_Database_Handler.Deserialize_JSON<List<Billing>>(archivedBills);
                        for (int iBill = 0; iBill < archivedBillsCollection.Count; ++iBill)
                        {
                            //Get mission from collection
                            Billing archivedBill = archivedBillsCollection[iBill];
                            //Add to the collection
                            SoftwareObjects.MissionsCollection.Add(archivedBill);
                        }
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        /// <summary>
        /// Functions
        /// Common
        /// Prevent type letter in the text box
        /// </summary>
        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            string pattern = "[0-9]|[,.]";
            Regex regex = new Regex(pattern);
            e.Handled = !regex.IsMatch(e.Text);
        }

        /// <summary>
        /// Functions
        /// Common
        /// Prevent special characters in the text box
        /// </summary>
        private void NoSpecialCharactersValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Char keyChar = (Char)System.Text.Encoding.ASCII.GetBytes(e.Text)[0];
            if (char.IsLetter(keyChar))
            {
                // char is letter
                e.Handled = false;
            }
            else if (char.IsDigit(keyChar))
            {
                // char is digit
                e.Handled = false;
            }
            else
            {
                // char is neither letter or digit.
                if (keyChar != '%' && keyChar != '_' && keyChar != '-')
                {
                    e.Handled = true;
                }
                else
                {
                    e.Handled = false;
                }
            }
        }

        #endregion

        #region Hosts and hostesses

        /// <summary>
        /// Functions
        /// Host and hostess
        /// Actualize the grid of Hostess from the actual collection
        /// </summary>
        private void Actualize_GridHostsAndHostessesFromCollection()
        {
            if (m_Database_Handler != null)
            {
                //Initialize rows and columns
                m_GridHostess_Column = 0;
                m_GridHostess_Row = 0;

                //Actualization
                try
                {
                    //Sort by last name
                    SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                    {
                        if (x.lastname == null && y.lastname == null) return 0;
                        else if (x.lastname == null) return -1;
                        else if (y.lastname == null) return 1;
                        else
                        {
                            if (m_IsSortHostess_Ascending == true)
                            {
                                return x.lastname.CompareTo(y.lastname);
                            }
                            else
                            {
                                return y.lastname.CompareTo(x.lastname);
                            }
                        }
                    });
                    //Add to grid
                    for (int iHostess = 0; iHostess < SoftwareObjects.HostsAndHotessesCollection.Count; ++iHostess)
                    {
                        Hostess hostess = SoftwareObjects.HostsAndHotessesCollection[iHostess];
                        Add_HostOrHostessToGrid(hostess);
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = new List<Hostess>();
                    return;
                }
            }
        }

        /// <summary>
        /// Functions
        /// Host and hostess
        /// Actualize the grid of hosts and hostesses from the collection saved in the database
        /// </summary>
        private void Actualize_GridHostsAndHostessesFromDatabase()
        {
            if (m_Database_Handler != null)
            {
                //Initialize rows and columns
                m_GridHostess_Column = 0;
                m_GridHostess_Row = 0;

                //Getting the hosts and hostesses collection  
                try
                {
                    string res = m_Database_Handler.Get_HostsAndHostessesFromDatabase();
                    if (res.Contains("OK"))
                    {
                        //Sort by last name
                        SoftwareObjects.HostsAndHotessesCollection.Sort(delegate (Hostess x, Hostess y)
                        {
                            if (x.lastname == null && y.lastname == null) return 0;
                            else if (x.lastname == null) return -1;
                            else if (y.lastname == null) return 1;
                            else
                            {
                                if (m_IsSortHostess_Ascending == true)
                                {
                                    return x.lastname.CompareTo(y.lastname);
                                }
                                else
                                {
                                    return y.lastname.CompareTo(x.lastname);
                                }
                            }
                        });
                        //Add to grid
                        for (int iHostess = 0; iHostess < SoftwareObjects.HostsAndHotessesCollection.Count; ++iHostess)
                        {
                            Hostess hostess = SoftwareObjects.HostsAndHotessesCollection[iHostess];
                            Add_HostOrHostessToGrid(hostess);
                        }
                    }
                    else if (res.Contains("Error"))
                    {
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("HostsAndHostessesActualizationErrorCaption"),
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        SoftwareObjects.HostsAndHotessesCollection = new List<Hostess>();
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.HostsAndHotessesCollection = new List<Hostess>();
                    return;
                }
            }
        }

        /// <summary>
        /// Functions
        /// Host and hostess
        /// Increment on grid Hostess columns
        /// </summary>
        private int m_GridHostess_Column = 0;
        /// <summary>
        /// Functions
        /// Host and hostess
        /// Increment on grid Hostess rows
        /// </summary>
        private int m_GridHostess_Row = 0;
        /// <summary>
        /// Functions
        /// Host and hostess
        /// Add a hostess to the grid, create a stack panel containing all the information of the hostess
        /// <param name="_HostOrHostess">Hostess</param>
        /// </summary>
        private void Add_HostOrHostessToGrid(Hostess _HostOrHostess)
        {
            try
            {
                //Treatment of the first row
                Grid_HostAndHostess_Details.RowDefinitions[0].MinHeight = 350;

                //Create the hostess button
                Button buttonHostess = new Button();
                buttonHostess.Padding = new Thickness(5);
                buttonHostess.Background = m_Color_HostAndHostess;
                buttonHostess.Click += IsPressed_HostAndHostess_Button;

                //Treatment of hostess info
                TextBlock hostessInfo = new TextBlock();
                Run names = new Run("\n" + _HostOrHostess.firstname + " " + _HostOrHostess.lastname);
                names.FontSize = 15;
                names.FontWeight = FontWeights.Bold;
                hostessInfo.Inlines.Add(names);
                hostessInfo.Inlines.Add(new LineBreak());
                string ageInfo = "";
                if (_HostOrHostess.birth_date.Split(' ').Length == 3)
                {
                    ageInfo = _HostOrHostess.birth_date;
                    int age = Age(_HostOrHostess.birth_date);
                    if (age != -1)
                    {
                        ageInfo += ", " + age.ToString() + " " + m_Global_Handler.Resources_Handler.Get_Resources("YearsOld");
                    }
                }
                string creationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_HostOrHostess.date_creation, m_Global_Handler.Language_Handler);
                string info = _HostOrHostess.address + "\n" +
                    _HostOrHostess.zipcode + ", " + _HostOrHostess.city + "\n" +
                    ageInfo + "\n" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CreationDate") + " : " + creationDate + "\n";
                hostessInfo.Inlines.Add(info);
                //Create stack button
                StackPanel buttonStack = new StackPanel();
                buttonStack.Orientation = Orientation.Horizontal;
                //Create image
                Image imgHostess = new Image();
                string repository = SoftwareObjects.GlobalSettings.photos_repository + "\\" + _HostOrHostess.id;
                if (Directory.Exists(repository))
                {
                    //Get photos
                    string[] files = Directory.GetFiles(repository);
                    //Display the first photo
                    if (files.Length > 0)
                    {
                        string photoFilename = files[0];
                        System.Drawing.Bitmap photoBmp = m_Global_Handler.Image_Handler.Load_Bitmap(photoFilename);
                        double factor = 100.0 / photoBmp.Height;
                        System.Drawing.Bitmap tmp = (System.Drawing.Bitmap)m_Global_Handler.Image_Handler.ResizeImage(photoBmp, (int)(factor * photoBmp.Width), (int)(factor * photoBmp.Height));
                        imgHostess.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(tmp);
                    }
                }
                //Effect
                System.Windows.Media.Effects.DropShadowEffect dropShadowEffect = new System.Windows.Media.Effects.DropShadowEffect();
                Color shadowColor = new Color();
                shadowColor.ScA = 1;
                shadowColor.ScB = 0;
                shadowColor.ScG = 0;
                shadowColor.ScR = 0;
                dropShadowEffect.Color = shadowColor;
                dropShadowEffect.Direction = 320;
                dropShadowEffect.ShadowDepth = 10;
                dropShadowEffect.BlurRadius = 10;
                dropShadowEffect.Opacity = 0.5;
                imgHostess.Effect = dropShadowEffect;
                //Add image
                buttonStack.Children.Add(imgHostess);

                //Add void
                Label voidLabel = new Label();
                voidLabel.Content = "\t";
                buttonStack.Children.Add(voidLabel);

                //Add label
                Label buttonLabel = new Label();
                buttonLabel.Content = hostessInfo;
                buttonStack.Children.Add(buttonLabel);

                //Add button
                buttonHostess.Content = buttonStack;
                buttonHostess.Tag = _HostOrHostess.id;

                //Create the main stack panel
                StackPanel mainStackPanel = new StackPanel();
                mainStackPanel.MinHeight = 45;
                mainStackPanel.MinWidth = 145;
                mainStackPanel.Margin = new Thickness(5);
                mainStackPanel.Tag = _HostOrHostess.id;
                mainStackPanel.Children.Add(buttonHostess);

                //Create the cellphone call button
                Button buttonCallCellPhone = new Button();
                buttonCallCellPhone.MinHeight = 20;
                buttonCallCellPhone.Padding = new Thickness(5);
                buttonCallCellPhone.Background = m_Color_HostAndHostess;
                buttonCallCellPhone.Click += IsPressed_HostAndHostess_PhoneButton;
                buttonCallCellPhone.Tag = "CellPhone";
                //Treatment of the image in the call button
                Image imgCellPhone = new Image();
                imgCellPhone.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Image_CellPhone);
                StackPanel stackCellPhonePanel = new StackPanel();
                stackCellPhonePanel.Orientation = Orientation.Horizontal;
                stackCellPhonePanel.Margin = new Thickness(5);
                TextBlock cellPhoneInfo = new TextBlock();
                string cellPhone = _HostOrHostess.cellphone;
                if (cellPhone.Length == 10)
                {
                    cellPhone = cellPhone.Insert(8, " ");
                    cellPhone = cellPhone.Insert(6, " ");
                    cellPhone = cellPhone.Insert(4, " ");
                    cellPhone = cellPhone.Insert(2, " ");
                }
                string cellPhoneText = "   " + cellPhone;
                cellPhoneInfo.Inlines.Add(cellPhoneText);
                stackCellPhonePanel.Children.Add(imgCellPhone);
                stackCellPhonePanel.Children.Add(cellPhoneInfo);
                buttonCallCellPhone.Content = stackCellPhonePanel;
                //Add to the main stack panel
                mainStackPanel.Children.Add(buttonCallCellPhone);

                //Create the email button
                Button buttonSendMail = new Button();
                buttonSendMail.MinHeight = 20;
                buttonSendMail.Padding = new Thickness(5);
                buttonSendMail.Background = m_Color_HostAndHostess;
                buttonSendMail.Click += IsPressed_HostAndHostess_EmailButton;
                buttonSendMail.Tag = "Email";
                //Treatment of the image in the email button
                Image imgEmail = new Image();
                imgEmail.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Image_Email);
                StackPanel stackEmail = new StackPanel();
                stackEmail.Orientation = Orientation.Horizontal;
                stackEmail.Margin = new Thickness(5);
                TextBlock emailInfo = new TextBlock();
                string emailText = "   " + _HostOrHostess.email;
                emailInfo.Inlines.Add(emailText);
                stackEmail.Children.Add(imgEmail);
                stackEmail.Children.Add(emailInfo);
                buttonSendMail.Content = stackEmail;
                //Add to the main stack panel
                mainStackPanel.Children.Add(buttonSendMail);

                //Treatment of the increments
                if (m_GridHostess_Column > Grid_HostAndHostess_Details.ColumnDefinitions.Count - 1)
                {
                    m_GridHostess_Column = 0;
                    m_GridHostess_Row += 1;
                    RowDefinition row = new RowDefinition();
                    row.MinHeight = 350;
                    Grid_HostAndHostess_Details.RowDefinitions.Add(row);
                    Grid.SetColumn(mainStackPanel, m_GridHostess_Column);
                    Grid.SetRow(mainStackPanel, m_GridHostess_Row);
                    m_GridHostess_Column += 1;
                }
                else
                {
                    Grid.SetColumn(mainStackPanel, m_GridHostess_Column);
                    Grid.SetRow(mainStackPanel, m_GridHostess_Row);
                    m_GridHostess_Column += 1;
                }

                //Treatment of the Grid_HostAndHostess column
                Grid_HostAndHostess_Details.Children.Add(mainStackPanel);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Host and hostess
        /// Get a hostess from his id
        /// <returns>Returns the hostess corresponding to the id</returns>
        /// </summary>
        private Hostess Get_SelectedHostOrHostessFromButton()
        {
            //Test
            if (m_Button_HostAndHostess_SelectedHostAndHostess == null || m_Button_HostAndHostess_SelectedHostAndHostess.Tag == null)
            {
                MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("HostessSelectedError"),
                    m_Global_Handler.Resources_Handler.Get_Resources("HostessSelectedErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            //Get the hostess from the id
            string id = (string)m_Button_HostAndHostess_SelectedHostAndHostess.Tag;
            Hostess hostess = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(id));

            //Actualize datagrid
            if (hostess != null)
            {
                //Fill the m_DataGrid_HostAndHostess_MissionsCollection
                m_DataGrid_HostAndHostess_MissionsCollection.Clear();
                for (int iBill = 0; iBill < SoftwareObjects.MissionsCollection.Count; ++iBill)
                {
                    Billing bill = SoftwareObjects.MissionsCollection[iBill];
                    if (bill.id_HostAndHostess == hostess.id)
                    {
                        string type = "";
                        if (bill.num_invoice == "" || bill.num_invoice == null)
                        {
                            type = m_Global_Handler.Resources_Handler.Get_Resources("Mission");
                            if (bill.date_Mission_archived != null && bill.date_Mission_archived != "")
                            {
                                //Not show archived mission
                                continue;
                            }
                        }
                        else
                        {
                            type = m_Global_Handler.Resources_Handler.Get_Resources("Invoice");
                            if (bill.date_invoice_archived != null && bill.date_invoice_archived != "")
                            {
                                //Not show archived invoice
                                continue;
                            }
                        }
                        string creationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(bill.date_Mission_creation, m_Global_Handler.Language_Handler);
                        string sentDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(bill.date_Mission_sent, m_Global_Handler.Language_Handler);
                        m_DataGrid_HostAndHostess_Missions data = new m_DataGrid_HostAndHostess_Missions("", "", "", "");
                        m_DataGrid_HostAndHostess_MissionsCollection.Add(data);
                    }
                }
                DataGrid_HostAndHostess_Missions.Items.Refresh();
            }

            //Return the hostess
            return hostess;
        }


        /// <summary>
        /// Functions
        /// Host and hostess
        /// Compute age from birth date
        /// <returns>Returns the hostess corresponding to the id</returns>
        /// </summary>
        private int Age(string _BirthDate)
        {
            try
            {
                string[] birthDate = _BirthDate.Split(' ');
                int day = int.Parse(birthDate[0]);
                string monthStr = birthDate[1];
                int month = DateTime.ParseExact(monthStr, "MMMM", CultureInfo.CurrentCulture).Month;
                int year = int.Parse(birthDate[2]);
                DateTime birthDateTime = new DateTime(year, 1, day);
                return DateTime.Now.Year - birthDateTime.Year -
                         (DateTime.Now.Month < birthDateTime.Month ? 1 :
                         DateTime.Now.Day < birthDateTime.Day ? 1 : 0);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return -1;
            }
        }
        #endregion

        #region Missions

        /// <summary>
        /// Functions
        /// Missions
        /// Actualize the grid of missions from the actual billing collection
        /// </summary>
        private void Actualize_GridMissionsFromMissionsCollection()
        {
            if (m_Database_Handler != null)
            {
                //Initialize rows and columns
                m_GridMissions_Column = 0;
                m_GridMissions_Row = 0;

                //Actualization
                try
                {
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);
                    for (int iMission = 0; iMission < m_Grid_Details_Missions_MissionsCollection.Count; ++iMission)
                    {
                        //Get mission
                        Billing mission = m_Grid_Details_Missions_MissionsCollection[iMission];
                        if (mission.date_Mission_archived == "0000-00-00")
                        {
                            mission.date_Mission_archived = null;
                        }
                        //Show archive or not
                        if (m_Mission_IsArchiveMode == true && mission.date_Mission_archived != null)
                        {
                            Add_MissionToGrid(mission);
                        }
                        if (m_Mission_IsArchiveMode == false && mission.date_Mission_archived == null)
                        {
                            Add_MissionToGrid(mission);
                        }
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Actualize the grid of missions from the collection saved in the database
        /// </summary>
        private void Actualize_GridMissionsFromDatabase()
        {
            if (m_Database_Handler != null)
            {
                //Fill bills collection
                Fill_BillsCollectionFromDatabase();

                //Initialize
                m_GridMissions_Column = 0;
                m_GridMissions_Row = 0;
                Grid_Missions_Details.Children.Clear();
                m_Grid_Details_Missions_MissionsCollection.Clear();

                //Getting the missions collection
                try
                {
                    string missions = m_Database_Handler.Get_MissionsFromDatabaseDatabase(m_Mission_IsArchiveMode);
                    if (missions.Contains("error"))
                    {
                        Exception error = new Exception(missions);
                        throw error;
                    }
                    else if (missions != "")
                    {
                        Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);
                        List<Billing> missionsCollection = m_Database_Handler.Deserialize_JSON<List<Billing>>(missions);
                        if (missionsCollection.Count == 0)
                        {
                            Exception error = new Exception("Error in getting the missions collection : string != \"\" and empty collection - Verify the string sent by the website");
                            throw error;
                        }
                        else
                        {
                            for (int iMission = 0; iMission < missionsCollection.Count; ++iMission)
                            {
                                //Get mission from collection
                                Billing mission = missionsCollection[iMission];
                                //Verify date archived
                                if (mission.date_Mission_archived == "0000-00-00")
                                {
                                    mission.date_Mission_archived = null;
                                }
                                //Show archive or not
                                if (m_Mission_IsArchiveMode == true && mission.date_Mission_archived != null)
                                {
                                    Add_MissionToGrid(mission);
                                    m_Grid_Details_Missions_MissionsCollection.Add(mission);
                                }
                                else if (m_Mission_IsArchiveMode == false && mission.date_Mission_archived == null)
                                {
                                    Add_MissionToGrid(mission);
                                    m_Grid_Details_Missions_MissionsCollection.Add(mission);
                                }
                            }
                        }

                        //Sort by status
                        Sort_MissionsByStatus(m_Grid_Details_Missions_MissionsCollection);
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Increment on grid missions columns
        /// </summary>
        private int m_GridMissions_Column = 0;
        /// <summary>
        /// Functions
        /// Missions
        /// Increment on grid missions rows
        /// </summary>
        private int m_GridMissions_Row = 0;
        /// <summary>
        /// Functions
        /// Missions
        /// Add an mission to the grid, create a stack panel containing all the information of the mission
        /// <param name="_Mission">Mission</param>
        /// </summary>
        private void Add_MissionToGrid(Billing _Mission)
        {
            try
            {
                //Treatment of the first row
                Grid_Missions_Details.RowDefinitions[0].MinHeight = 350;

                //Create the buttons
                Button buttonMission = new Button();
                Button buttonMissionsendMail = new Button();
                Manage_MissionButton(_Mission, buttonMission, buttonMissionsendMail, true, false);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Filter the grid of missions from the actual billing collection
        /// </summary>
        private void Filter_GridMissionsFromMissionsCollection(MissionStatus _Status)
        {
            try
            {
                //Get selected mission
                Billing missionSel = Get_SelectedMissionFromButton(false);

                //Initialize rows and columns
                m_GridMissions_Column = 0;
                m_GridMissions_Row = 0;

                //Initialize status
                m_Mission_SelectedStatus = _Status;

                //Get status
                List<Billing> createdMissions = new List<Billing>();
                List<Billing> generatedMissions = new List<Billing>();
                List<Billing> sentMissions = new List<Billing>();
                List<Billing> acceptedMissions = new List<Billing>();
                List<Billing> declinedMissions = new List<Billing>();
                List<Billing> billedMissions = new List<Billing>();
                for (int iMission = 0; iMission < m_Grid_Details_Missions_MissionsCollection.Count; ++iMission)
                {
                    Billing mission = m_Grid_Details_Missions_MissionsCollection[iMission];
                    if (mission.date_invoice_creation != "" && mission.date_invoice_creation != null && mission.date_invoice_creation != "0000-00-00")
                    {
                        billedMissions.Add(mission);
                    }
                    else if (mission.date_Mission_accepted != "" && mission.date_Mission_accepted != null && mission.date_Mission_accepted != "0000-00-00")
                    {
                        acceptedMissions.Add(mission);
                    }
                    else if (mission.date_Mission_declined != "" && mission.date_Mission_declined != null && mission.date_Mission_declined != "0000-00-00")
                    {
                        declinedMissions.Add(mission);
                    }
                    else if (mission.date_Mission_sent != "" && mission.date_Mission_sent != null && mission.date_Mission_sent != "0000-00-00")
                    {
                        sentMissions.Add(mission);
                    }
                    else if (mission.date_Mission_generated != "" && mission.date_Mission_generated != null && mission.date_Mission_generated != "0000-00-00")
                    {
                        generatedMissions.Add(mission);
                    }
                    else
                    {
                        createdMissions.Add(mission);
                    }
                }

                //Buttons
                Btn_Missions_Legend_Mission_Done.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Billed.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Created.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Declined.Background = m_Color_Button;

                //Get filtered collection
                List<Billing> filteredMissions = new List<Billing>();
                if (_Status == MissionStatus.CREATED)
                {
                    Btn_Missions_Legend_Mission_Created.Background = m_Color_SelectedMission;
                    filteredMissions = createdMissions;
                }
                else if (_Status == MissionStatus.DECLINED)
                {
                    Btn_Missions_Legend_Mission_Declined.Background = m_Color_SelectedMission;
                    filteredMissions = declinedMissions;
                }
                else if (_Status == MissionStatus.DONE)
                {
                    Btn_Missions_Legend_Mission_Done.Background = m_Color_SelectedMission;
                    filteredMissions = acceptedMissions;
                }
                else if (_Status == MissionStatus.BILLED)
                {
                    Btn_Missions_Legend_Mission_Billed.Background = m_Color_SelectedMission;
                    filteredMissions = billedMissions;
                }
                else if (_Status == MissionStatus.NONE)
                {
                    filteredMissions = m_Grid_Details_Missions_MissionsCollection;
                }

                //Filter
                Grid_Missions_Details.Children.Clear();
                Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);
                Billing missionFound = null;
                for (int iMission = 0; iMission < filteredMissions.Count; ++iMission)
                {
                    //Get mission
                    Billing mission = filteredMissions[iMission];
                    if (mission.date_Mission_archived == "0000-00-00")
                    {
                        mission.date_Mission_archived = null;
                    }
                    //Show archive or not
                    if (m_Mission_IsArchiveMode == true && mission.date_Mission_archived != null)
                    {
                        Add_MissionToGrid(mission);
                    }
                    if (m_Mission_IsArchiveMode == false && mission.date_Mission_archived == null)
                    {
                        Add_MissionToGrid(mission);
                    }
                    if (missionSel != null)
                    {
                        if (missionSel.id == mission.id)
                        {
                            missionFound = mission;
                        }
                    }
                }

                //Select mission if necessary
                Select_Mission(missionFound);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Get an mission from its button
        /// <returns>Returns the mission corresponding to the button</returns>
        /// </summary>
        private Billing Get_SelectedMissionFromButton(bool _ShowMessage = true)
        {
            //Tests
            if (m_Button_Mission_SelectedMission == null)
            {
                if (_ShowMessage == true)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("missionSelectedError"), m_Global_Handler.Resources_Handler.Get_Resources("missionSelectedErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
                return null;
            }
            StackPanel stack = (StackPanel)m_Button_Mission_SelectedMission.Parent;
            if (stack.Tag == null)
            {
                MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("missionSelectedError"), m_Global_Handler.Resources_Handler.Get_Resources("missionSelectedErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            //Get the mission from the id
            int id = (int)stack.Tag;
            Billing mission = SoftwareObjects.MissionsCollection.Find(x => x.id.Equals(id));

            return mission;
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Manage the mission button, create and add or modify
        /// </summary>
        private void Manage_MissionButton(Billing _Mission, Button _MissionButton, Button _MissionsendMailButton, bool _Is_NewButton, bool _Is_Edition)
        {
            try
            {
                //Treatment of the creation date
                string creationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_Mission_creation, m_Global_Handler.Language_Handler);

                //Treatment of the modification date
                if (_Is_NewButton == false && _Is_Edition == true)
                {
                    string modificationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(DateTime.Today.ToString(), m_Global_Handler.Language_Handler);
                    string format = "yyyy-MM-dd";
                    _Mission.date_Mission_modification = DateTime.Now.ToString(format);
                }

                //Treatment of email sent date
                string sendDate = "";
                if (_Mission.date_Mission_sent != null)
                {
                    sendDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_Mission_sent, m_Global_Handler.Language_Handler);
                }

                //Treatment of accepted/declined date
                string date = "";
                string acceptedDate = "";
                string declinedDate = "";
                string labelDate = "";
                if (_Mission.date_Mission_accepted != null)
                {
                    acceptedDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_Mission_accepted, m_Global_Handler.Language_Handler);
                    date = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_Mission_accepted, m_Global_Handler.Language_Handler);
                    labelDate = m_Global_Handler.Resources_Handler.Get_Resources("AcceptedDate");
                }
                else if (_Mission.date_Mission_declined != null)
                {
                    declinedDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_Mission_declined, m_Global_Handler.Language_Handler);
                    date = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_Mission_declined, m_Global_Handler.Language_Handler);
                    labelDate = m_Global_Handler.Resources_Handler.Get_Resources("DeclinedDate");
                }

                //Treatment of archived date
                string archivedDate = "";
                if (_Mission.date_Mission_archived != null)
                {
                    archivedDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_Mission_archived, m_Global_Handler.Language_Handler);
                }

                //Get hostess infos
                Hostess hostess = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(_Mission.id_HostAndHostess));
                if (hostess == null)
                {
                    Exception error = new Exception("No hostess associated to the mission");

                    m_Database_Handler.Delete_MissionToDatabase(_Mission.id);
                    throw error;
                }

                //Treatment of mission button
                TextBlock textInfo = new TextBlock();
                Run names = new Run("\n" + hostess.firstname + " " + hostess.lastname);
                names.FontSize = 15;
                names.FontWeight = FontWeights.Bold;
                textInfo.Inlines.Add(names);
                //textInfo.Inlines.Add(new LineBreak());
                //textInfo.Inlines.Add(company);
                textInfo.Inlines.Add(new LineBreak());
                Run subject = new Run(_Mission.subject + "\n" + _Mission.num_Mission);
                subject.FontSize = 13.5;
                subject.FontWeight = FontWeights.Bold;
                textInfo.Inlines.Add(subject);
                textInfo.Inlines.Add(new LineBreak());
                string discount = "";
                if (_Mission.discount != 0)
                {
                    discount = m_Global_Handler.Resources_Handler.Get_Resources("GlobalDiscount") + " : " + _Mission.discount.ToString() + "%";
                }
                string info = discount + "\n"
                    + m_Global_Handler.Resources_Handler.Get_Resources("AmountET") + " : " + _Mission.amount.ToString("0.00", CultureInfo.GetCultureInfo(m_Global_Handler.Language_Handler)) + " " + m_Global_Handler.Resources_Handler.Get_Resources("Currency") + "\n"
                    + m_Global_Handler.Resources_Handler.Get_Resources("AmountIT") + " : " + _Mission.grand_amount.ToString("0.00", CultureInfo.GetCultureInfo(m_Global_Handler.Language_Handler)) + " " + m_Global_Handler.Resources_Handler.Get_Resources("Currency") + "\n"
                    + m_Global_Handler.Resources_Handler.Get_Resources("CreationDate") + " : " + creationDate + "\n";
                if (sendDate != "")
                {
                    info = info + m_Global_Handler.Resources_Handler.Get_Resources("MissionDate") + " : " + sendDate + "\n";
                }
                if (acceptedDate != "")
                {
                    info = info + m_Global_Handler.Resources_Handler.Get_Resources("AcceptedDate") + " : " + acceptedDate + "\n";
                }
                if (declinedDate != "")
                {
                    info = info + m_Global_Handler.Resources_Handler.Get_Resources("DeclinedDate") + " : " + declinedDate + "\n";
                }
                if (m_Mission_IsArchiveMode == true)
                {
                    info += m_Global_Handler.Resources_Handler.Get_Resources("ArchiveDate") + " : " + archivedDate + "\n";
                }
                textInfo.Inlines.Add(info);

                //Create stack button
                StackPanel buttonStack = new StackPanel();
                buttonStack.Orientation = Orientation.Horizontal;

                //Create image
                Image imgMission = new Image();
                if (_Mission.num_invoice != null && _Mission.num_invoice != "")
                {
                    //Related invoice created
                    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Billed);
                }
                else if (_Mission.date_Mission_accepted != null && _Mission.date_Mission_accepted != "")
                {
                    //Mission accepted
                    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Accepted);
                }
                else if (_Mission.date_Mission_declined != null && _Mission.date_Mission_declined != "")
                {
                    //Mission declined
                    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Declined);
                }
                else if (_Mission.date_Mission_sent != null && _Mission.date_Mission_sent != "")
                {
                    //Mission sent by email icon
                    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Sent);
                }
                else if (_Mission.date_Mission_generated != null && _Mission.date_Mission_generated != "")
                {
                    //Mission generated
                    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Generated);
                }
                else
                {
                    //Default icon
                    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Created);
                }

                //Effect
                System.Windows.Media.Effects.DropShadowEffect dropShadowEffect = new System.Windows.Media.Effects.DropShadowEffect();
                Color shadowColor = new Color();
                shadowColor.ScA = 1;
                shadowColor.ScB = 0;
                shadowColor.ScG = 0;
                shadowColor.ScR = 0;
                dropShadowEffect.Color = shadowColor;
                dropShadowEffect.Direction = 320;
                dropShadowEffect.ShadowDepth = 10;
                dropShadowEffect.BlurRadius = 10;
                dropShadowEffect.Opacity = 0.5;
                imgMission.Effect = dropShadowEffect;

                //Add image
                buttonStack.Children.Add(imgMission);

                //Add void
                Label voidLabel = new Label();
                voidLabel.Content = "\t";
                buttonStack.Children.Add(voidLabel);

                //Add text
                Label buttonLabel = new Label();
                buttonLabel.Content = textInfo;
                buttonStack.Children.Add(buttonLabel);

                //Fill button
                _MissionButton.Content = buttonStack;

                //Treatment of the email button
                Image imgEmail = new Image();
                imgEmail.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Image_Email);
                StackPanel stackEmail = new StackPanel();
                stackEmail.Orientation = Orientation.Horizontal;
                stackEmail.Margin = new Thickness(5);
                TextBlock emailInfo = new TextBlock();
                string emailText = "   " + hostess.email;
                emailInfo.Inlines.Add(emailText);
                stackEmail.Children.Add(imgEmail);
                stackEmail.Children.Add(emailInfo);
                _MissionsendMailButton.Content = stackEmail;

                if (_Is_NewButton == true)
                {
                    //Create the stack panel
                    StackPanel stackPanel = new StackPanel();
                    stackPanel.MinHeight = 45;
                    stackPanel.MinWidth = 145;
                    stackPanel.Margin = new Thickness(5);
                    stackPanel.Tag = _Mission.id;

                    //Properties of the mission button
                    _MissionButton.Padding = new Thickness(5);
                    _MissionButton.Background = m_Color_Mission;
                    if (_Mission.date_Mission_archived == "" || _Mission.date_Mission_archived == null)
                    {
                        //In progress mission
                        _MissionButton.Background = m_Color_Mission;
                    }
                    else
                    {
                        //Archived mission
                        _MissionButton.Background = m_Color_ArchivedMission;
                    }
                    _MissionButton.Click += IsPressed_MissionButton;
                    stackPanel.Children.Add(_MissionButton);

                    //Properties of the email button                    
                    _MissionsendMailButton.MinHeight = 30;
                    _MissionsendMailButton.Padding = new Thickness(5);
                    if (_Mission.date_Mission_archived == "" || _Mission.date_Mission_archived == null)
                    {
                        //In progress mission
                        _MissionsendMailButton.Background = m_Color_Mission;
                    }
                    else
                    {
                        //Archived mission
                        _MissionsendMailButton.Background = m_Color_ArchivedMission;
                    }
                    _MissionsendMailButton.Click += IsPressed_MissionEmailButton;
                    _MissionsendMailButton.Tag = "Email";
                    stackPanel.Children.Add(_MissionsendMailButton);

                    //Treatment of the increments
                    if (m_GridMissions_Column > Grid_Missions_Details.ColumnDefinitions.Count - 1)
                    {
                        m_GridMissions_Column = 0;
                        m_GridMissions_Row += 1;
                        RowDefinition row = new RowDefinition();
                        row.MinHeight = 350;
                        Grid_Missions_Details.RowDefinitions.Add(row);
                        Grid.SetColumn(stackPanel, m_GridMissions_Column);
                        Grid.SetRow(stackPanel, m_GridMissions_Row);
                        _MissionButton.Tag = m_GridMissions_Row.ToString() + "|" + m_GridMissions_Column.ToString();
                        m_GridMissions_Column += 1;
                    }
                    else
                    {
                        Grid.SetColumn(stackPanel, m_GridMissions_Column);
                        Grid.SetRow(stackPanel, m_GridMissions_Row);
                        _MissionButton.Tag = m_GridMissions_Row.ToString() + "|" + m_GridMissions_Column.ToString();
                        m_GridMissions_Column += 1;
                    }

                    //Treatment of the Grid_Mission column
                    Grid_Missions_Details.Children.Add(stackPanel);
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Select an mission
        /// </summary>
        private void Select_Mission(Billing _Mission)
        {
            try
            {
                //Verifications
                if (_Mission == null)
                {
                    //Treatment of the creation date
                    Txt_Missions_CreationDate.Text = "";

                    //Hostess text box
                    Txt_Missions_ClientCompanyName.Text = "";
                    Txt_Missions_FirstName.Text = "";
                    Txt_Missions_LastName.Text = "";

                    //Editing service fields and datagrid
                    m_DataGrid_Missions_ServicesCollection.Clear();
                    DataGrid_Missions_Services.Items.Refresh();

                    //Disable the buttons
                    Btn_Missions_Duplicate.IsEnabled = false;
                    Btn_Missions_Edit.IsEnabled = false;
                    Btn_Missions_Delete.IsEnabled = false;
                    Btn_Missions_Close.IsEnabled = false;

                    Thread.Sleep(100);
                    return;
                }

                //Manage the button
                Brush colorMission = m_Color_Mission;
                Brush colorSelectedMission = m_Color_SelectedMission;
                if (m_Mission_IsArchiveMode == true)
                {
                    colorMission = m_Color_ArchivedMission;
                    colorSelectedMission = m_Color_SelectedArchivedMission;
                }
                if (m_Button_Mission_SelectedMission != null)
                {
                    m_Button_Mission_SelectedMission.Background = colorMission;
                }
                if (m_Button_Mission_SelectedEmail != null)
                {
                    m_Button_Mission_SelectedEmail.Background = colorMission;
                }
                if (m_Button_Mission_SelectedInvoice != null)
                {
                    m_Button_Mission_SelectedInvoice.Background = colorMission;
                }
                //Get the mission stack panel
                double row = 0;
                for (int iMission = 0; iMission < Grid_Missions_Details.Children.Count; ++iMission)
                {
                    StackPanel stackSel = (StackPanel)Grid_Missions_Details.Children[iMission];
                    if ((int)stackSel.Tag == _Mission.id)
                    {
                        m_Button_Mission_SelectedMission = (Button)stackSel.Children[0];
                        m_Button_Mission_SelectedMission.Background = colorSelectedMission;
                        m_Button_Mission_SelectedEmail = (Button)stackSel.Children[1];
                        m_Button_Mission_SelectedEmail.Background = colorSelectedMission;
                        if (stackSel.Children.Count > 2)
                        {
                            m_Button_Mission_SelectedInvoice = (Button)stackSel.Children[2];
                            m_Button_Mission_SelectedInvoice.Background = colorSelectedMission;
                        }
                        stackSel.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                        stackSel.Arrange(new Rect(0, 0, stackSel.DesiredSize.Width, stackSel.DesiredSize.Height));
                        row = Grid.GetRow(stackSel) * stackSel.ActualHeight;
                        break;
                    }
                }

                //Treatment of the creation date
                string creationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_Mission_creation, m_Global_Handler.Language_Handler);
                if (m_Is_InitializationDone == true)
                {
                    Txt_Missions_CreationDate.Text = creationDate;
                }

                //Get hostess infos
                Hostess hostess = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(_Mission.id_HostAndHostess));
                if (hostess == null)
                {
                    Exception error = new Exception("No hostess associated to the mission");

                    m_Database_Handler.Delete_MissionToDatabase(_Mission.id);
                    throw error;
                }

                //Hostess text box
                //Txt_Missions_ClientCompanyName.Text = hostess.company;
                Txt_Missions_FirstName.Text = hostess.firstname;
                Txt_Missions_LastName.Text = hostess.lastname;

                //Scroll to the selected invoice
                Scrl_Grid_Missions_Details.ScrollToVerticalOffset(row);

                //Enable the buttons
                if (_Mission.date_Mission_archived == null || _Mission.date_Mission_archived == "")
                {
                    Btn_Missions_Duplicate.IsEnabled = true;
                    Btn_Missions_Edit.IsEnabled = true;
                    Btn_Missions_Delete.IsEnabled = true;
                    Btn_Missions_Close.IsEnabled = true;
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Sort the grid of missions by status
        /// </summary>
        private void Sort_MissionsByStatus(List<Billing> _SavedCollection)
        {
            try
            {
                //Get status
                List<Billing> createdMissions = new List<Billing>();
                List<Billing> generatedMissions = new List<Billing>();
                List<Billing> sendMissions = new List<Billing>();
                List<Billing> acceptedMissions = new List<Billing>();
                List<Billing> declinedMissions = new List<Billing>();
                List<Billing> billedMissions = new List<Billing>();
                for (int iMission = 0; iMission < m_Grid_Details_Missions_MissionsCollection.Count; ++iMission)
                {
                    Billing mission = m_Grid_Details_Missions_MissionsCollection[iMission];
                    if (mission.date_invoice_creation != "" && mission.date_invoice_creation != null)
                    {
                        billedMissions.Add(mission);
                    }
                    else if (mission.date_Mission_accepted != "" && mission.date_Mission_accepted != null)
                    {
                        acceptedMissions.Add(mission);
                    }
                    else if (mission.date_Mission_declined != "" && mission.date_Mission_declined != null)
                    {
                        declinedMissions.Add(mission);
                    }
                    else if (mission.date_Mission_sent != "" && mission.date_Mission_sent != null)
                    {
                        sendMissions.Add(mission);
                    }
                    else if (mission.date_Mission_generated != "" && mission.date_Mission_generated != null)
                    {
                        generatedMissions.Add(mission);
                    }
                    else
                    {
                        createdMissions.Add(mission);
                    }
                }
                //Sort each category by num mission
                List<List<Billing>> listCollections = new List<List<Billing>>();
                listCollections.Add(createdMissions);
                listCollections.Add(generatedMissions);
                listCollections.Add(sendMissions);
                listCollections.Add(declinedMissions);
                listCollections.Add(acceptedMissions);
                listCollections.Add(billedMissions);
                for (int iList = 0; iList < listCollections.Count; ++iList)
                {
                    List<Billing> list = listCollections[iList];
                    list.Sort(delegate (Billing x, Billing y)
                    {
                        return x.num_Mission.CompareTo(y.num_Mission);
                    });
                }

                //Add to grid layout collection
                m_Grid_Details_Missions_MissionsCollection.Clear();
                m_Grid_Details_Missions_MissionsCollection.AddRange(listCollections[0]);
                m_Grid_Details_Missions_MissionsCollection.AddRange(listCollections[1]);
                m_Grid_Details_Missions_MissionsCollection.AddRange(listCollections[2]);
                m_Grid_Details_Missions_MissionsCollection.AddRange(listCollections[3]);
                m_Grid_Details_Missions_MissionsCollection.AddRange(listCollections[4]);
                m_Grid_Details_Missions_MissionsCollection.AddRange(listCollections[5]);

                //Clear the grid
                Grid_Missions_Details.Children.Clear();
                Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                //Actualize the collection
                Actualize_GridMissionsFromMissionsCollection();

                //Select Status in the combo box
                Cmb_Missions_SortBy.SelectedItem = m_Global_Handler.Resources_Handler.Get_Resources("Status");
            }
            catch (Exception exception)
            {
                m_Global_Handler.Error_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                m_Grid_Details_Missions_MissionsCollection = _SavedCollection;

                //Clear the grid
                Grid_Missions_Details.Children.Clear();
                Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                //Actualize the collection
                Actualize_GridMissionsFromMissionsCollection();

                return;
            }
        }

        #endregion

        #region Settings

        #region

        /// <summary>
        /// Functions
        /// Settings
        /// Actualize the settings from the database
        /// </summary>
        private void Actualize_SettingsFromDatabase()
        {
            //Load settings from database
            string settings = m_Database_Handler.Get_SettingsFromDatabase();

            //Verification
            if (settings.Contains("OK"))
            {
                //Fill text boxes
                Txt_Settings_General_Database.Text = SoftwareObjects.GlobalSettings.database_definition;
                Txt_Settings_General_Photos.Text = SoftwareObjects.GlobalSettings.photos_repository;

                return;
            }
            else
            {
                MessageBox.Show(settings, m_Global_Handler.Resources_Handler.Get_Resources("ActualizeSettingsErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        #endregion

        #endregion

        #endregion

    }
}