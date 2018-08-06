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
        /// Collection containing the missions displayed
        /// </summary>
        private List<Mission> m_DisplayedMissionsCollection = new List<Mission>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Class describing a shift contained in the missions category datagrid
        /// </summary>
        public class m_Datagrid_Mission_Shifts
        {
            /// <summary>
            /// Constructor
            /// </summary>
            public m_Datagrid_Mission_Shifts(string _Id, string _Date, string _HostOrHostess, string _StartTime, string _EndTime)
            {
                id = _Id;
                date = _Date;
                hostorhostess = _HostOrHostess;
                start_time = _StartTime;
                end_time = _EndTime;
            }

            /// <summary>
            /// Id
            /// </summary>
            public string id { set; get; }
            /// <summary>
            /// Date
            /// </summary>
            public string date { set; get; }
            /// <summary>
            /// Date
            /// </summary>
            public string hostorhostess { set; get; }
            /// <summary>
            /// Start time
            /// </summary>
            public string start_time { set; get; }
            /// <summary>
            /// Start time
            /// </summary>
            public string end_time { set; get; }
        }

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection of services contained in the missions category datagrid
        /// </summary>
        ObservableCollection<m_Datagrid_Mission_Shifts> m_Datagrid_Missions_ShiftsCollection =
            new ObservableCollection<m_Datagrid_Mission_Shifts>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection containing the hosts and hostesses displayed
        /// </summary>
        private List<Hostess> m_DisplayedHostsAndHostessesCollection = new List<Hostess>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Class describing a mission contained in the host and hostess category datagrid
        /// </summary>
        public class m_Datagrid_HostAndHostess_Missions
        {
            /// <summary>
            /// Constructor
            /// </summary>
            public m_Datagrid_HostAndHostess_Missions(string _Id, string _Client, string _City, string _Date)
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
        ObservableCollection<m_Datagrid_HostAndHostess_Missions> m_DataGrid_HostAndHostess_MissionsCollection =
            new ObservableCollection<m_Datagrid_HostAndHostess_Missions>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection containing the clients displayed
        /// </summary>
        private List<Client> m_DisplayedClientsCollection = new List<Client>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Class describing a mission contained in the clients category datagrid
        /// </summary>
        public class m_Datagrid_Clients_Missions
        {
            /// <summary>
            /// Constructor
            /// </summary>
            public m_Datagrid_Clients_Missions(string _Id, string _City, string _Date, string _Description)
            {
                id = _Id;
                city = _City;
                date = _Date;
                description = _Description;
            }

            /// <summary>
            /// Id
            /// </summary>
            public string id { set; get; }
            /// <summary>
            /// City
            /// </summary>
            public string city { set; get; }
            /// <summary>
            /// Date
            /// </summary>
            public string date { set; get; }
            /// <summary>
            /// Description
            /// </summary>
            public string description { set; get; }
        }

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection of missions done by the selected host or hostess
        /// </summary>
        ObservableCollection<m_Datagrid_Clients_Missions> m_DataGrid_Clients_MissionsCollection =
            new ObservableCollection<m_Datagrid_Clients_Missions>();

        /// <summary>
        /// Initialization
        /// Datagrid
        /// Collection of missions panels contained in the missions grid
        /// </summary>
        List<Mission> m_Grid_Details_Missions_MissionsCollection = new List<Mission>();

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
        /// Selected mission button
        /// </summary>
        public Button m_Button_Mission_SelectedMission = null;

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected client cellphone button
        /// </summary>
        public Button m_Button_Client_SelectedCellPhone = null;

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected client button
        /// </summary>
        public Button m_Button_Client_SelectedClient = null;

        /// <summary>
        /// Initialization
        /// Selected buton
        /// Selected client email button
        /// </summary>
        public Button m_Button_Client_SelectedEmail = null;

        #endregion

        #region Specific variables

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Boolean indicating if the host or hostess category is in archive mode or not
        /// </summary>
        private bool m_HostsAndHostesses_IsArchiveMode = false;

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Boolean indicating if the clients category is in archive mode or not
        /// </summary>
        private bool m_Clients_IsArchiveMode = false;

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
        /// Selected client id
        /// </summary>
        public string m_Id_SelectedMission = "-1";

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Selected hostess id
        /// </summary>
        public string m_Id_SelectedHostAndHostess = "-1";

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Selected client id
        /// </summary>
        public string m_Id_SelectedClient = "-1";

        /// <summary>
        /// Initialization
        /// Specific variable
        /// Boolean indicating the sorting way of the clients
        /// </summary>
        bool m_IsSortClient_Ascending = true;

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
                m_Global_Handler.Log_Handler.WriteAction("New session");

                //Colors
                BrushConverter converter = new BrushConverter();
                Brush panelColor = (Brush)converter.ConvertFromString("#FFFFFFFF");
                m_Color_MainButton = (Brush)converter.ConvertFromString("#FF3B3839");
                m_Color_SelectedMainButton = (Brush)converter.ConvertFromString("#FF595C61");
                m_Color_Button = (Brush)converter.ConvertFromString("#FF6F0178");
                m_Color_Red = (Brush)converter.ConvertFromString("#FFF03535");
                m_Color_HostAndHostess = panelColor;
                m_Color_SelectedHostAndHostess = (Brush)converter.ConvertFromString("#FFF73D7B");
                m_Color_ArchivedMission = (Brush)converter.ConvertFromString("#FF595C61");
                m_Color_SelectedArchivedMission = (Brush)converter.ConvertFromString("#FF66A2CD");
                m_Color_Mission = panelColor;
                m_Color_SelectedMission = (Brush)converter.ConvertFromString("#FFF73D7B");

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

                //Actualize the clients collection from the database
                Actualize_GridClientsFromDatabase();

                //Actualize the missions collection
                Actualize_GridMissionsFromDatabase();

                //Fill the shifts collection
                Fill_ShiftsCollectionFromDatabase();

                //Setting the contents of objects
                if (m_Global_Handler != null)
                {
                    Define_Content();
                }
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
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
                TextBlock hostsandhostesses = new TextBlock();
                hostsandhostesses.Text = m_Global_Handler.Resources_Handler.Get_Resources("HostsAndHostesses");
                hostsandhostesses.TextAlignment = TextAlignment.Center;
                hostsandhostesses.TextWrapping = TextWrapping.Wrap;
                Btn_Software_HostAndHostess.Content = hostsandhostesses;
                Btn_Software_Missions.Content = m_Global_Handler.Resources_Handler.Get_Resources("Missions");
                Btn_Software_Settings.Content = m_Global_Handler.Resources_Handler.Get_Resources("Settings");


                #region Missions

                //Buttons
                Btn_Missions_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("ArchiveMission");
                Btn_Missions_Create.Content = m_Global_Handler.Resources_Handler.Get_Resources("CreateMission");
                Btn_Missions_Delete.Content = m_Global_Handler.Resources_Handler.Get_Resources("DeleteMission");
                Btn_Missions_Duplicate.Content = m_Global_Handler.Resources_Handler.Get_Resources("DuplicateMission");
                Btn_Missions_Edit.Content = m_Global_Handler.Resources_Handler.Get_Resources("EditMission");
                Btn_Missions_ShowArchived.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionsShowArchives");
                Btn_Missions_ShowInProgress.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionsShowInProgress");
                Btn_Missions_GenerateExcelStatement.Content = m_Global_Handler.Resources_Handler.Get_Resources("GenerateExcelStatement");

                //Combo boxes
                List<string> cmbMission = new List<string>();
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("Customer"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("City"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("CreationDate"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("ZipCode"));
                cmbMission.Add(m_Global_Handler.Resources_Handler.Get_Resources("Status"));
                cmbMission.Sort();
                foreach (string cmbMissionstr in cmbMission)
                {
                    Cmb_Missions_SortBy.Items.Add(cmbMissionstr);
                }

                //Datagrid
                Datagrid_Missions_Shifts.ItemsSource = m_Datagrid_Missions_ShiftsCollection;

                //Labels
                Lbl_Missions_CreationDate.Content = m_Global_Handler.Resources_Handler.Get_Resources("CreationDate");
                Lbl_Missions_Research.Content = m_Global_Handler.Resources_Handler.Get_Resources("ResearchMission");
                Lbl_Missions_Sort.Content = m_Global_Handler.Resources_Handler.Get_Resources("SortBy");
                Lbl_Missions_Image_Mission_Done.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionLegendDone");
                Lbl_Missions_Image_Mission_Billed.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionLegendBilled");
                Lbl_Missions_Image_Mission_Created.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionLegendCreated");
                Lbl_Missions_Image_Mission_Declined.Content = m_Global_Handler.Resources_Handler.Get_Resources("MissionLegendDeclined");

                #endregion

                #region Host and hostess

                //Buttons
                Btn_HostAndHostess_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("Archive");
                Btn_HostAndHostess_Create.Content = m_Global_Handler.Resources_Handler.Get_Resources("Create");
                Btn_HostAndHostess_Delete.Content = m_Global_Handler.Resources_Handler.Get_Resources("Delete");
                Btn_HostAndHostess_Edit.Content = m_Global_Handler.Resources_Handler.Get_Resources("Edit");
                Btn_HostAndHostess_GenerateExcelStatement.Content = m_Global_Handler.Resources_Handler.Get_Resources("GenerateExcelStatement");
                Btn_HostAndHostess_Import.Content = m_Global_Handler.Resources_Handler.Get_Resources("ImportHostAndHostessFromExcel");
                Btn_HostAndHostess_ShowArchived.Content = m_Global_Handler.Resources_Handler.Get_Resources("Archives");
                Btn_HostAndHostess_ShowInProgress.Content = m_Global_Handler.Resources_Handler.Get_Resources("InProgress");

                //Combobox
                List<string> cmbHostess = new List<string>();
                cmbHostess.Add(m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostessLastName"));
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
                Btn_Clients_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("Archive");
                Btn_Clients_Create.Content = m_Global_Handler.Resources_Handler.Get_Resources("Create");
                Btn_Clients_Delete.Content = m_Global_Handler.Resources_Handler.Get_Resources("Delete");
                Btn_Clients_Edit.Content = m_Global_Handler.Resources_Handler.Get_Resources("Edit");
                Btn_Clients_GenerateExcelStatement.Content = m_Global_Handler.Resources_Handler.Get_Resources("GenerateExcelStatement");
                Btn_Clients_ShowArchived.Content = m_Global_Handler.Resources_Handler.Get_Resources("Archives");
                Btn_Clients_ShowInProgress.Content = m_Global_Handler.Resources_Handler.Get_Resources("InProgress");

                //Combobox
                List<string> cmbClients = new List<string>();
                cmbClients.Add(m_Global_Handler.Resources_Handler.Get_Resources("CorporateName"));
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
                Lbl_Clients_Research.Content = m_Global_Handler.Resources_Handler.Get_Resources("ResearchClient");
                Lbl_Clients_Sort.Content = m_Global_Handler.Resources_Handler.Get_Resources("SortBy");

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
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
                m_Global_Handler.Log_Handler.WriteAction("End session");
                Closing -= Software_Window_Closing_Event;
                e.Cancel = true;
                var anim = new DoubleAnimation(0, (Duration)TimeSpan.FromSeconds(0.3));
                anim.Completed += (s, _) => this.Close();
                this.BeginAnimation(UIElement.OpacityProperty, anim);
                Environment.Exit(0);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                m_Global_Handler.Log_Handler.WriteAction("Contact customer service");

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

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
        private void Btn_Missions_Archive_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the mission
                Mission missionSel = Get_SelectedMissionFromButton();
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
                    string res = m_Database_Handler.Archive_Mission(missionSel.id);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Action
                        m_Global_Handler.Log_Handler.WriteAction("Mission " + missionSel.id + " archived");

                        //Actualize grid from collection
                        Actualize_GridMissionsFromDatabase();

                        //Clear the fields
                        Cmb_Missions_SortBy.Text = "";
                        Txt_Missions_Client.Text = "";
                        Txt_Missions_CreationDate.Text = "";
                        Txt_Missions_EndDate.Text = "";
                        Txt_Missions_Research.Text = "";
                        Txt_Missions_StartDate.Text = "";
                        m_Datagrid_Missions_ShiftsCollection.Clear();
                        Datagrid_Missions_Shifts.Items.Refresh();

                        //Null buttons
                        m_Button_Mission_SelectedMission = null;

                        //Disable the buttons
                        Btn_Missions_Archive.IsEnabled = false;
                        Btn_Missions_Delete.IsEnabled = false;
                        Btn_Missions_Duplicate.IsEnabled = false;
                        Btn_Missions_Edit.IsEnabled = false;

                        //Close the wait window
                        Thread.Sleep(500);
                        windowWait.Stop();

                        return;
                    }
                    else
                    {
                        //Close the wait window
                        Thread.Sleep(500);
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
                    string res = m_Database_Handler.Restore_Mission(missionSel.id);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Action
                        m_Global_Handler.Log_Handler.WriteAction("Mission " + missionSel.id + " restored");

                        //Actualize grid from collection
                        Actualize_GridMissionsFromDatabase();
                        m_Mission_SelectedStatus = MissionStatus.NONE;

                        //Clear the fields
                        Cmb_Missions_SortBy.Text = "";
                        Txt_Missions_Client.Text = "";
                        Txt_Missions_CreationDate.Text = "";
                        Txt_Missions_EndDate.Text = "";
                        Txt_Missions_Research.Text = "";
                        Txt_Missions_StartDate.Text = "";
                        m_Datagrid_Missions_ShiftsCollection.Clear();
                        Datagrid_Missions_Shifts.Items.Refresh();

                        //Null buttons
                        m_Button_Mission_SelectedMission = null;

                        //Disable the buttons
                        Btn_Missions_Archive.IsEnabled = false;
                        Btn_Missions_Delete.IsEnabled = false;
                        Btn_Missions_Duplicate.IsEnabled = false;
                        Btn_Missions_Edit.IsEnabled = false;

                        //Close the wait window
                        windowWait.Stop();

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                WindowMission.MainWindow_Mission missionWindow = new WindowMission.MainWindow_Mission(m_Global_Handler, m_Database_Handler, m_Cities_DocFrance, false, null);
                Nullable<bool> resShow = missionWindow.ShowDialog();

                //Validation
                if (resShow == true)
                {
                    if (missionWindow.m_Mission != null)
                    {
                        //Open the wait window
                        windowWait.Start(m_Global_Handler, "MissionCreationPrincipalMessage", "MissionCreationSecondaryMessage");

                        //Refresh datagrid
                        m_Datagrid_Missions_ShiftsCollection.Clear();

                        //Add mission to grid
                        Mission missionToAdd = missionWindow.m_Mission;
                        Add_MissionToGrid(missionToAdd);

                        //Action
                        m_Global_Handler.Log_Handler.WriteAction("Mission " + missionToAdd.client_name + " - From " + missionToAdd.start_date + " to " + missionToAdd.end_date + " created");

                        //Select the last hostess
                        if (m_Button_Mission_SelectedMission != null)
                        {
                            m_Button_Mission_SelectedMission.Background = m_Color_Mission;
                        }
                        StackPanel stack = Grid_Missions_Details.Children.Cast<StackPanel>().First(f => Grid.GetRow(f) == m_GridHostess_Row && Grid.GetColumn(f) == m_GridHostess_Column - 1);
                        for (int iChild = 0; iChild < stack.Children.Count; ++iChild)
                        {
                            Button childButton = (Button)stack.Children[iChild];
                            childButton.Background = m_Color_SelectedMission;
                            if (childButton.Tag.ToString() != "" && childButton.Tag.ToString() != "CellPhone" && childButton.Tag.ToString() != "Phone" && childButton.Tag.ToString() != "Email")
                            {
                                m_Button_Mission_SelectedMission = childButton;
                                m_Id_SelectedMission = (string)childButton.Tag;
                            }

                            //Display the last mission
                            Scrl_Grid_Missions_Details.ScrollToBottom();
                        }

                        //Select the mission
                        Select_Mission(missionToAdd);

                        //Enable the buttons
                        Btn_Missions_Delete.IsEnabled = true;
                        Btn_Missions_Duplicate.IsEnabled = true;
                        Btn_Missions_Edit.IsEnabled = true;

                        //Close the wait window
                        Thread.Sleep(500);
                        windowWait.Stop();

                        return;
                    }
                }
            }
            catch (Exception exception)
            {
                windowWait.Stop();
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                Mission missionSel = Get_SelectedMissionFromButton();
                if (missionSel == null)
                {
                    return;
                }

                //Confirm the delete
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissionConfirmDelete"),
                    m_Global_Handler.Resources_Handler.Get_Resources("MissionConfirmDeleteCaption"),
                    MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
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
                    m_Global_Handler.Log_Handler.WriteAction("Mission " + missionSel.client_name + " - From " + missionSel.start_date + " to " + missionSel.end_date + " deleted");

                    //Delete the associated shifts
                    List<String> listShiftsId = new List<string>();
                    if (missionSel.id_list_shifts != null)
                    {
                        listShiftsId = new List<string>(missionSel.id_list_shifts.Split(';'));
                    }
                    for (int iShift = 0; iShift < listShiftsId.Count; ++iShift)
                    {
                        m_Database_Handler.Delete_ShiftFromDatabase(listShiftsId[iShift]);
                    }

                    //Actualize and filter
                    Actualize_GridMissionsFromDatabase();

                    //Clear the fields
                    m_Id_SelectedMission = "-1";
                    Cmb_Missions_SortBy.Text = "";
                    Txt_Missions_Client.Text = "";
                    Txt_Missions_CreationDate.Text = "";
                    Txt_Missions_EndDate.Text = "";
                    Txt_Missions_Research.Text = "";
                    Txt_Missions_StartDate.Text = "";
                    m_Datagrid_Missions_ShiftsCollection.Clear();
                    Datagrid_Missions_Shifts.Items.Refresh();

                    //Null buttons
                    m_Button_Mission_SelectedMission = null;

                    //Disable the buttons
                    Btn_Missions_Archive.IsEnabled = false;
                    Btn_Missions_Delete.IsEnabled = false;
                    Btn_Missions_Duplicate.IsEnabled = false;
                    Btn_Missions_Edit.IsEnabled = false;

                    //Close the wait window
                    Thread.Sleep(500);
                    windowWait.Stop();

                    return;
                }
                else if (res.Contains("error"))
                {
                    //Close the wait window
                    Thread.Sleep(500);
                    windowWait.Stop();

                    //Treatment of the error
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

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

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                Mission missionSel = Get_SelectedMissionFromButton();
                if (missionSel == null)
                {
                    return;
                }

                //Open the wait window
                windowWait.Start(m_Global_Handler, "MissionDuplicationPrincipalMessage", "MissionDuplicationSecondaryMessage");

                //Get the created mission
                Mission newMission = new Mission(missionSel);
                newMission.id = newMission.Create_MissionId();
                newMission.date_creation = DateTime.Now.ToString();

                //Add to database
                string res = m_Database_Handler.Add_MissionToDatabase(newMission.address, newMission.city, newMission.client_name, newMission.country,
                    newMission.description, newMission.end_date, newMission.id, newMission.start_date, newMission.state, newMission.zipcode);

                //Treat the result
                if (res.Contains("OK"))
                {
                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Mission " + missionSel.client_name + " - From " + missionSel.start_date + " to " + missionSel.end_date + " duplicated");

                    //Add to collection
                    SoftwareObjects.MissionsCollection.Add(newMission);

                    //Actualize and filter
                    Actualize_GridMissionsFromDatabase();

                    //Select the new mission
                    Select_Mission(newMission);

                    //Close the window wait
                    Thread.Sleep(500);
                    windowWait.Stop();

                    return;
                }
                else
                {
                    //Close the wait window
                    Thread.Sleep(500);
                    windowWait.Stop();

                    //Error connecting to web site
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }

            }
            catch (Exception exception)
            {
                Thread.Sleep(500);
                windowWait.Stop();
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                Mission missionSel = Get_SelectedMissionFromButton();
                if (missionSel == null)
                {
                    return;
                }

                //Open the mission window
                WindowMission.MainWindow_Mission missionWindow = new WindowMission.MainWindow_Mission(m_Global_Handler, m_Database_Handler, m_Cities_DocFrance, true, missionSel);
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

                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Mission " + missionSel.id + " edited");

                    //State of mission
                    MissionStatus missionstate = MissionStatus.NONE;

                    //Treatment of mission button
                    StackPanel panel = (StackPanel)m_Button_Mission_SelectedMission.Parent;
                    Button buttonMission = (Button)panel.Children[0];
                    Manage_MissionButton(missionSel, buttonMission, false);

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
                    Thread.Sleep(500);
                    windowWait.Stop();

                    return;
                }
            }
            catch (Exception exception)
            {
                //Close the window
                Thread.Sleep(500);
                windowWait.Stop();

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                List<Mission> missionsToExport = new List<Mission>();
                for (int iMission = 0; iMission < m_DisplayedMissionsCollection.Count; ++iMission)
                {
                    Mission mission = (Mission)Datagrid_Missions_Shifts.Items[iMission];
                    DateTime dateCreation = Convert.ToDateTime(mission.date_creation);
                    if (dateCreation >= firstDatePicker && dateCreation <= endDatePicker)
                    {
                        missionsToExport.Add(mission);
                    }
                }

                //Sort the collection
                missionsToExport.Sort(delegate (Mission x, Mission y)
                {
                    if (x.id == null && y.id == null) return 0;
                    else if (x.id == null) return 1;
                    else if (y.id == null) return -1;
                    else
                    {
                        return x.date_creation.CompareTo(y.date_creation);
                    }
                });

                //Write in a file with tab separated
                List<string> lines = new List<string>();
                string columnsHeader = m_Global_Handler.Resources_Handler.Get_Resources("Id") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CreationDate") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Customer") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("StartDate") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("EndDate") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Address") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("ZipCode") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("City") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Status");
                lines.Add(columnsHeader);
                for (int iMission = 0; iMission < missionsToExport.Count; ++iMission)
                {
                    Mission mission = missionsToExport[iMission];
                    string line = mission.id + "\t" + mission.date_creation + "\t" + mission.client_name + "\t" + mission.start_date + "\t" + mission.end_date + "\t" +
                        mission.address + "\t" + mission.zipcode + "\t" + mission.city + "\t"; //TODO status
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
                m_Global_Handler.Log_Handler.WriteAction("Missions exported to excel file " + fileNameXLS);

                //Close the wait window
                windowWait.Stop();

                //Open the file
                Process.Start(fileNameXLS);
            }
            catch (Exception exception)
            {
                //Close the wait window
                windowWait.Stop();

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Missions
        /// Click on show archived missions button
        /// </summary>
        private void Btn_Missions_ShowArchived_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window wait
                windowWait.Start(m_Global_Handler, "MissionsShowArchivedPrincipalMessage", "MissionsShowArchivedSecondaryMessage");

                //Mode
                m_Mission_IsArchiveMode = true;

                //Actualize
                Actualize_GridMissionsFromDatabase();
                m_Mission_SelectedStatus = MissionStatus.NONE;

                //Clear the fields
                m_Id_SelectedMission = "-1";
                Cmb_Missions_SortBy.Text = "";
                Txt_Missions_Client.Text = "";
                Txt_Missions_CreationDate.Text = "";
                Txt_Missions_EndDate.Text = "";
                Txt_Missions_Research.Text = "";
                Txt_Missions_StartDate.Text = "";
                m_Datagrid_Missions_ShiftsCollection.Clear();
                Datagrid_Missions_Shifts.Items.Refresh();

                //Manage buttons
                Btn_Missions_Duplicate.IsEnabled = false;
                Btn_Missions_Edit.IsEnabled = false;
                Btn_Missions_Delete.IsEnabled = false;
                Btn_Missions_Archive.IsEnabled = false;
                Btn_Missions_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("Restore");
                Btn_Missions_Legend_Mission_Created.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Billed.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Declined.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Done.Background = m_Color_Button;

                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                windowWait.Start(m_Global_Handler, "MissionsShowInProgressPrincipalMessage", "MissionsShowInProgressSecondaryMessage");

                //Mode
                m_Mission_IsArchiveMode = false;

                //Actualize
                Actualize_GridMissionsFromDatabase();
                m_Mission_SelectedStatus = MissionStatus.NONE;

                //Clear the fields
                m_Id_SelectedMission = "-1";
                Cmb_Missions_SortBy.Text = "";
                Txt_Missions_Client.Text = "";
                Txt_Missions_CreationDate.Text = "";
                Txt_Missions_EndDate.Text = "";
                Txt_Missions_Research.Text = "";
                Txt_Missions_StartDate.Text = "";
                m_Datagrid_Missions_ShiftsCollection.Clear();
                Datagrid_Missions_Shifts.Items.Refresh();

                //Manage buttons
                Btn_Missions_Archive.IsEnabled = false;
                Btn_Missions_Delete.IsEnabled = false;
                Btn_Missions_Duplicate.IsEnabled = false;
                Btn_Missions_Edit.IsEnabled = false;
                Btn_Missions_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("Archive");
                Btn_Missions_Legend_Mission_Billed.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Created.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Declined.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Done.Background = m_Color_Button;

                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
            List<Mission> savedCollection = new List<Mission>();
            for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
            {
                //Get mission
                Mission mission = SoftwareObjects.MissionsCollection[iMission];
                //Archived
                if (m_Mission_IsArchiveMode == true && mission.archived == 1)
                {
                    savedCollection.Add(mission);
                }
                //In progress
                else if (m_Mission_IsArchiveMode == false && mission.archived == 0)
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
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("Customer"))
            {
                //Sort by last name
                try
                {
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Mission x, Mission y)
                    {
                        if (x.client_name == null && y.client_name == null) return 0;
                        else if (x.client_name == null) return -1;
                        else if (y.client_name == null) return 1;
                        else
                        {
                            if (m_IsSortMission_Ascending == false)
                                return x.client_name.CompareTo(y.client_name);
                            else
                                return y.client_name.CompareTo(x.client_name);
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Mission x, Mission y)
                    {
                        if (x.city == null && y.city == null) return 0;
                        else if (x.city == null) return -1;
                        else if (y.city == null) return 1;
                        {
                            if (m_IsSortMission_Ascending == false)
                                return x.city.CompareTo(y.city);
                            else
                                return y.city.CompareTo(x.city);
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Mission x, Mission y)
                    {
                        if (x.date_creation == null && y.date_creation == null) return 0;
                        else if (x.date_creation == null) return 1;
                        else if (y.date_creation == null) return -1;
                        else
                        {
                            DateTime xDate = Convert.ToDateTime(x.date_creation);
                            DateTime yDate = Convert.ToDateTime(y.date_creation);
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    m_Grid_Details_Missions_MissionsCollection = savedCollection;

                    //Clear the grid
                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridMissionsFromMissionsCollection();

                    return;
                }
            }
            else if (Cmb_Missions_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("ZipCode"))
            {
                //Sort by subject
                try
                {
                    m_Grid_Details_Missions_MissionsCollection.Sort(delegate (Mission x, Mission y)
                    {
                        if (x.zipcode == null && y.zipcode == null) return 0;
                        else if (x.zipcode == null) return 1;
                        else if (y.zipcode == null) return -1;
                        else
                        {
                            if (m_IsSortMission_Ascending == false)
                                return x.zipcode.CompareTo(y.zipcode);
                            else
                                return y.zipcode.CompareTo(x.zipcode);
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
        private void Datagrid_Missions_Shifts_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            string columnHeader = e.Column.Header.ToString();
            if (columnHeader == "date")
            {
                e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Date");
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            if (columnHeader == "hostorhostess")
            {
                e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostess");
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                DataGridTextColumn dataColumn = (DataGridTextColumn)e.Column;
                dataColumn.ElementStyle = new Style(typeof(TextBlock));
                dataColumn.ElementStyle.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap));
            }
            else if (columnHeader == "start_time")
            {
                e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("StartTime");
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            else if (columnHeader == "end_time")
            {
                e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("EndTime");
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            else
            {
                e.Column.Visibility = Visibility.Collapsed;
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
                    string idSel = (string)stackSel.Tag;

                    //Get the mission
                    Mission missionSel = SoftwareObjects.MissionsCollection.Find(x => x.id.Equals(idSel));
                    if (missionSel == null)
                    {
                        Exception error = new Exception("Mission not found !");
                        throw error;
                    }

                    if (missionSel != null)
                    {
                        //Fill the fields
                        Txt_Missions_Client.Text = missionSel.client_name;
                        Txt_Missions_CreationDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.date_creation, m_Global_Handler.Language_Handler);
                        Txt_Missions_EndDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.end_date, m_Global_Handler.Language_Handler);
                        Txt_Missions_StartDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.start_date, m_Global_Handler.Language_Handler);

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
                        }

                        if (m_Mission_IsArchiveMode == false)
                        {
                            //Enable the buttons
                            Btn_Missions_Archive.IsEnabled = true;
                            Btn_Missions_Delete.IsEnabled = true;
                            Btn_Missions_Duplicate.IsEnabled = true;
                            Btn_Missions_Edit.IsEnabled = true;
                        }
                        else
                        {
                            //Disable the buttons but Restore
                            Btn_Missions_Archive.IsEnabled = true;
                            Btn_Missions_Delete.IsEnabled = false;
                            Btn_Missions_Duplicate.IsEnabled = false;
                            Btn_Missions_Edit.IsEnabled = false;
                        }
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
            m_Id_SelectedMission = "-1";
            Cmb_Missions_SortBy.Text = "";
            Txt_Missions_Client.Text = "";
            Txt_Missions_CreationDate.Text = "";
            Txt_Missions_EndDate.Text = "";
            Txt_Missions_Research.Text = "";
            Txt_Missions_StartDate.Text = "";
            m_Datagrid_Missions_ShiftsCollection.Clear();
            Datagrid_Missions_Shifts.Items.Refresh();

            //Null buttons
            m_Button_Mission_SelectedMission = null;

            //Disable the buttons
            Btn_Missions_Archive.IsEnabled = false;
            Btn_Missions_Delete.IsEnabled = false;
            Btn_Missions_Duplicate.IsEnabled = false;
            Btn_Missions_Edit.IsEnabled = false;

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
                    List<Mission> foundMissionsList = new List<Mission>();
                    for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
                    {
                        //Get the mission
                        Mission processedMission = SoftwareObjects.MissionsCollection[iMission];

                        //Research
                        if (processedMission.client_name.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                        else if (processedMission.city.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                        else if (processedMission.address.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                        else if (processedMission.zipcode.ToLower().Contains(researchedText))
                        {
                            foundMissionsList.Add(processedMission);
                            continue;
                        }
                    }

                    //Clear displayed collection
                    m_DisplayedMissionsCollection.Clear();

                    //Displaying the found missions list
                    Grid_Missions_Details.Children.Clear();
                    actualizationMissionsDone = false;
                    if (foundMissionsList.Count > 0)
                    {
                        //Initialize counter for columns and rows of Grid_Missions
                        m_GridMission_Column = 0;
                        m_GridMission_Row = 0;
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        #endregion

        #region Hosts and hostesses

        /// <summary>
        /// Event
        /// Hostess
        /// Click on archive hostess
        /// </summary>
        private void Btn_HostAndHostess_Archive_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the host or hostess
                Hostess hostorHostessSel = Get_SelectedHostOrHostessFromButton();
                if (hostorHostessSel == null)
                {
                    return;
                }

                //In progress mode
                if (m_HostsAndHostesses_IsArchiveMode == false)
                {
                    //Open the window wait
                    windowWait.Start(m_Global_Handler, "HostOrHostessArchivePrincipalMessage", "HostOrHostessArchiveSecondaryMessage");

                    //Add to database    
                    string res = m_Database_Handler.Archive_HostOrHostess(hostorHostessSel.id);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Action
                        m_Global_Handler.Log_Handler.WriteAction("Host/hostess " + hostorHostessSel.firstname + " " + hostorHostessSel.lastname + " archived");

                        //Modify host or hostess archive mode
                        hostorHostessSel.archived = 1;

                        //Actualize grid from collection
                        Actualize_GridHostsAndHostessesFromDatabase();

                        //Clear the fields
                        Txt_HostAndHostess_Research.Text = "";
                        m_DataGrid_HostAndHostess_MissionsCollection.Clear();
                        DataGrid_HostAndHostess_Missions.Items.Refresh();

                        //Null buttons
                        m_Button_HostAndHostess_SelectedCellPhone = null;
                        m_Button_HostAndHostess_SelectedEmail = null;
                        m_Button_HostAndHostess_SelectedHostAndHostess = null;

                        //Disable the buttons
                        Btn_HostAndHostess_Archive.IsEnabled = false;
                        Btn_HostAndHostess_Create.IsEnabled = false;
                        Btn_HostAndHostess_Edit.IsEnabled = false;
                        Btn_HostAndHostess_Delete.IsEnabled = false;

                        //Close the wait window
                        windowWait.Stop();

                        return;
                    }
                    else if (res.Contains("error"))
                    {
                        //Close the wait window
                        windowWait.Stop();

                        //Treatment of the error
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

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
                    windowWait.Start(m_Global_Handler, "HostOrHostessRestorePrincipalMessage", "HostOrHostessRestoreSecondaryMessage");

                    //Add to database    
                    string res = m_Database_Handler.Restore_HostOrHostess(hostorHostessSel.id);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Action
                        m_Global_Handler.Log_Handler.WriteAction("Host/hostess " + hostorHostessSel.firstname + " " + hostorHostessSel.lastname + " restored");

                        //Modify host or hostess archive mode
                        hostorHostessSel.archived = 0;

                        //Actualize grid from collection
                        Actualize_GridHostsAndHostessesFromDatabase();

                        //Clear the fields
                        Txt_HostAndHostess_Research.Text = "";
                        m_DataGrid_HostAndHostess_MissionsCollection.Clear();
                        DataGrid_HostAndHostess_Missions.Items.Refresh();

                        //Null buttons
                        m_Button_HostAndHostess_SelectedCellPhone = null;
                        m_Button_HostAndHostess_SelectedEmail = null;
                        m_Button_HostAndHostess_SelectedHostAndHostess = null;

                        //Disable the buttons
                        Btn_HostAndHostess_Archive.IsEnabled = false;
                        Btn_HostAndHostess_Create.IsEnabled = false;
                        Btn_HostAndHostess_Edit.IsEnabled = false;
                        Btn_HostAndHostess_Delete.IsEnabled = false;

                        //Close the wait window
                        windowWait.Stop();

                        return;
                    }
                    else if (res.Contains("error"))
                    {
                        //Close the wait window
                        windowWait.Stop();

                        //Treatment of the error
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                Close();
            }
        }

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
                    m_Global_Handler.Log_Handler.WriteAction("Host/Hostess " + hostOrHostessToAdd.firstname + " " + hostOrHostessToAdd.lastname + " created");

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                List<Mission> associatedMissions = new List<Mission>();
                for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
                {
                    Mission missionSel = SoftwareObjects.MissionsCollection[iMission];
                    if (missionSel.id == hostOrHostess.id) //TODO
                    {
                        associatedMissions.Add(missionSel);
                    }
                }
                if (associatedMissions.Count > 0)
                {
                    string message = m_Global_Handler.Resources_Handler.Get_Resources("HostessForbiddenDeleteText") + "\n";
                    message += m_Global_Handler.Resources_Handler.Get_Resources("Mission") + ": ";
                    for (int iMission = 0; iMission < associatedMissions.Count; ++iMission)
                    {
                        message += " " + associatedMissions[iMission].client_name;
                    }
                    MessageBox.Show(this, message, m_Global_Handler.Resources_Handler.Get_Resources("HostessForbiddenDeleteCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                //Confirm the delete
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostessConfirmDeleteText"),
                    m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostessConfirmDeleteCaption"), MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                //Edit in internet database
                string res = m_Database_Handler.Delete_HostAndHostessFromDatabase(hostOrHostess.id);

                //Treat the result
                if (res.Contains("OK"))
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "HostOrHostessDeletionPrincipalMessage", "HostOrHostessDeletionSecondaryMessage");

                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Host/Hostess " + hostOrHostess.firstname + " " + hostOrHostess.lastname + " deleted");

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    m_Global_Handler.Log_Handler.WriteAction("Host/Hostess " + hostOrHostess.firstname + " " + hostOrHostess.lastname + " edited");

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                string fileNameXLS = m_Global_Handler.Resources_Handler.Get_Resources("HostsAndHostessesStatement") + " - " + today;
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
                windowWait.Start(m_Global_Handler, "HostsAndHostessesExcelGenerationPrincipalMessage", "HostsAndHostessesExcelGenerationSecondaryMessage");

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
                string columnsHeader = m_Global_Handler.Resources_Handler.Get_Resources("FirstName") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("LastName") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CellPhone") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Sex") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Email") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Address") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("ZipCode") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("City") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Country") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("IdPaycheck") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("BirthDate") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("BirthCity") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("SocialSecurityNumber") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Size") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("SizeShirt") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("SizePants") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("SizeShoes") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Car") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("DriverLicence") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("English") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("German") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Spanish") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Italian") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Others") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Street") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Event") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Permanent");
                lines.Add(columnsHeader);
                for (int iHostess = 0; iHostess < m_DisplayedHostsAndHostessesCollection.Count; ++iHostess)
                {
                    Hostess hostess = m_DisplayedHostsAndHostessesCollection[iHostess];
                    string hasCar = (hostess.has_car == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string hasDriverLicence = (hostess.has_driver_licence == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string speakEnglish = (hostess.language_english == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string speakGerman = (hostess.language_german == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string speakSpanish = (hostess.language_spanish == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string speakItalian = (hostess.language_italian == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string isProfileStreet = (hostess.profile_street == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string isProfileEvent = (hostess.profile_event == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string isProfilePermanent = (hostess.profile_permanent == true) ? m_Global_Handler.Resources_Handler.Get_Resources("Yes") : m_Global_Handler.Resources_Handler.Get_Resources("No");
                    string line = hostess.firstname + "\t" + hostess.lastname + "\t" + hostess.cellphone + "\t" + hostess.sex + "\t" + hostess.email + "\t" + hostess.address + "\t" +
                        hostess.zipcode + "\t" + hostess.city + "\t" + hostess.country + "\t" + hostess.id_paycheck + "\t" + hostess.birth_date + "\t" + hostess.birth_city + "\t" +
                        hostess.social_number + "\t" + hostess.size + "\t" + hostess.size_shirt + "\t" + hostess.size_pants + "\t" + hostess.size_shoes + "\t" + hasCar + "\t" +
                        hasDriverLicence + "\t" + speakEnglish + "\t" + speakGerman + "\t" + speakSpanish + "\t" + speakItalian + "\t" + hostess.language_others + "\t" +
                        isProfileStreet + "\t" + isProfileEvent + "\t" + isProfilePermanent;
                    ;
                    lines.Add(line);
                }

                //Generate the excel file
                int result = m_Global_Handler.Excel_Handler.Generate_ExcelStatement(fileNameXLS, lines);
                if (result == -1)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("HostsAndHostessesStatementExcelGenerationFailedText"),
                                m_Global_Handler.Resources_Handler.Get_Resources("HostsAndHostessesStatementExcelGenerationFailedCaption"),
                                MessageBoxButton.OK, MessageBoxImage.Error);
                    //Close the wait window
                    windowWait.Stop();

                    return;
                }

                //Action
                m_Global_Handler.Log_Handler.WriteAction("Hosts and hostesses statement generated to Excel file " + fileNameXLS);

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Click on Import hosts and hostesses form an Excel file
        /// </summary>
        private void Btn_HostAndHostess_Import_Click(object sender, RoutedEventArgs e)
        {
            //Verify excel format
            MessageBoxResult resultTemplate = MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("HostAndHostessExcelImportTemplateText"),
                m_Global_Handler.Resources_Handler.Get_Resources("HostAndHostessExcelImportTemplateCaption"), MessageBoxButton.OKCancel, MessageBoxImage.Information);
            if (resultTemplate == MessageBoxResult.Cancel)
            {
                return;
            }

            //Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xls";
            dlg.Filter = "Excel Files (*.xls;*.xlsx)|*.xls; *.xlsx";

            //Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display it
            if (result == true)
            {
                //Creation of the wait window
                WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

                //Action
                m_Global_Handler.Log_Handler.WriteAction("Hosts and hostesses from " + dlg.FileName);


                try
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "HostAndHostessExcelImportPrincipalMessage", "HostAndHostessExcelImportSecondaryMessage");

                    // Open document 
                    string excelFilename = dlg.FileName;
                    string extension = Path.GetExtension(excelFilename);
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(excelFilename, 0, true, 5, "", "", true,
                        Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;

                    //Read and fill collection
                    int index = 2;
                    while (((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 1]).Value != null)
                    {
                        //Creation of service
                        Hostess hostess = new Hostess();
                        hostess.firstname = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 1]).Text);
                        hostess.lastname = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 2]).Text);
                        hostess.cellphone = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 3]).Value);
                        hostess.sex = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 4]).Value);
                        hostess.email = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 5]).Value);
                        hostess.address = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 6]).Text);
                        hostess.zipcode = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 7]).Text);
                        hostess.city = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 8]).Value);
                        hostess.country = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 9]).Value);
                        hostess.id_paycheck = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 10]).Value);
                        hostess.birth_date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 11]).Value);
                        hostess.birth_city = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 12]).Value);
                        hostess.social_number = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 13]).Value);
                        hostess.size = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 14]).Value);
                        hostess.size_shirt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 15]).Value);
                        hostess.size_pants = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 16]).Value);
                        hostess.size_shoes = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 17]).Value);
                        string hasCar = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 15]).Value);
                        if (hasCar.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.has_car = true;
                        }
                        else
                        {
                            hostess.has_car = false;
                        }
                        string hasDriverLicence = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 18]).Value);
                        if (hasDriverLicence.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.has_driver_licence = true;
                        }
                        else
                        {
                            hostess.has_driver_licence = false;
                        }
                        string english = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 19]).Value);
                        if (english.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.language_english = true;
                        }
                        else
                        {
                            hostess.language_english = false;
                        }
                        string german = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 20]).Value);
                        if (german.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.language_german = true;
                        }
                        else
                        {
                            hostess.language_german = false;
                        }
                        string spanish = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 21]).Value);
                        if (spanish.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.language_spanish = true;
                        }
                        else
                        {
                            hostess.language_spanish = false;
                        }
                        string italian = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 22]).Value);
                        if (italian.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.language_italian = true;
                        }
                        else
                        {
                            hostess.language_italian = false;
                        }
                        hostess.language_others = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 23]).Value);
                        string streetProfile = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 24]).Value);
                        if (streetProfile.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.profile_street = true;
                        }
                        else
                        {
                            hostess.profile_street = false;
                        }
                        string eventProfile = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 25]).Value);
                        if (eventProfile.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.profile_event = true;
                        }
                        else
                        {
                            hostess.profile_event = false;
                        }
                        string permanentProfile = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[index, 25]).Value);
                        if (permanentProfile.ToLower() == m_Global_Handler.Resources_Handler.Get_Resources("Yes").ToLower())
                        {
                            hostess.profile_permanent = true;
                        }
                        else
                        {
                            hostess.profile_permanent = false;
                        }

                        //Verification
                        bool is_HostessAlreadyAdded = false;
                        for (int iHostess = 0; iHostess < SoftwareObjects.HostsAndHotessesCollection.Count; ++iHostess)
                        {
                            Hostess hostessInCollection = SoftwareObjects.HostsAndHotessesCollection[iHostess];
                            if (hostessInCollection.social_number == hostess.social_number)
                            {
                                is_HostessAlreadyAdded = true;
                                break;
                            }
                        }
                        if (is_HostessAlreadyAdded == true)
                        {
                            //Incrementation
                            ++index;

                            //Continue loop
                            continue;
                        }

                        //Creation of the id
                        hostess.id = hostess.Create_HostOrHostessId();

                        //Creation date
                        hostess.date_creation = DateTime.Now.ToString();

                        //Add to internet database        
                        string res = m_Database_Handler.Add_HostAndHostessToDatabase(hostess.address, hostess.birth_city,
                            hostess.birth_date, hostess.cellphone, hostess.city, hostess.country, hostess.email,
                            hostess.firstname, hostess.has_car, hostess.has_driver_licence, hostess.id, hostess.id_paycheck,
                            hostess.language_english, hostess.language_german, hostess.language_italian, hostess.language_others,
                            hostess.language_spanish, hostess.lastname, hostess.profile_event, hostess.profile_permanent,
                            hostess.profile_street, hostess.sex, hostess.size, hostess.size_pants, hostess.size_shirt,
                            hostess.size_shoes, hostess.social_number, "", hostess.zipcode);

                        //Treat the result
                        if (res.Contains("OK"))
                        {
                            //Incrementation
                            ++index;
                        }
                        else if (res.Contains("Error"))
                        {
                            //Close the wait window
                            windowWait.Stop();

                            //Incrementation
                            ++index;

                            //Treatment of the error
                            MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

                            //Continue loop
                            continue;
                        }
                    }

                    //Actualization
                    Actualize_GridHostsAndHostessesFromDatabase();

                    //Close the wait window
                    windowWait.Stop();

                    //Message box and disable the edit and delete buttons
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("HostAndHostessExcelImportEndedText"),
                        m_Global_Handler.Resources_Handler.Get_Resources("HostAndHostessExcelImportEndedCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    //Close the excel file without saving
                    workBook.Close(false, Type.Missing, Type.Missing);

                    return;

                }
                catch (Exception exception)
                {
                    //Close the wait window
                    windowWait.Stop();

                    //Write the error into log
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Click on show archived hosts and hostesses
        /// </summary>
        private void Btn_HostAndHostess_ShowArchived_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window wait
                windowWait.Start(m_Global_Handler, "HostsAndHostessesShowArchivedPrincipalMessage", "HostsAndHostessesShowArchivedSecondaryMessage");

                //Mode
                m_HostsAndHostesses_IsArchiveMode = true;

                //Actualize
                Actualize_GridHostsAndHostessesFromDatabase();

                //Clear the fields
                Txt_HostAndHostess_Research.Text = "";
                m_DataGrid_HostAndHostess_MissionsCollection.Clear();
                DataGrid_HostAndHostess_Missions.Items.Refresh();

                //Null buttons
                m_Button_HostAndHostess_SelectedCellPhone = null;
                m_Button_HostAndHostess_SelectedEmail = null;
                m_Button_HostAndHostess_SelectedHostAndHostess = null;

                //Disable the buttons
                Btn_HostAndHostess_Archive.IsEnabled = false;
                Btn_HostAndHostess_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("Restore");
                Btn_HostAndHostess_Create.IsEnabled = false;
                Btn_HostAndHostess_Edit.IsEnabled = false;
                Btn_HostAndHostess_Delete.IsEnabled = false;

                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Hostess
        /// Click on show in progress hosts and hostesses
        /// </summary>
        private void Btn_HostAndHostess_ShowInProgress_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window wait
                windowWait.Start(m_Global_Handler, "HostsAndHostessesShowInProgressPrincipalMessage", "HostsAndHostessesShowInProgressSecondaryMessage");

                //Mode
                m_HostsAndHostesses_IsArchiveMode = false;

                //Actualize
                Actualize_GridHostsAndHostessesFromDatabase();

                //Clear the fields
                Txt_HostAndHostess_Research.Text = "";
                m_DataGrid_HostAndHostess_MissionsCollection.Clear();
                DataGrid_HostAndHostess_Missions.Items.Refresh();

                //Null buttons
                m_Button_HostAndHostess_SelectedCellPhone = null;
                m_Button_HostAndHostess_SelectedEmail = null;
                m_Button_HostAndHostess_SelectedHostAndHostess = null;

                //Disable the buttons
                Btn_HostAndHostess_Archive.IsEnabled = false;
                Btn_HostAndHostess_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("Archive");
                Btn_HostAndHostess_Create.IsEnabled = false;
                Btn_HostAndHostess_Edit.IsEnabled = false;
                Btn_HostAndHostess_Delete.IsEnabled = false;

                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
            else if (Cmb_HostAndHostess_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostessLastName"))
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                        m_Datagrid_HostAndHostess_Missions hostessGrid = (m_Datagrid_HostAndHostess_Missions)grid.SelectedItems[0];
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
                        m_Datagrid_HostAndHostess_Missions billDatagrid = (m_Datagrid_HostAndHostess_Missions)grid.SelectedCells[0].Item;

                        //Get the missions associated to the selected host or hostess
                        List<Mission> missions = new List<Mission>();
                        for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
                        {
                            Mission bill = SoftwareObjects.MissionsCollection[iMission];
                            if (bill.id == Hostessel.id)//TODO
                            {
                                missions.Add(bill);
                            }
                        }
                        m_Grid_Details_Missions_MissionsCollection = missions;

                        //Get selected bill from collection
                        Mission missionSel = SoftwareObjects.MissionsCollection.Find(x => x.id.Equals(billDatagrid.id));
                        if (missionSel != null)
                        {
                            //Mission
                            if (missionSel.id == "" || missionSel.id == null)
                            {
                                //Graphism
                                Btn_Software_Missions.Background = m_Color_MainButton;
                                Btn_Software_HostAndHostess.Background = m_Color_SelectedMainButton;

                                //Visibility
                                Grid_HostAndHostess.Visibility = Visibility.Collapsed;
                                Grid_Missions.Visibility = Visibility.Visible;

                                //Archived or in progress
                                if (missionSel.archived == 1)
                                {
                                    m_Mission_IsArchiveMode = true;
                                }

                                //Clear displayed collection
                                m_DisplayedMissionsCollection.Clear();

                                //Adding associated missions
                                Grid_Missions_Details.Children.Clear();
                                m_GridMission_Column = 0;
                                m_GridMission_Row = 0;
                                for (int iMission = 0; iMission < missions.Count; ++iMission)
                                {
                                    if (m_Mission_IsArchiveMode == true && missions[iMission].archived == 1)
                                    {
                                        Add_MissionToGrid(missions[iMission]);
                                    }
                                    else if (m_Mission_IsArchiveMode == false && missions[iMission].archived == 0)
                                    {
                                        Add_MissionToGrid(missions[iMission]);
                                    }
                                }

                                //Fields
                                Txt_Missions_Client.Text = missionSel.client_name;
                                Txt_Missions_CreationDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.date_creation, m_Global_Handler.Language_Handler);
                                Txt_Missions_EndDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.end_date, m_Global_Handler.Language_Handler);
                                Txt_Missions_StartDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(missionSel.start_date, m_Global_Handler.Language_Handler);

                                //Select the mission
                                Filter_GridMissionsFromMissionsCollection(MissionStatus.NONE);
                                Select_Mission(missionSel);

                                //Enable the buttons
                                Btn_Missions_Duplicate.IsEnabled = true;
                                Btn_Missions_Edit.IsEnabled = true;
                                Btn_Missions_Delete.IsEnabled = true;
                                Btn_Missions_Archive.IsEnabled = true;
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_ArchivedMission;
                        }
                    }
                    if (m_Button_HostAndHostess_SelectedCellPhone != null)
                    {
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_ArchivedMission;
                        }
                    }
                    if (m_Button_HostAndHostess_SelectedEmail != null)
                    {
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedEmail.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedEmail.Background = m_Color_ArchivedMission;
                        }
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
                    //Select the host or hostess
                    Get_SelectedHostOrHostessFromButton();

                    if (m_HostsAndHostesses_IsArchiveMode == false)
                    {
                        //Enable the buttons
                        Btn_HostAndHostess_Archive.IsEnabled = true;
                        Btn_HostAndHostess_Edit.IsEnabled = true;
                        Btn_HostAndHostess_Delete.IsEnabled = true;
                    }
                    else
                    {
                        //Enable the buttons
                        Btn_HostAndHostess_Archive.IsEnabled = true;
                        Btn_HostAndHostess_Edit.IsEnabled = false;
                        Btn_HostAndHostess_Delete.IsEnabled = false;
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    windowWait.Start(m_Global_Handler, "HostOrHostessEmailPrincipalMessage", "HostOrHostessEmailSecondaryMessage");

                    //Get the email adress
                    Button emailButton = (Button)sender;
                    StackPanel emailStackPanel = (StackPanel)emailButton.Content;
                    TextBlock emailTextBox = (TextBlock)emailStackPanel.Children[1];
                    string emailAddress = emailTextBox.Text;

                    //Manage the buttons
                    if (m_Button_HostAndHostess_SelectedHostAndHostess != null)
                    {
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_ArchivedMission;
                        }
                    }
                    if (m_Button_HostAndHostess_SelectedCellPhone != null)
                    {
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_ArchivedMission;
                        }
                    }
                    if (m_Button_HostAndHostess_SelectedEmail != null)
                    {
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedEmail.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedEmail.Background = m_Color_ArchivedMission;
                        }
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
                    Btn_HostAndHostess_Archive.IsEnabled = true;
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
                    m_Global_Handler.Log_Handler.WriteAction("Mail sent to " + emailAddress);

                    //Close stop window
                    windowWait.Stop();

                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    windowWait.Start(m_Global_Handler, "HostOrHostessPhonePrincipalMessage", "HostOrHostessPhoneSecondaryMessage");

                    //Get the email adress
                    Button phoneButton = (Button)sender;
                    StackPanel phoneStackPanel = (StackPanel)phoneButton.Content;
                    TextBlock phoneTextBox = (TextBlock)phoneStackPanel.Children[1];
                    string phoneNumber = phoneTextBox.Text.Trim();

                    //Manage the buttons
                    if (m_Button_HostAndHostess_SelectedHostAndHostess != null)
                    {
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedHostAndHostess.Background = m_Color_ArchivedMission;
                        }
                    }
                    if (m_Button_HostAndHostess_SelectedCellPhone != null)
                    {
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedCellPhone.Background = m_Color_ArchivedMission;
                        }
                    }
                    if (m_Button_HostAndHostess_SelectedEmail != null)
                    {
                        if (m_HostsAndHostesses_IsArchiveMode == false)
                        {
                            m_Button_HostAndHostess_SelectedEmail.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_HostAndHostess_SelectedEmail.Background = m_Color_ArchivedMission;
                        }
                    }
                    m_Button_HostAndHostess_SelectedHostAndHostess = null;
                    m_Button_HostAndHostess_SelectedCellPhone = null;
                    m_Button_HostAndHostess_SelectedEmail = null;
                    StackPanel stackSel = (StackPanel)phoneButton.Parent;
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
                    Btn_HostAndHostess_Archive.IsEnabled = true;
                    Btn_HostAndHostess_Edit.IsEnabled = true;
                    Btn_HostAndHostess_Delete.IsEnabled = true;

                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Skype call made to " + hostOrHostess.firstname + ", " + hostOrHostess.cellphone);

                    //Close stop window
                    windowWait.Stop();

                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                    m_DisplayedHostsAndHostessesCollection.Clear();
                    for (int iHostess = 0; iHostess < foundHostessList.Count; ++iHostess)
                    {
                        Add_HostOrHostessToGrid(foundHostessList[iHostess]);
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        #endregion

        #region Clients

        /// <summary>
        /// Event
        /// Clients
        /// Click on archive client
        /// </summary>
        private void Btn_Clients_Archive_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the client
                Client clientSel = Get_SelectedClientFromButton();
                if (clientSel == null)
                {
                    return;
                }

                //In progress mode
                if (m_Clients_IsArchiveMode == false)
                {
                    //Open the window wait
                    windowWait.Start(m_Global_Handler, "ClientArchivePrincipalMessage", "ClientArchiveSecondaryMessage");

                    //Add to database    
                    string res = m_Database_Handler.Archive_Client(clientSel.id);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Action
                        m_Global_Handler.Log_Handler.WriteAction("Client " + clientSel.corporate_name + " archived");

                        //Modify client archive mode
                        clientSel.archived = 1;

                        //Actualize grid from collection
                        Actualize_GridClientsFromDatabase();

                        //Clear the fields
                        Txt_Clients_Research.Text = "";
                        m_DataGrid_Clients_MissionsCollection.Clear();
                        DataGrid_Clients_Missions.Items.Refresh();

                        //Null buttons
                        m_Button_Client_SelectedCellPhone = null;
                        m_Button_Client_SelectedEmail = null;
                        m_Button_Client_SelectedClient = null;

                        //Disable the buttons
                        Btn_Clients_Archive.IsEnabled = false;
                        Btn_Clients_Create.IsEnabled = false;
                        Btn_Clients_Edit.IsEnabled = false;
                        Btn_Clients_Delete.IsEnabled = false;

                        //Close the wait window
                        windowWait.Stop();

                        return;
                    }
                    else if (res.Contains("error"))
                    {
                        //Close the wait window
                        windowWait.Stop();

                        //Treatment of the error
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

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
                    windowWait.Start(m_Global_Handler, "ClientRestorePrincipalMessage", "ClientRestoreSecondaryMessage");

                    //Add to database    
                    string res = m_Database_Handler.Restore_Client(clientSel.id);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Action
                        m_Global_Handler.Log_Handler.WriteAction("Client " + clientSel.corporate_name + " restored");

                        //Modify client archive mode
                        clientSel.archived = 0;

                        //Actualize grid from collection
                        Actualize_GridClientsFromDatabase();

                        //Clear the fields
                        Txt_Clients_Research.Text = "";
                        m_DataGrid_Clients_MissionsCollection.Clear();
                        DataGrid_Clients_Missions.Items.Refresh();

                        //Null buttons
                        m_Button_Client_SelectedCellPhone = null;
                        m_Button_Client_SelectedEmail = null;
                        m_Button_Client_SelectedClient = null;

                        //Disable the buttons
                        Btn_Clients_Archive.IsEnabled = false;
                        Btn_Clients_Create.IsEnabled = false;
                        Btn_Clients_Edit.IsEnabled = false;
                        Btn_Clients_Delete.IsEnabled = false;

                        //Close the wait window
                        windowWait.Stop();

                        return;
                    }
                    else if (res.Contains("error"))
                    {
                        //Close the wait window
                        windowWait.Stop();

                        //Treatment of the error
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                Close();
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on add client
        /// </summary>
        private void Btn_Clients_Create_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window
                WindowClient.MainWindow_Client client_CreateWindow = new WindowClient.MainWindow_Client(m_Global_Handler,
                    m_Database_Handler, m_Cities_DocFrance, false, null);
                Nullable<bool> resShow = client_CreateWindow.ShowDialog();

                //Validation
                if (resShow == true)
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "ClientCreationPrincipalMessage", "ClientCreationSecondaryMessage");

                    //Refresh datagrid
                    m_DataGrid_Clients_MissionsCollection.Clear();

                    //Add hostess to grid
                    Client clientToAdd = client_CreateWindow.m_Client;
                    Add_ClientToGrid(client_CreateWindow.m_Client);

                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Client " + clientToAdd.corporate_name + " created");

                    //Select the last hostess
                    if (m_Button_Client_SelectedClient != null)
                    {
                        m_Button_Client_SelectedClient.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_Client_SelectedCellPhone != null)
                    {
                        m_Button_Client_SelectedCellPhone.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_Client_SelectedEmail != null)
                    {
                        m_Button_Client_SelectedEmail.Background = m_Color_HostAndHostess;
                    }
                    StackPanel stack = Grid_Clients_Details.Children.Cast<StackPanel>().First(f => Grid.GetRow(f) == m_GridClient_Row && Grid.GetColumn(f) == m_GridClient_Column - 1);
                    for (int iChild = 0; iChild < stack.Children.Count; ++iChild)
                    {
                        Button childButton = (Button)stack.Children[iChild];
                        childButton.Background = m_Color_SelectedHostAndHostess;
                        if (childButton.Tag.ToString() != "" && childButton.Tag.ToString() != "CellPhone" && childButton.Tag.ToString() != "Phone" && childButton.Tag.ToString() != "Email")
                        {
                            m_Button_Client_SelectedClient = childButton;
                            m_Id_SelectedClient = (string)childButton.Tag;
                        }
                        else if (childButton.Tag.ToString() == "CellPhone")
                        {
                            m_Button_Client_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_Client_SelectedEmail = childButton;
                        }

                        //Display the last hostess
                        ScrollViewer_Clients_Details.ScrollToBottom();
                    }

                    //Enable the buttons
                    Btn_Clients_Archive.IsEnabled = true;
                    Btn_Clients_Edit.IsEnabled = true;
                    Btn_Clients_Delete.IsEnabled = true;

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on delete client
        /// </summary>
        private void Btn_Clients_Delete_Click(object sender, RoutedEventArgs e)
        {

            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the client
                Client client = Get_SelectedClientFromButton();
                if (client == null)
                {
                    return;
                }

                //Verification of the presence of missions or invoices
                List<Mission> associatedMissions = new List<Mission>();
                for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
                {
                    Mission billSel = SoftwareObjects.MissionsCollection[iMission];
                    if (billSel.id == client.id) // TODO
                    {
                        associatedMissions.Add(billSel);
                    }
                }
                if (associatedMissions.Count > 0)
                {
                    string message = m_Global_Handler.Resources_Handler.Get_Resources("ClientForbiddenDeleteText") + "\n";
                    message += m_Global_Handler.Resources_Handler.Get_Resources("Mission") + ": ";
                    for (int iMission = 0; iMission < associatedMissions.Count; ++iMission)
                    {
                        message += " " + associatedMissions[iMission].id;
                    }
                    MessageBox.Show(this, message, m_Global_Handler.Resources_Handler.Get_Resources("ClientForbiddenDeleteCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                //Confirm the delete
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("ClientConfirmDeleteText"),
                    m_Global_Handler.Resources_Handler.Get_Resources("ClientConfirmDeleteCaption"), MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                //Edit in internet database
                string res = m_Database_Handler.Delete_ClientFromDatabase(client.id);

                //Treat the result
                if (res.Contains("OK"))
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "ClientDeletionPrincipalMessage", "ClientDeletionSecondaryMessage");

                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Client " + client.corporate_name + " deleted");

                    //Delete the hostess to the collection
                    SoftwareObjects.ClientsCollection.Remove(client);
                    Grid_Clients_Details.Children.Clear();
                    Actualize_GridClientsFromDatabase();

                    //Delete photo repository
                    string repository = SoftwareObjects.GlobalSettings.photos_repository + "\\" + client.id;
                    if (Directory.Exists(repository))
                    {
                        Directory.Delete(repository, true);
                    }

                    //Clear the button
                    m_Button_Client_SelectedClient = null;
                    m_Button_Client_SelectedEmail = null;
                    m_Button_Client_SelectedCellPhone = null;

                    //Disable the buttons
                    Btn_Clients_Archive.IsEnabled = false;
                    Btn_Clients_Edit.IsEnabled = false;
                    Btn_Clients_Delete.IsEnabled = false;

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on edit client
        /// </summary>
        private void Btn_Clients_Edit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Get the hostess
                Client client = Get_SelectedClientFromButton();
                if (client == null)
                {
                    return;
                }

                //Open the window
                WindowClient.MainWindow_Client client_EditWindow = new WindowClient.MainWindow_Client(m_Global_Handler, m_Database_Handler,
                    m_Cities_DocFrance, true, client);
                Nullable<bool> resShow = client_EditWindow.ShowDialog();

                //Validation
                if (resShow == true)
                {
                    //Modify the hostess
                    client = client_EditWindow.m_Client;

                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Client " + client.corporate_name + " edited");

                    //Treatment of client info
                    TextBlock clientInfo = new TextBlock();
                    Run names = new Run("\n" + client.corporate_name);
                    names.FontSize = 15;
                    names.FontWeight = FontWeights.Bold;
                    clientInfo.Inlines.Add(names);
                    clientInfo.Inlines.Add(new LineBreak());
                    string creationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(client.date_creation, m_Global_Handler.Language_Handler);
                    string info = client.address + "\n" +
                        client.zipcode + ", " + client.city + "\n" +
                        m_Global_Handler.Resources_Handler.Get_Resources("CreationDate") + " : " + creationDate + "\n";
                    clientInfo.Inlines.Add(info);
                    //Create stack button
                    StackPanel buttonStack = new StackPanel();
                    buttonStack.Orientation = Orientation.Horizontal;

                    //Add void
                    Label voidLabel = new Label();
                    voidLabel.Content = "\t";
                    buttonStack.Children.Add(voidLabel);

                    //Add label
                    Label buttonLabel = new Label();
                    buttonLabel.Content = clientInfo;
                    buttonStack.Children.Add(buttonLabel);

                    //Add button
                    m_Button_Client_SelectedClient.Content = buttonStack;

                    //Get buttons
                    StackPanel panel = (StackPanel)m_Button_Client_SelectedClient.Parent;
                    Button buttonCellPhone = null;
                    for (int iChild = 1; iChild < panel.Children.Count; ++iChild)
                    {
                        Button childButton = (Button)panel.Children[iChild];
                        if (childButton.Tag.ToString() == "CellPhone")
                        {
                            buttonCellPhone = childButton;
                            m_Button_Client_SelectedCellPhone = childButton;
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
                    string cellPhone = client.phone;
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
                }

                return;
            }
            catch (Exception exception)
            {
                //Write the error into log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on generate an excel staement of clients
        /// /// </summary>
        private void Btn_Clients_GenerateExcelStatement_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Get the name of the pdf
                string today = m_Global_Handler.DateAndTime_Handler.Treat_Date(DateTime.Today.ToString("yyyy-MM-dd"), m_Global_Handler.Language_Handler);
                today = today.Replace("/", "-");
                string fileNameXLS = m_Global_Handler.Resources_Handler.Get_Resources("ClientsStatement") + " - " + today;
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
                windowWait.Start(m_Global_Handler, "ClientsExcelGenerationPrincipalMessage", "ClientsExcelGenerationSecondaryMessage");

                //Sort the collection
                SoftwareObjects.ClientsCollection.Sort(delegate (Client x, Client y)
                {
                    if (x.corporate_name == null && y.corporate_name == null) return 0;
                    else if (x.corporate_name == null) return 1;
                    else if (y.corporate_name == null) return -1;
                    else
                    {
                        return x.corporate_name.CompareTo(y.corporate_name);
                    }
                });

                //Write in a file with tab separated
                List<string> lines = new List<string>();
                string columnsHeader = m_Global_Handler.Resources_Handler.Get_Resources("CorporateName") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CorporateNumber") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("VATNumber") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Phone") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Address") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("ZipCode") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("City") + "\t" +
                    m_Global_Handler.Resources_Handler.Get_Resources("Country");
                lines.Add(columnsHeader);
                for (int iClient = 0; iClient < m_DisplayedClientsCollection.Count; ++iClient)
                {
                    Client client = m_DisplayedClientsCollection[iClient];
                    string line = client.corporate_name + "\t" + client.corporate_number + "\t" + client.vat_number + "\t" + client.phone + "\t" + client.address + "\t" + client.zipcode + "\t" +
                       client.city + "\t" + client.country;
                    lines.Add(line);
                }

                //Generate the excel file
                int result = m_Global_Handler.Excel_Handler.Generate_ExcelStatement(fileNameXLS, lines);
                if (result == -1)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("ClientsStatementExcelGenerationFailedText"),
                                m_Global_Handler.Resources_Handler.Get_Resources("ClientsStatementExcelGenerationFailedCaption"),
                                MessageBoxButton.OK, MessageBoxImage.Error);
                    //Close the wait window
                    windowWait.Stop();

                    return;
                }

                //Action
                m_Global_Handler.Log_Handler.WriteAction("Clients statement generated to Excel file " + fileNameXLS);

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on show archived clients
        /// </summary>
        private void Btn_Clients_ShowArchived_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window wait
                windowWait.Start(m_Global_Handler, "ClientsShowArchivedPrincipalMessage", "ClientsShowArchivedSecondaryMessage");

                //Mode
                m_Clients_IsArchiveMode = true;

                //Actualize
                Actualize_GridClientsFromDatabase();

                //Clear the fields
                Txt_Clients_Research.Text = "";
                m_DataGrid_Clients_MissionsCollection.Clear();
                DataGrid_Clients_Missions.Items.Refresh();

                //Null buttons
                m_Button_Client_SelectedCellPhone = null;
                m_Button_Client_SelectedEmail = null;
                m_Button_Client_SelectedClient = null;

                //Disable the buttons
                Btn_Clients_Archive.IsEnabled = false;
                Btn_Clients_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("Restore");
                Btn_Clients_Create.IsEnabled = false;
                Btn_Clients_Edit.IsEnabled = false;
                Btn_Clients_Delete.IsEnabled = false;

                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Click on show in progress clients
        /// </summary>
        private void Btn_Clients_ShowInProgress_Click(object sender, RoutedEventArgs e)
        {
            //Creation of the wait window
            WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

            try
            {
                //Open the window wait
                windowWait.Start(m_Global_Handler, "ClientsShowInProgressPrincipalMessage", "ClientsShowInProgressSecondaryMessage");

                //Mode
                m_Clients_IsArchiveMode = false;

                //Actualize
                Actualize_GridClientsFromDatabase();

                //Clear the fields
                Txt_Clients_Research.Text = "";
                m_DataGrid_Clients_MissionsCollection.Clear();
                DataGrid_Clients_Missions.Items.Refresh();

                //Null buttons
                m_Button_Client_SelectedCellPhone = null;
                m_Button_Client_SelectedEmail = null;
                m_Button_Client_SelectedClient = null;

                //Disable the buttons
                Btn_Clients_Archive.IsEnabled = false;
                Btn_Clients_Archive.Content = m_Global_Handler.Resources_Handler.Get_Resources("Archive");
                Btn_Clients_Create.IsEnabled = false;
                Btn_Clients_Edit.IsEnabled = false;
                Btn_Clients_Delete.IsEnabled = false;

                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();
            }
            catch (Exception exception)
            {
                //Close the wait window
                Thread.Sleep(500);
                windowWait.Stop();

                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
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
            m_IsSortClient_Ascending = !m_IsSortClient_Ascending;

            //Save in case of problems
            List<Client> SavedCollection = new List<Client>();
            SavedCollection = SoftwareObjects.ClientsCollection;

            //No sort
            if (Cmb_Clients_SortBy.SelectedIndex == -1)
            {
                //Clear the grid
                Grid_Clients_Details.Children.Clear();
                Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                //Actualize the collection
                Actualize_GridClientsFromCollection();
            }
            else if (Cmb_Clients_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("CorporateName"))
            {
                //Sort by corporate name
                try
                {
                    SoftwareObjects.ClientsCollection.Sort(delegate (Client x, Client y)
                    {
                        if (x.corporate_name == null && y.corporate_name == null) return 0;
                        else if (x.corporate_name == null) return -1;
                        else if (y.corporate_name == null) return 1;
                        else
                        {
                            if (m_IsSortClient_Ascending == true)
                            {
                                return x.corporate_name.CompareTo(y.corporate_name);
                            }
                            else
                            {
                                return y.corporate_name.CompareTo(x.corporate_name);
                            }
                        }
                    });

                    //Clear the grid
                    Grid_Clients_Details.Children.Clear();
                    Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridClientsFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.ClientsCollection = SavedCollection;

                    //Clear the grid
                    Grid_Clients_Details.Children.Clear();
                    Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridClientsFromCollection();

                    return;
                }
            }
            else if (Cmb_Clients_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("ZipCode"))
            {
                //Sort by last name
                try
                {
                    SoftwareObjects.ClientsCollection.Sort(delegate (Client x, Client y)
                    {
                        if (x.zipcode == null && y.zipcode == null) return 0;
                        else if (x.zipcode == null) return -1;
                        else if (y.zipcode == null) return 1;
                        else
                        {
                            if (m_IsSortClient_Ascending == true)
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
                    Grid_Clients_Details.Children.Clear();
                    Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridClientsFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.ClientsCollection = SavedCollection;

                    //Clear the grid
                    Grid_Clients_Details.Children.Clear();
                    Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridClientsFromCollection();

                    return;
                }
            }
            else if (Cmb_Clients_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("City"))
            {
                //Sort by city
                try
                {
                    SoftwareObjects.ClientsCollection.Sort(delegate (Client x, Client y)
                    {
                        if (x.city == null && y.city == null) return 0;
                        else if (x.city == null) return -1;
                        else if (y.city == null) return 1;
                        else
                        {
                            if (m_IsSortClient_Ascending == true)
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
                    Grid_Clients_Details.Children.Clear();
                    Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridClientsFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.ClientsCollection = SavedCollection;

                    //Clear the grid
                    Grid_Clients_Details.Children.Clear();
                    Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridClientsFromCollection();

                    return;
                }
            }
            else if (Cmb_Clients_SortBy.SelectedItem.ToString() == m_Global_Handler.Resources_Handler.Get_Resources("CreationDate"))
            {
                //Sort by creation date
                try
                {
                    SoftwareObjects.ClientsCollection.Sort(delegate (Client x, Client y)
                    {
                        if (x.date_creation == null && y.date_creation == null) return 0;
                        else if (x.date_creation == null) return -1;
                        else if (y.date_creation == null) return 1;
                        else
                        {
                            if (m_IsSortClient_Ascending == true)
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
                    Grid_Clients_Details.Children.Clear();
                    Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridClientsFromCollection();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.ClientsCollection = SavedCollection;

                    //Clear the grid
                    Grid_Clients_Details.Children.Clear();
                    Grid_Clients_Details.RowDefinitions.RemoveRange(0, Grid_Clients_Details.RowDefinitions.Count - 1);

                    //Actualize the collection
                    Actualize_GridClientsFromCollection();

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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Mouse down on the main client button event
        /// Select the client
        /// </summary>
        private void IsPressed_Clients_Button(object sender, RoutedEventArgs ev)
        {
            if (sender != null)
            {
                try
                {
                    //Get the hostess's id
                    Button buttonSel = (Button)sender;
                    string idSel = (string)buttonSel.Tag;
                    m_Id_SelectedClient = idSel;

                    //Manage the buttons
                    if (m_Button_Client_SelectedClient != null)
                    {
                        if (m_Clients_IsArchiveMode == false)
                        {
                            m_Button_Client_SelectedClient.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_Client_SelectedClient.Background = m_Color_ArchivedMission;
                        }
                    }
                    if (m_Button_Client_SelectedCellPhone != null)
                    {
                        if (m_Clients_IsArchiveMode == false)
                        {
                            m_Button_Client_SelectedCellPhone.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_Client_SelectedCellPhone.Background = m_Color_ArchivedMission;
                        }
                    }
                    if (m_Button_Client_SelectedEmail != null)
                    {
                        if (m_Clients_IsArchiveMode == false)
                        {
                            m_Button_Client_SelectedEmail.Background = m_Color_HostAndHostess;
                        }
                        else
                        {
                            m_Button_Client_SelectedEmail.Background = m_Color_ArchivedMission;
                        }
                    }
                    m_Button_Client_SelectedClient = null;
                    m_Button_Client_SelectedCellPhone = null;
                    m_Button_Client_SelectedEmail = null;
                    StackPanel stackSel = (StackPanel)buttonSel.Parent;
                    for (int iChild = 0; iChild < stackSel.Children.Count; ++iChild)
                    {
                        Button childButton = (Button)stackSel.Children[iChild];
                        childButton.Background = m_Color_SelectedHostAndHostess;
                        if (iChild == 0)
                        {
                            m_Button_Client_SelectedClient = childButton;
                        }
                        else if (childButton.Tag.ToString() == "CellPhone")
                        {
                            m_Button_Client_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_Client_SelectedEmail = childButton;
                        }
                    }
                    //Select the hostess
                    Get_SelectedClientFromButton();

                    if (m_Clients_IsArchiveMode == false)
                    {
                        //Enable the buttons
                        Btn_Clients_Archive.IsEnabled = true;
                        Btn_Clients_Edit.IsEnabled = true;
                        Btn_Clients_Delete.IsEnabled = true;
                    }
                    else
                    {
                        //Disable the buttons
                        Btn_Clients_Archive.IsEnabled = true;
                        Btn_Clients_Edit.IsEnabled = false;
                        Btn_Clients_Delete.IsEnabled = false;
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Mouse down on the client's email button event
        /// Select client and open a new email with the selected client email adress
        /// </summary>
        private void IsPressed_Clients_EmailButton(object sender, RoutedEventArgs ev)
        {
            if (sender != null)
            {
                //Creation of the wait window
                WindowWait.MainWindow_Wait windowWait = new WindowWait.MainWindow_Wait();

                try
                {
                    //Open the wait window
                    windowWait.Start(m_Global_Handler, "ClientEmailPrincipalMessage", "ClientEmailSecondaryMessage");

                    //Get the email adress
                    Button emailButton = (Button)sender;
                    StackPanel emailStackPanel = (StackPanel)emailButton.Content;
                    TextBlock emailTextBox = (TextBlock)emailStackPanel.Children[1];
                    string emailAddress = emailTextBox.Text;

                    //Manage the buttons
                    if (m_Button_Client_SelectedClient != null)
                    {
                        m_Button_Client_SelectedClient.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_Client_SelectedCellPhone != null)
                    {
                        m_Button_Client_SelectedCellPhone.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_Client_SelectedEmail != null)
                    {
                        m_Button_Client_SelectedEmail.Background = m_Color_HostAndHostess;
                    }
                    m_Button_Client_SelectedClient = null;
                    m_Button_Client_SelectedCellPhone = null;
                    m_Button_Client_SelectedEmail = null;
                    StackPanel stackSel = (StackPanel)emailButton.Parent;
                    for (int iChildren = 0; iChildren < stackSel.Children.Count; ++iChildren)
                    {
                        Button childButton = (Button)stackSel.Children[iChildren];
                        childButton.Background = m_Color_SelectedHostAndHostess;
                        if (childButton.Tag.ToString() != "" && childButton.Tag.ToString() != "CellPhone" && childButton.Tag.ToString() != "Email")
                        {
                            m_Button_Client_SelectedClient = childButton;
                            m_Id_SelectedClient = (string)childButton.Tag;
                        }
                        else if (childButton.Tag.ToString() == "CellPhone")
                        {
                            m_Button_Client_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_Client_SelectedEmail = childButton;
                        }
                    }

                    //Select the hostess
                    Get_SelectedClientFromButton();

                    if (m_Clients_IsArchiveMode == false)
                    {
                        //Enable the buttons
                        Btn_Clients_Archive.IsEnabled = true;
                        Btn_Clients_Edit.IsEnabled = true;
                        Btn_Clients_Delete.IsEnabled = true;

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
                        m_Global_Handler.Log_Handler.WriteAction("Mail sent to " + emailAddress);
                    }
                    else
                    {
                        //Disable the buttons
                        Btn_Clients_Archive.IsEnabled = true;
                        Btn_Clients_Edit.IsEnabled = false;
                        Btn_Clients_Delete.IsEnabled = false;
                    }

                    //Close stop window
                    windowWait.Stop();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    windowWait.Stop();
                    return;
                }
            }
        }

        /// <summary>
        /// Event
        /// Clients
        /// Mouse down on the client's phone button(s) event
        /// Select the client and open skype
        /// </summary>
        private void IsPressed_Clients_PhoneButton(object sender, RoutedEventArgs ev)
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
                    windowWait.Start(m_Global_Handler, "ClientPhonePrincipalMessage", "ClientPhoneSecondaryMessage");

                    //Get the email adress
                    Button phoneButton = (Button)sender;
                    StackPanel phoneStackPanel = (StackPanel)phoneButton.Content;
                    TextBlock phoneTextBox = (TextBlock)phoneStackPanel.Children[1];
                    string phoneNumber = phoneTextBox.Text.Trim();

                    //Manage the buttons
                    if (m_Button_Client_SelectedClient != null)
                    {
                        m_Button_Client_SelectedClient.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_Client_SelectedCellPhone != null)
                    {
                        m_Button_Client_SelectedCellPhone.Background = m_Color_HostAndHostess;
                    }
                    if (m_Button_Client_SelectedEmail != null)
                    {
                        m_Button_Client_SelectedEmail.Background = m_Color_HostAndHostess;
                    }
                    m_Button_Client_SelectedClient = null;
                    m_Button_Client_SelectedCellPhone = null;
                    m_Button_Client_SelectedEmail = null;
                    StackPanel stackSel = (StackPanel)phoneButton.Parent;
                    for (int iChildren = 0; iChildren < stackSel.Children.Count; ++iChildren)
                    {
                        Button childButton = (Button)stackSel.Children[iChildren];
                        childButton.Background = m_Color_SelectedHostAndHostess;
                        if (childButton.Tag.ToString() != "" && childButton.Tag.ToString() != "CellPhone" && childButton.Tag.ToString() != "Email")
                        {
                            m_Button_Client_SelectedClient = childButton;
                            m_Id_SelectedClient = (string)childButton.Tag;
                        }
                        else if (childButton.Tag.ToString() == "CellPhone")
                        {
                            m_Button_Client_SelectedCellPhone = childButton;
                        }
                        else if (childButton.Tag.ToString() == "Email")
                        {
                            m_Button_Client_SelectedEmail = childButton;
                        }
                    }

                    if (m_Clients_IsArchiveMode == false)
                    {
                        //Enable the buttons
                        Btn_Clients_Archive.IsEnabled = true;
                        Btn_Clients_Edit.IsEnabled = true;
                        Btn_Clients_Delete.IsEnabled = true;

                        //Select the hostess
                        Client client = Get_SelectedClientFromButton();

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

                        //Action
                        m_Global_Handler.Log_Handler.WriteAction("Skype call made to " + client.corporate_name + ", " + client.phone);
                    }
                    else
                    {
                        //Disable the buttons
                        Btn_Clients_Archive.IsEnabled = true;
                        Btn_Clients_Edit.IsEnabled = false;
                        Btn_Clients_Delete.IsEnabled = false;
                    }

                    //Close stop window
                    windowWait.Stop();

                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    windowWait.Stop();
                    return;
                }
            }
        }

        /// <summary>
        /// Boolean tagging the full actualization avoiding to do it multiple time
        /// This variable is only used in the research client text box text changed event
        /// </summary>
        bool actualizationClientDone = true;
        /// <summary>
        /// Event
        /// Clients
        /// Text changed in the research client text box
        /// Only Hostess containing the text (at least 2 characters) are shown 
        /// </summary>
        private void Txt_Clients_Research_TextChanged(object sender, TextChangedEventArgs e)
        {
            //Get the researched text
            string researchedText = Txt_Clients_Research.Text;
            researchedText = researchedText.ToLower();

            //Clear the fields
            m_Button_Client_SelectedClient = null;

            //Verifications - 3 letters minimum
            if (researchedText.Length < 2)
            {
                if (!actualizationClientDone)
                {
                    Grid_Clients_Details.Children.Clear();
                    Actualize_GridClientsFromDatabase();
                    actualizationClientDone = true;
                    return;
                }
            }
            else
            {
                try
                {
                    //Research in each private member of the list of Client
                    List<Client> foundClientList = new List<Client>();
                    for (int iClient = 0; iClient < SoftwareObjects.ClientsCollection.Count; ++iClient)
                    {
                        Client processedClient = SoftwareObjects.ClientsCollection[iClient];
                        if (processedClient.address.ToLower().Contains(researchedText))
                        {
                            foundClientList.Add(processedClient);
                            continue;
                        }
                        else if (processedClient.city.ToLower().Contains(researchedText))
                        {
                            foundClientList.Add(processedClient);
                            continue;
                        }
                        else if (processedClient.country.ToLower().Contains(researchedText))
                        {
                            foundClientList.Add(processedClient);
                            continue;
                        }
                        else if (processedClient.corporate_name.ToLower().Contains(researchedText))
                        {
                            foundClientList.Add(processedClient);
                            continue;
                        }
                        else if (processedClient.state.ToLower().Contains(researchedText))
                        {
                            foundClientList.Add(processedClient);
                            continue;
                        }
                        else if (processedClient.zipcode.ToLower().Contains(researchedText))
                        {
                            foundClientList.Add(processedClient);
                            continue;
                        }
                    }

                    //Clear displayed collection
                    m_DisplayedClientsCollection.Clear();

                    //Displaying the found Client list
                    Grid_Clients_Details.Children.Clear();
                    actualizationClientDone = false;
                    if (foundClientList.Count > 0)
                    {
                        //Clear fields
                        Txt_Clients_Research.Text = "";

                        //Initialize counter for columns and rows of Grid_Client
                        m_GridClient_Column = 0;
                        m_GridClient_Row = 0;
                    }
                    for (int iClient = 0; iClient < foundClientList.Count; ++iClient)
                    {
                        Add_ClientToGrid(foundClientList[iClient]);
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, Arg_Ex);
                    MessageBox.Show(Arg_Ex.Message, m_Global_Handler.Resources_Handler.Get_Resources("DatabaseDefinitionModificationErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                    Txt_Settings_General_Database.Text = SoftwareObjects.GlobalSettings.database_definition;
                    return;
                }
                catch (MySqlException SQL_Ex)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, SQL_Ex);
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
                m_Global_Handler.Log_Handler.WriteAction("Settings: Database " + SoftwareObjects.GlobalSettings.database_definition + " saved");
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                                m_Global_Handler.Log_Handler.WriteAction("Settings: Photos saved to " + SoftwareObjects.GlobalSettings.photos_repository);
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                m_GridMission_Column = 0;
                m_GridMission_Row = 0;

                //Actualization
                try
                {
                    //Clear displayed collection
                    m_DisplayedMissionsCollection.Clear();

                    Grid_Missions_Details.Children.Clear();
                    Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);
                    for (int iMission = 0; iMission < m_Grid_Details_Missions_MissionsCollection.Count; ++iMission)
                    {
                        //Get mission
                        Mission mission = m_Grid_Details_Missions_MissionsCollection[iMission];
                        //Show archive or not
                        if (m_Mission_IsArchiveMode == true && mission.archived == 1)
                        {
                            Add_MissionToGrid(mission);
                        }
                        if (m_Mission_IsArchiveMode == false && mission.archived == 0)
                        {
                            Add_MissionToGrid(mission);
                        }
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                //Initialize rows and columns
                m_GridMission_Column = 0;
                m_GridMission_Row = 0;
                Grid_Missions_Details.Children.Clear();

                //Clear displayed collection
                m_DisplayedMissionsCollection.Clear();

                //Getting the hosts and hostesses collection  
                try
                {
                    string res = m_Database_Handler.Get_MissionsFromDatabase();
                    if (res.Contains("OK"))
                    {
                        //Add to grid
                        for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
                        {
                            Mission mission = SoftwareObjects.MissionsCollection[iMission];
                            if (m_Mission_IsArchiveMode == true && mission.archived == 1)
                            {
                                Add_MissionToGrid(mission);
                            }
                            else if (m_Mission_IsArchiveMode == false && mission.archived == 0)
                            {
                                Add_MissionToGrid(mission);
                            }
                        }
                    }
                    else if (res.Contains("Error"))
                    {
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("MissionsActualizationErrorCaption"),
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        SoftwareObjects.MissionsCollection = new List<Mission>();
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.MissionsCollection = new List<Mission>();
                    return;
                }
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Increment on grid missions columns
        /// </summary>
        private int m_GridMission_Column = 0;
        /// <summary>
        /// Functions
        /// Missions
        /// Increment on grid missions rows
        /// </summary>
        private int m_GridMission_Row = 0;
        /// <summary>
        /// Functions
        /// Missions
        /// Add an mission to the grid, create a stack panel containing all the information of the mission
        /// <param name="_Mission">Mission</param>
        /// </summary>
        private void Add_MissionToGrid(Mission _Mission)
        {
            try
            {
                //Treatment of the first row
                Grid_Missions_Details.RowDefinitions[0].MinHeight = 200;

                //Add to displayed collection
                m_DisplayedMissionsCollection.Add(_Mission);

                //Create the buttons
                Button buttonMission = new Button();
                Manage_MissionButton(_Mission, buttonMission, true);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                Mission missionSel = Get_SelectedMissionFromButton(false);

                //Initialize rows and columns
                m_GridMission_Column = 0;
                m_GridMission_Row = 0;

                //Initialize status
                m_Mission_SelectedStatus = _Status;

                //Get status
                List<Mission> createdMissions = new List<Mission>();
                List<Mission> generatedMissions = new List<Mission>();
                List<Mission> sentMissions = new List<Mission>();
                List<Mission> acceptedMissions = new List<Mission>();
                List<Mission> declinedMissions = new List<Mission>();
                List<Mission> billedMissions = new List<Mission>();
                for (int iMission = 0; iMission < m_Grid_Details_Missions_MissionsCollection.Count; ++iMission)
                {
                    //TODO
                    Mission mission = m_Grid_Details_Missions_MissionsCollection[iMission];
                    //if (mission.date_invoice_creation != "" && mission.date_invoice_creation != null && mission.date_invoice_creation != "0000-00-00")
                    //{
                    //    billedMissions.Add(mission);
                    //}
                    //else if (mission.date_Mission_accepted != "" && mission.date_Mission_accepted != null && mission.date_Mission_accepted != "0000-00-00")
                    //{
                    //    acceptedMissions.Add(mission);
                    //}
                    //else if (mission.date_Mission_declined != "" && mission.date_Mission_declined != null && mission.date_Mission_declined != "0000-00-00")
                    //{
                    //    declinedMissions.Add(mission);
                    //}
                    //else if (mission.date_Mission_sent != "" && mission.date_Mission_sent != null && mission.date_Mission_sent != "0000-00-00")
                    //{
                    //    sentMissions.Add(mission);
                    //}
                    //else if (mission.date_Mission_generated != "" && mission.date_Mission_generated != null && mission.date_Mission_generated != "0000-00-00")
                    //{
                    //    generatedMissions.Add(mission);
                    //}
                    //else
                    //{
                    createdMissions.Add(mission);
                    //}
                }

                //Buttons
                Btn_Missions_Legend_Mission_Done.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Billed.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Created.Background = m_Color_Button;
                Btn_Missions_Legend_Mission_Declined.Background = m_Color_Button;

                //Get filtered collection
                List<Mission> filteredMissions = new List<Mission>();
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

                //Clear displayed collection
                m_DisplayedMissionsCollection.Clear();

                //Filter
                Grid_Missions_Details.Children.Clear();
                Grid_Missions_Details.RowDefinitions.RemoveRange(0, Grid_Missions_Details.RowDefinitions.Count - 1);
                Mission missionFound = null;
                for (int iMission = 0; iMission < filteredMissions.Count; ++iMission)
                {
                    //Get mission
                    Mission mission = filteredMissions[iMission];
                    //Show archive or not
                    if (m_Mission_IsArchiveMode == true && mission.archived == 1)
                    {
                        Add_MissionToGrid(mission);
                    }
                    if (m_Mission_IsArchiveMode == false && mission.archived == 0)
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Get a mission from its button
        /// <returns>Returns the mission corresponding to the button</returns>
        /// </summary>
        private Mission Get_SelectedMissionFromButton(bool _ShowMessage = true)
        {
            try
            {
                //Tests
                if (m_Button_Mission_SelectedMission == null)
                {
                    if (_ShowMessage == true)
                    {
                        MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissionSelectedError"), m_Global_Handler.Resources_Handler.Get_Resources("MissionSelectedErrorCaption"),
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    return null;
                }
                StackPanel stack = (StackPanel)m_Button_Mission_SelectedMission.Parent;
                if (stack.Tag == null)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissionSelectedError"), m_Global_Handler.Resources_Handler.Get_Resources("MissionSelectedErrorCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return null;
                }

                //Get the mission from the id
                string id = (string)stack.Tag;
                Mission mission = SoftwareObjects.MissionsCollection.Find(x => x.id.Equals(id));

                return mission;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return null;
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Manage the mission button, create and add or modify
        /// </summary>
        private void Manage_MissionButton(Mission _Mission, Button _MissionButton, bool _Is_NewButton)
        {
            try
            {
                //Treatment of dates
                string creationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_creation, m_Global_Handler.Language_Handler);
                string startDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.start_date, m_Global_Handler.Language_Handler);
                string endDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.end_date, m_Global_Handler.Language_Handler);

                //Treatment of mission button
                TextBlock textInfo = new TextBlock();
                Run names = new Run("\n" + _Mission.client_name);
                names.FontSize = 15;
                names.FontWeight = FontWeights.Bold;
                textInfo.Inlines.Add(names);
                textInfo.Inlines.Add(new LineBreak());
                string[] descriptionTable = _Mission.description.Split(' ');
                string descriptionStr = "\n";
                for (int iWord = 0; iWord < descriptionTable.Length; ++iWord)
                {
                    descriptionStr = descriptionStr + descriptionTable[iWord] + " ";
                    if (iWord % 4 == 0 && iWord != 0)
                    {
                        descriptionStr = descriptionStr + "\n";
                    }
                }
                Run description = new Run(descriptionStr);
                description.FontSize = 14;
                description.FontWeight = FontWeights.Bold;
                textInfo.Inlines.Add(description);
                textInfo.Inlines.Add(new LineBreak());
                Run dates = new Run("\n" + startDate + "\t" + endDate);
                dates.FontSize = 13;
                textInfo.Inlines.Add(dates);
                textInfo.Inlines.Add(new LineBreak());
                string info = _Mission.address + "\n" +
                    _Mission.zipcode + ", " + _Mission.city + "\n" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CreationDate") + " : " + creationDate + "\n";
                textInfo.Inlines.Add(info);
                textInfo.TextWrapping = TextWrapping.Wrap;

                //Create stack button
                StackPanel buttonStack = new StackPanel();
                buttonStack.Orientation = Orientation.Horizontal;

                //Create image TODO
                Image imgMission = new Image();
                //if (_Mission.num_invoice != null && _Mission.num_invoice != "")
                //{
                //    //Related invoice created
                //    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Billed);
                //}
                //else if (_Mission.date_Mission_accepted != null && _Mission.date_Mission_accepted != "")
                //{
                //    //Mission accepted
                //    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Accepted);
                //}
                //else if (_Mission.date_Mission_declined != null && _Mission.date_Mission_declined != "")
                //{
                //    //Mission declined
                //    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Declined);
                //}
                //else if (_Mission.date_Mission_sent != null && _Mission.date_Mission_sent != "")
                //{
                //    //Mission sent by email icon
                //    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Sent);
                //}
                //else if (_Mission.date_Mission_generated != null && _Mission.date_Mission_generated != "")
                //{
                //    //Mission generated
                //    imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Generated);
                //}
                //else
                //{
                //Default icon
                imgMission.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Icon_Mission_Created);
                //}

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
                    if (_Mission.archived == 0)
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

                    //Treatment of the increments
                    if (m_GridMission_Column > Grid_Missions_Details.ColumnDefinitions.Count - 1)
                    {
                        m_GridMission_Column = 0;
                        m_GridMission_Row += 1;
                        RowDefinition row = new RowDefinition();
                        row.MinHeight = 200;
                        Grid_Missions_Details.RowDefinitions.Add(row);
                        Grid.SetColumn(stackPanel, m_GridMission_Column);
                        Grid.SetRow(stackPanel, m_GridMission_Row);
                        _MissionButton.Tag = m_GridMission_Row.ToString() + "|" + m_GridMission_Column.ToString();
                        m_GridMission_Column += 1;
                    }
                    else
                    {
                        Grid.SetColumn(stackPanel, m_GridMission_Column);
                        Grid.SetRow(stackPanel, m_GridMission_Row);
                        _MissionButton.Tag = m_GridMission_Row.ToString() + "|" + m_GridMission_Column.ToString();
                        m_GridMission_Column += 1;
                    }

                    //Treatment of the Grid_Mission column
                    Grid_Missions_Details.Children.Add(stackPanel);
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Select an mission
        /// </summary>
        private void Select_Mission(Mission _Mission)
        {
            try
            {
                //Verification
                if (_Mission == null)
                {
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
                //Get the mission stack panel
                double row = 0;
                for (int iMission = 0; iMission < Grid_Missions_Details.Children.Count; ++iMission)
                {
                    StackPanel stackSel = (StackPanel)Grid_Missions_Details.Children[iMission];
                    if ((string)stackSel.Tag == _Mission.id)
                    {
                        m_Button_Mission_SelectedMission = (Button)stackSel.Children[0];
                        m_Button_Mission_SelectedMission.Background = colorSelectedMission;
                        stackSel.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                        stackSel.Arrange(new Rect(0, 0, stackSel.DesiredSize.Width, stackSel.DesiredSize.Height));
                        row = Grid.GetRow(stackSel) * stackSel.ActualHeight;
                        break;
                    }
                }

                //Hostess text box
                Txt_Missions_Client.Text = _Mission.client_name;
                Txt_Missions_CreationDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.date_creation, m_Global_Handler.Language_Handler);
                Txt_Missions_EndDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.end_date, m_Global_Handler.Language_Handler);
                Txt_Missions_StartDate.Text = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Mission.start_date, m_Global_Handler.Language_Handler);

                //Scroll to the selected invoice
                Scrl_Grid_Missions_Details.ScrollToVerticalOffset(row);

                //Enable the buttons
                if (_Mission.archived == 0)
                {
                    Btn_Missions_Duplicate.IsEnabled = true;
                    Btn_Missions_Edit.IsEnabled = true;
                    Btn_Missions_Delete.IsEnabled = true;
                    Btn_Missions_Archive.IsEnabled = true;
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Missions
        /// Sort the grid of missions by status
        /// </summary>
        private void Sort_MissionsByStatus(List<Mission> _SavedCollection)
        {
            try
            {
                //Get status
                List<Mission> createdMissions = new List<Mission>();
                List<Mission> generatedMissions = new List<Mission>();
                List<Mission> sendMissions = new List<Mission>();
                List<Mission> acceptedMissions = new List<Mission>();
                List<Mission> declinedMissions = new List<Mission>();
                List<Mission> billedMissions = new List<Mission>();
                for (int iMission = 0; iMission < m_Grid_Details_Missions_MissionsCollection.Count; ++iMission)
                {
                    //TODO
                    //Mission mission = m_Grid_Details_Missions_MissionsCollection[iMission];
                    //if (mission.date_invoice_creation != "" && mission.date_invoice_creation != null)
                    //{
                    //    billedMissions.Add(mission);
                    //}
                    //else if (mission.date_Mission_accepted != "" && mission.date_Mission_accepted != null)
                    //{
                    //    acceptedMissions.Add(mission);
                    //}
                    //else if (mission.date_Mission_declined != "" && mission.date_Mission_declined != null)
                    //{
                    //    declinedMissions.Add(mission);
                    //}
                    //else if (mission.date_Mission_sent != "" && mission.date_Mission_sent != null)
                    //{
                    //    sendMissions.Add(mission);
                    //}
                    //else if (mission.date_Mission_generated != "" && mission.date_Mission_generated != null)
                    //{
                    //    generatedMissions.Add(mission);
                    //}
                    //else
                    //{
                    //    createdMissions.Add(mission);
                    //}
                }
                //Sort each category by num mission
                List<List<Mission>> listCollections = new List<List<Mission>>();
                listCollections.Add(createdMissions);
                listCollections.Add(generatedMissions);
                listCollections.Add(sendMissions);
                listCollections.Add(declinedMissions);
                listCollections.Add(acceptedMissions);
                listCollections.Add(billedMissions);
                for (int iList = 0; iList < listCollections.Count; ++iList)
                {
                    List<Mission> list = listCollections[iList];
                    list.Sort(delegate (Mission x, Mission y)
                    {
                        return x.id.CompareTo(y.id);
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                Grid_HostAndHostess_Details.Children.Clear();

                //Actualization
                try
                {
                    //Clear displayed collection
                    m_DisplayedHostsAndHostessesCollection.Clear();

                    //Add to grid
                    for (int iHostess = 0; iHostess < SoftwareObjects.HostsAndHotessesCollection.Count; ++iHostess)
                    {
                        Hostess hostess = SoftwareObjects.HostsAndHotessesCollection[iHostess];
                        if (m_HostsAndHostesses_IsArchiveMode == true && hostess.archived == 1)
                        {
                            Add_HostOrHostessToGrid(hostess);
                        }
                        else if (m_HostsAndHostesses_IsArchiveMode == false && hostess.archived == 0)
                        {
                            Add_HostOrHostessToGrid(hostess);
                        }
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                Grid_HostAndHostess_Details.Children.Clear();

                //Clear displayed collection
                m_DisplayedHostsAndHostessesCollection.Clear();

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
                            if (m_HostsAndHostesses_IsArchiveMode == true && hostess.archived == 1)
                            {
                                Add_HostOrHostessToGrid(hostess);
                            }
                            else if (m_HostsAndHostesses_IsArchiveMode == false && hostess.archived == 0)
                            {
                                Add_HostOrHostessToGrid(hostess);
                            }
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
                //Add to visible collection
                m_DisplayedHostsAndHostessesCollection.Add(_HostOrHostess);

                //Treatment of the first row
                Grid_HostAndHostess_Details.RowDefinitions[0].MinHeight = 350;

                //Create the hostess button
                Button buttonHostess = new Button();
                buttonHostess.Padding = new Thickness(5);
                if (m_HostsAndHostesses_IsArchiveMode == false)
                {
                    buttonHostess.Background = m_Color_Mission;
                }
                else
                {
                    buttonHostess.Background = m_Color_ArchivedMission;
                }
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
                if (m_HostsAndHostesses_IsArchiveMode == false)
                {
                    buttonCallCellPhone.Background = m_Color_Mission;
                }
                else
                {
                    buttonCallCellPhone.Background = m_Color_ArchivedMission;
                }
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
                if (m_HostsAndHostesses_IsArchiveMode == false)
                {
                    buttonSendMail.Background = m_Color_Mission;
                }
                else
                {
                    buttonSendMail.Background = m_Color_ArchivedMission;
                }
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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
            try
            {
                //Test
                if (m_Button_HostAndHostess_SelectedHostAndHostess == null || m_Button_HostAndHostess_SelectedHostAndHostess.Tag == null)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostessSelectedError"),
                        m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostessSelectedErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
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
                    for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
                    {
                        Mission missionSel = SoftwareObjects.MissionsCollection[iMission];
                        if (missionSel.id == hostess.id) //TODO
                        {
                            m_Datagrid_HostAndHostess_Missions data = new m_Datagrid_HostAndHostess_Missions(missionSel.id, missionSel.client_name, missionSel.city, missionSel.start_date);
                            m_DataGrid_HostAndHostess_MissionsCollection.Add(data);
                        }
                    }
                    DataGrid_HostAndHostess_Missions.Items.Refresh();
                }

                //Return the hostess
                return hostess;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return null;
            }
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
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return -1;
            }
        }

        #endregion

        #region Clients

        /// <summary>
        /// Functions
        /// Clients
        /// Actualize the grid of Clients from the actual collection
        /// </summary>
        private void Actualize_GridClientsFromCollection()
        {
            if (m_Database_Handler != null)
            {
                //Initialize rows and columns
                m_GridClient_Column = 0;
                m_GridClient_Row = 0;
                Grid_Clients_Details.Children.Clear();

                //Clear displayed collection
                m_DisplayedClientsCollection.Clear();

                //Actualization
                try
                {
                    //Add to grid
                    for (int iClient = 0; iClient < SoftwareObjects.ClientsCollection.Count; ++iClient)
                    {
                        Client client = SoftwareObjects.ClientsCollection[iClient];
                        if (m_Clients_IsArchiveMode == true && client.archived == 1)
                        {
                            Add_ClientToGrid(client);
                        }
                        else if (m_Clients_IsArchiveMode == false && client.archived == 0)
                        {
                            Add_ClientToGrid(client);
                        }
                    }
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.ClientsCollection = new List<Client>();
                    return;
                }
            }
        }

        /// <summary>
        /// Functions
        /// Clients
        /// Actualize the grid of clients from the collection saved in the database
        /// </summary>
        private void Actualize_GridClientsFromDatabase()
        {
            if (m_Database_Handler != null)
            {
                //Initialize rows and columns
                m_GridClient_Column = 0;
                m_GridClient_Row = 0;
                Grid_Clients_Details.Children.Clear();

                //Clear displayed collection
                m_DisplayedClientsCollection.Clear();

                //Getting the hosts and hostesses collection  
                try
                {
                    string res = m_Database_Handler.Get_ClientsFromDatabase();
                    if (res.Contains("OK"))
                    {
                        //Sort by last name
                        SoftwareObjects.ClientsCollection.Sort(delegate (Client x, Client y)
                        {
                            if (x.corporate_name == null && y.corporate_name == null) return 0;
                            else if (x.corporate_name == null) return -1;
                            else if (y.corporate_name == null) return 1;
                            else
                            {
                                if (m_IsSortClient_Ascending == true)
                                {
                                    return x.corporate_name.CompareTo(y.corporate_name);
                                }
                                else
                                {
                                    return y.corporate_name.CompareTo(x.corporate_name);
                                }
                            }
                        });
                        //Add to grid
                        for (int iClient = 0; iClient < SoftwareObjects.ClientsCollection.Count; ++iClient)
                        {
                            Client client = SoftwareObjects.ClientsCollection[iClient];
                            if (m_Clients_IsArchiveMode == true && client.archived == 1)
                            {
                                Add_ClientToGrid(client);
                            }
                            else if (m_Clients_IsArchiveMode == false && client.archived == 0)
                            {
                                Add_ClientToGrid(client);
                            }
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
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                    SoftwareObjects.ClientsCollection = new List<Client>();
                    return;
                }
            }
        }

        /// <summary>
        /// Functions
        /// Clients
        /// Increment on grid Client columns
        /// </summary>
        private int m_GridClient_Column = 0;
        /// <summary>
        /// Functions
        /// Clients
        /// Increment on grid Client rows
        /// </summary>
        private int m_GridClient_Row = 0;
        /// <summary>
        /// Functions
        /// Clients
        /// Add a client to the grid, create a stack panel containing all the information of the client
        /// <param name="_Client">Hostess</param>
        /// </summary>
        private void Add_ClientToGrid(Client _Client)
        {
            try
            {
                //Add to displayed collection
                m_DisplayedClientsCollection.Add(_Client);

                //Treatment of the first row
                Grid_Clients_Details.RowDefinitions[0].MinHeight = 350;

                //Create the hostess button
                Button buttonClient = new Button();
                buttonClient.Padding = new Thickness(5);
                if (m_Clients_IsArchiveMode == false)
                {
                    buttonClient.Background = m_Color_Mission;
                }
                else
                {
                    buttonClient.Background = m_Color_ArchivedMission;
                }
                buttonClient.Click += IsPressed_Clients_Button;

                //Treatment of hostess info
                TextBlock clientInfo = new TextBlock();
                Run names = new Run("\n" + _Client.corporate_name);
                names.FontSize = 15;
                names.FontWeight = FontWeights.Bold;
                clientInfo.Inlines.Add(names);
                clientInfo.Inlines.Add(new LineBreak());
                string creationDate = m_Global_Handler.DateAndTime_Handler.Treat_Date(_Client.date_creation, m_Global_Handler.Language_Handler);
                string info = _Client.address + "\n" +
                    _Client.zipcode + ", " + _Client.city + "\n" +
                    m_Global_Handler.Resources_Handler.Get_Resources("CreationDate") + " : " + creationDate + "\n";
                clientInfo.Inlines.Add(info);
                //Create stack button
                StackPanel buttonStack = new StackPanel();
                buttonStack.Orientation = Orientation.Horizontal;

                //Add label
                Label buttonLabel = new Label();
                buttonLabel.Content = clientInfo;
                buttonStack.Children.Add(buttonLabel);

                //Add button
                buttonClient.Content = buttonStack;
                buttonClient.Tag = _Client.id;

                //Create the main stack panel
                StackPanel mainStackPanel = new StackPanel();
                mainStackPanel.MinHeight = 45;
                mainStackPanel.MinWidth = 145;
                mainStackPanel.Margin = new Thickness(5);
                mainStackPanel.Tag = _Client.id;
                mainStackPanel.Children.Add(buttonClient);

                //Create the cellphone call button
                Button buttonCallCellPhone = new Button();
                buttonCallCellPhone.MinHeight = 20;
                buttonCallCellPhone.Padding = new Thickness(5);
                if (m_Clients_IsArchiveMode == false)
                {
                    buttonCallCellPhone.Background = m_Color_Mission;
                }
                else
                {
                    buttonCallCellPhone.Background = m_Color_ArchivedMission;
                }
                buttonCallCellPhone.Click += IsPressed_Clients_PhoneButton;
                buttonCallCellPhone.Tag = "CellPhone";
                //Treatment of the image in the call button
                Image imgCellPhone = new Image();
                imgCellPhone.Source = m_Global_Handler.Image_Handler.Convert_ToBitmapSource(Properties.Resources.Image_CellPhone);
                StackPanel stackCellPhonePanel = new StackPanel();
                stackCellPhonePanel.Orientation = Orientation.Horizontal;
                stackCellPhonePanel.Margin = new Thickness(5);
                TextBlock cellPhoneInfo = new TextBlock();
                string phone = _Client.phone;
                if (phone.Length == 10)
                {
                    phone = phone.Insert(8, " ");
                    phone = phone.Insert(6, " ");
                    phone = phone.Insert(4, " ");
                    phone = phone.Insert(2, " ");
                }
                string cellPhoneText = "   " + phone;
                cellPhoneInfo.Inlines.Add(cellPhoneText);
                stackCellPhonePanel.Children.Add(imgCellPhone);
                stackCellPhonePanel.Children.Add(cellPhoneInfo);
                buttonCallCellPhone.Content = stackCellPhonePanel;
                //Add to the main stack panel
                mainStackPanel.Children.Add(buttonCallCellPhone);

                //Treatment of the increments
                if (m_GridClient_Column > Grid_Clients_Details.ColumnDefinitions.Count - 1)
                {
                    m_GridClient_Column = 0;
                    m_GridClient_Row += 1;
                    RowDefinition row = new RowDefinition();
                    row.MinHeight = 350;
                    Grid_Clients_Details.RowDefinitions.Add(row);
                    Grid.SetColumn(mainStackPanel, m_GridClient_Column);
                    Grid.SetRow(mainStackPanel, m_GridClient_Row);
                    m_GridClient_Column += 1;
                }
                else
                {
                    Grid.SetColumn(mainStackPanel, m_GridClient_Column);
                    Grid.SetRow(mainStackPanel, m_GridClient_Row);
                    m_GridClient_Column += 1;
                }

                //Treatment of the Grid_Client column
                Grid_Clients_Details.Children.Add(mainStackPanel);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Clients
        /// Get a client from his id
        /// <returns>Returns the client corresponding to the id</returns>
        /// </summary>
        private Client Get_SelectedClientFromButton()
        {
            try
            {
                //Test
                if (m_Button_Client_SelectedClient == null || m_Button_Client_SelectedClient.Tag == null)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("ClientSelectedError"),
                        m_Global_Handler.Resources_Handler.Get_Resources("ClientSelectedErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Error);
                    return null;
                }

                //Get the hostess from the id
                string id = (string)m_Button_Client_SelectedClient.Tag;
                Client client = SoftwareObjects.ClientsCollection.Find(x => x.id.Equals(id));

                //Actualize datagrid
                if (client != null)
                {
                    //Fill the m_DataGrid_Client_MissionsCollection
                    m_DataGrid_Clients_MissionsCollection.Clear();
                    for (int iMission = 0; iMission < SoftwareObjects.MissionsCollection.Count; ++iMission)
                    {
                        Mission missionSel = SoftwareObjects.MissionsCollection[iMission];
                        if (missionSel.id == client.id) //TODO
                        {
                            m_Datagrid_Clients_Missions data = new m_Datagrid_Clients_Missions(missionSel.id, missionSel.city, missionSel.start_date, missionSel.description);
                            m_DataGrid_Clients_MissionsCollection.Add(data);
                        }
                    }
                    DataGrid_Clients_Missions.Items.Refresh();
                }

                //Return the client
                return client;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return null;
            }
        }

        #endregion

        #region Shifts

        /// <summary>
        /// Functions
        /// Shifts
        /// Fill the shifts collection from database
        /// </summary>
        private void Fill_ShiftsCollectionFromDatabase()
        {
            try
            {
                if (m_Database_Handler != null)
                {
                    SoftwareObjects.ShiftsCollection.Clear();
                    string res = m_Database_Handler.Get_ShiftsFromDatabase();
                    if (res.Contains("OK"))
                    {
                        return;
                    }
                    else if (res.Contains("Error"))
                    {
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("ShiftsFillCollectionErrorCaption"),
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        SoftwareObjects.ShiftsCollection = new List<Shift>();
                        return;
                    }
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
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