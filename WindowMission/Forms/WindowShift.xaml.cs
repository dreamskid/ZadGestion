using GeneralClasses;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

namespace WindowMission
{
    /// <summary>
    /// Class MainWindow_Shift from the namespace WindowMission
    /// </summary>
    public partial class MainWindowShift : Window
    {
        #region Initialization

        #region Variables

        /// <summary>
        /// Initialization
        /// Variables
        /// Confirmation of quit the window
        /// </summary>
        private bool m_ConfirmQuit = false;

        /// <summary>
        /// Initialization
        /// Variables
        /// Internet handler
        /// </summary>
        private static Database.Database m_Database_Handler = null;

        /// <summary>
        /// Class for Datagrid_Shifts
        /// </summary>
        public class m_Datagrid_Shifts
        {

            /// <summary>
            /// Constructor for m_Datagrid_Shifts
            /// </summary>
            public m_Datagrid_Shifts(int _Id, string _Date, string _StartTime, string EndTime)
            {
                id = _Id;
                date = _Date;
                start_time = _StartTime;
                end_time = EndTime;
            }

            /// <summary>
            /// Id
            /// </summary>
            public int id { set; get; }
            /// <summary>
            /// Date
            /// </summary>
            public string date { set; get; }
            /// <summary>
            /// Start of the shift time
            /// </summary>
            public string start_time { set; get; }
            /// <summary>
            /// End of the shift time
            /// </summary>
            public string end_time { set; get; }
        }

        /// <summary>
        /// List of shifts
        /// </summary>
        private List<m_Datagrid_Shifts> m_Datagrid_ShiftsCollection = new List<m_Datagrid_Shifts>();

        /// <summary>
        /// Initialization
        /// Variables
        /// Global handlers for common objects
        /// </summary>
        private static Handlers m_Global_Handler = null;

        /// <summary>
        /// Initialization
        /// Variables
        /// Contact modified or created
        /// </summary>
        public Mission m_Mission = null;

        /// <summary>
        /// Initialization
        /// Variables
        /// List of controls
        /// </summary>
        private List<Control> m_ListOfFields = new List<Control>();

        /// <summary>
        /// Initialization
        /// Variables
        /// List of shifts
        /// </summary>
        private List<Shift> m_ListOfShifts = new List<Shift>();

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
            public m_Datagrid_Mission_Shifts(string _Date, string _HostOrHostess, string _StartTime, string _EndTime)
            {
                date = _Date;
                hostorhostess = _HostOrHostess;
                start_time = _StartTime;
                end_time = _EndTime;
            }

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

        #endregion

        #region Functions

        /// <summary>
        /// Initialization
        /// Functions
        /// Constructor for the mission main window
        /// </summary>
        public MainWindowShift(Handlers _Global_Handler, Database.Database _Database_Handler, Mission _Mission)
        {
            try
            {
                //Intialize the components
                InitializeComponent();
                this.Closing += new CancelEventHandler(Window_Closing);

                //Initialize variables
                m_Global_Handler = _Global_Handler;
                m_Database_Handler = _Database_Handler;
                m_Mission = _Mission;

                //Load shifts
                List<String> listShiftsId = new List<string>(m_Mission.id_list_shifts.Split(';'));
                m_ListOfShifts = m_Database_Handler.Get_ShiftsFromListOfId(listShiftsId);

                //Fill infos mission
                Txt_Shifts_Mission.Text = m_Mission.description;
                Txt_Shifts_Mission_City.Text = m_Mission.city;
                Txt_Shifts_Mission_Client.Text = m_Mission.client_name;
                Txt_Shifts_Mission_EndDate.Text = m_Mission.end_date;
                Txt_Shifts_Mission_StartDate.Text = m_Mission.start_date;

                //Fill combo boxes
                for (int iHour = 0; iHour <= 24; ++iHour)
                {
                    Cmb_Shifts_Shift_EndHour_Hour.Items.Add(iHour.ToString("00")); ;
                    Cmb_Shifts_Shift_StartHour_Hour.Items.Add(iHour.ToString("00")); ;
                }
                for (int iMin = 0; iMin <= 59; ++iMin)
                {
                    Cmb_Shifts_Shift_EndHour_Min.Items.Add(iMin.ToString("00")); ;
                    Cmb_Shifts_Shift_StartHour_Min.Items.Add(iMin.ToString("00")); ;
                }

                //Fill shifts datagrid
                m_Datagrid_Missions_ShiftsCollection.Clear();
                for (int iShift = 0; iShift < m_ListOfShifts.Count; ++iShift)
                {
                    Shift shiftSel = m_ListOfShifts[iShift];
                    m_Datagrid_Mission_Shifts data = new m_Datagrid_Mission_Shifts(shiftSel.date, SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(shiftSel.id_list_hostsandhostesses)).lastname, shiftSel.start_time, shiftSel.end_time);
                    m_Datagrid_Missions_ShiftsCollection.Add(data);
                }
                Datagrid_Shifts.Items.Refresh();

                //Define content
                Define_Content();
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Initialization
        /// Functions
        /// Define each content for labels, button, text in the form
        /// </summary>
        private void Define_Content()
        {
            try
            {
                //Title
                this.Title = m_Global_Handler.Resources_Handler.Get_Resources("Shifts");

                //Check box
                Chk_Shifts_Shift_Suit.Content = m_Global_Handler.Resources_Handler.Get_Resources("LoanSuit");

                //Labels
                Lbl_Shifts_Mission_City.Content = m_Global_Handler.Resources_Handler.Get_Resources("City");
                Lbl_Shifts_Mission_Client.Content = m_Global_Handler.Resources_Handler.Get_Resources("Customer");
                Lbl_Shifts_Mission_Description.Content = m_Global_Handler.Resources_Handler.Get_Resources("Description");
                Lbl_Shifts_Mission_EndDate.Content = m_Global_Handler.Resources_Handler.Get_Resources("EndDate");
                Lbl_Shifts_Mission_StartDate.Content = m_Global_Handler.Resources_Handler.Get_Resources("StartDate");
                Lbl_Shifts_Shifts.Content = m_Global_Handler.Resources_Handler.Get_Resources("Shifts");
                Lbl_Shifts_Shift_Date.Content = m_Global_Handler.Resources_Handler.Get_Resources("Date");
                Lbl_Shifts_Shift_EndHour.Content = m_Global_Handler.Resources_Handler.Get_Resources("EndHour");
                Lbl_Shifts_Shift_HostOrHostess.Content = m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostess");
                Lbl_Shifts_Shift_HourlyRate.Content = m_Global_Handler.Resources_Handler.Get_Resources("HourlyRate");
                Lbl_Shifts_Shift_Pause.Content = m_Global_Handler.Resources_Handler.Get_Resources("PauseDuration");
                Lbl_Shifts_Shift_StartHour.Content = m_Global_Handler.Resources_Handler.Get_Resources("StartHour");

                //Buttons
                Btn_Shifts_Add.Content = m_Global_Handler.Resources_Handler.Get_Resources("Add");
                Btn_Shifts_Delete.Content = m_Global_Handler.Resources_Handler.Get_Resources("Delete");
                Btn_Shifts_Modify.Content = m_Global_Handler.Resources_Handler.Get_Resources("Modify");
                Btn_Shifts_Save.Content = m_Global_Handler.Resources_Handler.Get_Resources("Save");
                Btn_Shifts_QuitWithoutSaving.Content = m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSaving");

                //Datagrid
                Datagrid_Shifts.ItemsSource = m_Datagrid_Missions_ShiftsCollection;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Initialization
        /// Functions
        /// Get the list of countries
        /// <return>The list of countries</return>
        /// </summary>
        private List<string> Get_CountryList()
        {
            try
            {
                //Create the list
                List<string> cultureList = new List<string>();

                //Get the list
                CultureInfo[] cultures = CultureInfo.GetCultures(CultureTypes.SpecificCultures);
                foreach (CultureInfo culture in cultures)
                {
                    RegionInfo region = new RegionInfo(culture.LCID);

                    if (!(cultureList.Contains(region.NativeName)))
                    {
                        cultureList.Add(region.NativeName);
                    }
                }

                //Return the list
                return cultureList;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return new List<string>();
            }
        }

        #endregion

        #endregion

        #region Events

        /// <summary>
        /// Event
        /// Click on quit wihout saving button
        /// </summary>
        private void Btn_Shifts_QuitWithoutSaving_Click(object sender, RoutedEventArgs e)
        {
            //Confirm
            MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSavingMessage"),
                            m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSavingCaption"),
                            MessageBoxButton.YesNo, MessageBoxImage.Exclamation);

            //Quit
            if (result == MessageBoxResult.Yes)
            {
                m_ConfirmQuit = true;
                this.DialogResult = false;
                Close();
            }
        }

        /// <summary>
        /// Event
        /// Click on the save button
        /// </summary>
        private void Btn_Shifts_Save_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// Event
        /// Datagrid shifts autogenerating columns
        /// </summary>
        private void Datagrid_Shifts_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            //TODO
            string columnHeader = e.Column.Header.ToString();
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
        /// Closing window
        /// </summary>
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            //Already quit without saving
            if (m_ConfirmQuit == true)
            {
                return;
            }

            //Confirm
            MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSavingMessage"),
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