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
        /// Initialization
        /// Variables
        /// Global handlers for common objects
        /// </summary>
        private static Handlers m_Global_Handler = null;

        /// <summary>
        /// Initialization
        /// Variables
        /// List of controls
        /// </summary>
        private List<Control> m_ListOfFields = new List<Control>();

        /// <summary>
        /// Initialization
        /// Variables
        /// List of hosts and hostesses in string format
        /// </summary>
        private List<string> m_List_HostsAndHostesses = new List<string>();
        private List<string> m_List_IdHostsAndHostesses = new List<string>();

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
        /// Variables
        /// Mission
        /// </summary>
        public Mission m_Mission = null;

        /// <summary>
        /// Initialization
        /// Variables
        /// Shift modified or created
        /// </summary>
        public Shift m_Shift = new Shift();

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
                List<String> listShiftsId = new List<string>();
                if (m_Mission.id_list_shifts != null)
                {
                    listShiftsId = new List<string>(m_Mission.id_list_shifts.Split(';'));
                }
                m_ListOfShifts = m_Database_Handler.Get_ShiftsFromListOfId(listShiftsId);

                //Fill infos mission
                Txt_Shifts_Mission.Text = m_Mission.description;
                Txt_Shifts_Mission_City.Text = m_Mission.city;
                Txt_Shifts_Mission_Client.Text = m_Mission.client_name;
                Txt_Shifts_Mission_EndDate.Text = m_Mission.end_date;
                Txt_Shifts_Mission_StartDate.Text = m_Mission.start_date;
                if (m_Mission.start_date != "")
                {
                    Cld_Shifts_Shift_Date.SelectedDate = Convert.ToDateTime(m_Mission.start_date);
                    Cld_Shifts_Shift_Date.DisplayDate = Convert.ToDateTime(m_Mission.start_date);
                }
                //Fill combo boxes
                for (int iHour = 0; iHour <= 24; ++iHour)
                {
                    Cmb_Shifts_Shift_EndHour_Hour.Items.Add(iHour.ToString("00")); ;
                    Cmb_Shifts_Shift_StartHour_Hour.Items.Add(iHour.ToString("00")); ;
                }
                Cmb_Shifts_Shift_EndHour_Hour.SelectedIndex = 12;
                Cmb_Shifts_Shift_StartHour_Hour.SelectedIndex = 12;
                for (int iMin = 0; iMin < 4; ++iMin)
                {
                    int min = iMin * 15;
                    Cmb_Shifts_Shift_EndHour_Min.Items.Add(min.ToString("00")); ;
                    Cmb_Shifts_Shift_StartHour_Min.Items.Add(min.ToString("00")); ;
                }
                Cmb_Shifts_Shift_EndHour_Min.SelectedIndex = 0;
                Cmb_Shifts_Shift_StartHour_Min.SelectedIndex = 0;

                for (int iHostess = 0; iHostess < SoftwareObjects.HostsAndHotessesCollection.Count; ++iHostess)
                {
                    Hostess hostess = SoftwareObjects.HostsAndHotessesCollection[iHostess];
                    string name = hostess.zipcode + " \t " + hostess.firstname + " " + hostess.lastname;
                    m_List_HostsAndHostesses.Add(name);
                    m_List_IdHostsAndHostesses.Add(hostess.id);
                    Cmb_Shifts_Shift_HostOrHostess.Items.Add(name);
                }
                //Fill shifts datagrid
                m_Datagrid_Missions_ShiftsCollection.Clear();
                for (int iShift = 0; iShift < m_ListOfShifts.Count; ++iShift)
                {
                    Shift shiftSel = m_ListOfShifts[iShift];
                    m_Datagrid_Mission_Shifts data = new m_Datagrid_Mission_Shifts(shiftSel.id, shiftSel.date,
                        SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(shiftSel.id_hostorhostess)).firstname + " " +
                        SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(shiftSel.id_hostorhostess)).lastname,
                        shiftSel.start_time, shiftSel.end_time);
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
                Btn_Shifts_Quit.Content = m_Global_Handler.Resources_Handler.Get_Resources("Quit");

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
        /// Click on the add button
        /// </summary>
        private void Btn_Shifts_Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Fill parameters
                DateTime dateSelected = DateTime.Today;
                if (Cld_Shifts_Shift_Date.SelectedDate != null)
                {
                    dateSelected = (DateTime)Cld_Shifts_Shift_Date.SelectedDate;
                }
                m_Shift.date = dateSelected.ToString("dd/MM/yyyy");
                m_Shift.end_time = Cmb_Shifts_Shift_EndHour_Hour.Text + ":" + Cmb_Shifts_Shift_EndHour_Min.Text;
                m_Shift.start_time = Cmb_Shifts_Shift_StartHour_Hour.Text + ":" + Cmb_Shifts_Shift_StartHour_Min.Text;
                m_Shift.hourly_rate = Txt_Shifts_Shift_HourlyRate.Text;
                m_Shift.pause = Txt_Shifts_Shift_Pause.Text;
                if (Cmb_Shifts_Shift_HostOrHostess.SelectedIndex != -1)
                {
                    m_Shift.id_hostorhostess = m_List_IdHostsAndHostesses[Cmb_Shifts_Shift_HostOrHostess.SelectedIndex];
                }
                else
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("NoHostSelectedErrorText"),
                        m_Global_Handler.Resources_Handler.Get_Resources("NoHostSelectedErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                m_Shift.id = m_Mission.id + "_" + m_Shift.date + "_" + m_Shift.start_time + "_" + m_Shift.end_time + "_" + m_Shift.id_hostorhostess;
                m_Shift.id_mission = m_Mission.id;
                m_Shift.suit = (bool)Chk_Shifts_Shift_Suit.IsChecked;
                if (m_Mission.id_list_shifts == "")
                {
                    m_Mission.id_list_shifts = m_Shift.id;
                }
                else
                {
                    m_Mission.id_list_shifts += ";" + m_Shift.id;
                }

                //Add to internet database
                string res = m_Database_Handler.Add_ShiftToDatabase(m_Shift.date, m_Shift.end_time, m_Shift.hourly_rate, m_Shift.id, m_Shift.id_hostorhostess, m_Shift.id_mission,
                    m_Shift.pause, m_Shift.start_time, m_Shift.suit);

                //Treat the result
                if (res.Contains("OK"))
                {
                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Shift " + m_Shift.id + " added");

                    //Add to collection
                    SoftwareObjects.ShiftsCollection.Add(m_Shift);

                    //Add to datagrid
                    string hostOrHostess = Cmb_Shifts_Shift_HostOrHostess.Text.Split('\t')[1];
                    m_Datagrid_Missions_ShiftsCollection.Add(new m_Datagrid_Mission_Shifts(m_Shift.id, m_Shift.date, hostOrHostess, m_Shift.start_time, m_Shift.end_time));
                    Datagrid_Shifts.Items.Refresh();
                    Datagrid_Shifts.SelectedItem = Datagrid_Shifts.Items[Datagrid_Shifts.Items.Count - 1];

                    //Edit the mission to include the id of the new shift
                    m_Database_Handler.Edit_MissionToDatabase(m_Mission.address, m_Mission.city,
                        m_Mission.client_name, m_Mission.country, m_Mission.description, m_Mission.end_date, m_Mission.id,
                        m_Mission.id_list_shifts, m_Mission.start_date, m_Mission.state, m_Mission.zipcode);
                }
                else if (res.Contains("Error"))
                {
                    //Treatment of the error
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                    m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, res);
                    return;
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
        /// Click on the delete button
        /// </summary>
        private void Btn_Shifts_Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Verification
                if (m_Shift.id == null)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("NoShiftSelectedErrorText"),
                        m_Global_Handler.Resources_Handler.Get_Resources("NoShiftSelectedErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                //Confirm the delete
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("ShiftConfirmDelete"),
                    m_Global_Handler.Resources_Handler.Get_Resources("ShiftConfirmDeleteCaption"),
                    MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                //Delete in database
                string res = m_Database_Handler.Delete_ShiftFromDatabase(m_Shift.id);

                //Treat the result
                if (res.Contains("OK"))
                {
                    //Action
                    m_Global_Handler.Log_Handler.WriteAction("Shift " + m_Shift.id + " deleted");

                    //Delete from the datagrid
                    int index = Datagrid_Shifts.SelectedIndex;
                    m_Datagrid_Missions_ShiftsCollection.Remove((m_Datagrid_Mission_Shifts)Datagrid_Shifts.SelectedItem);
                    Datagrid_Shifts.Items.Refresh();
                    if (index - 1 >= 0)
                    {
                        //Select the one before the deleted item
                        Datagrid_Shifts.SelectedIndex = index - 1;
                    }
                    else if (Datagrid_Shifts.Items.Count > 0)
                    {
                        //Select the first one
                        Datagrid_Shifts.SelectedItem = 0;
                    }

                    //Delete from the mission
                    m_Mission.id_list_shifts = m_Mission.id_list_shifts.Replace(m_Shift.id + ";", "");
                    m_Database_Handler.Edit_MissionToDatabase(m_Mission.address, m_Mission.city,
                        m_Mission.client_name, m_Mission.country, m_Mission.description, m_Mission.end_date, m_Mission.id,
                        m_Mission.id_list_shifts, m_Mission.start_date, m_Mission.state, m_Mission.zipcode);

                    //Delete from the collection
                    SoftwareObjects.ShiftsCollection.Remove(m_Shift);

                    //Clear boxes
                    Cld_Shifts_Shift_Date.SelectedDate = DateTime.Today;
                    Cld_Shifts_Shift_Date.DisplayDate = DateTime.Today;
                    Txt_Shifts_Shift_HourlyRate.Text = "";
                    Txt_Shifts_Shift_Pause.Text = "";
                    Cmb_Shifts_Shift_EndHour_Hour.Text = "";
                    Cmb_Shifts_Shift_EndHour_Min.Text = "";
                    Cmb_Shifts_Shift_HostOrHostess.Text = "";
                    Cmb_Shifts_Shift_StartHour_Hour.Text = "";
                    Cmb_Shifts_Shift_StartHour_Min.Text = "";
                    Chk_Shifts_Shift_Suit.IsChecked = false;

                    return;
                }
                else if (res.Contains("error"))
                {
                    //Treatment of the error
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                    m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, res);
                    return;
                }
                else
                {
                    //Error connecting to web site
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                    m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, res);
                    return;
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
        /// Click on the modify button
        /// </summary>
        private void Btn_Shifts_Modify_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Fill parameters
                DateTime dateSelected = DateTime.Today;
                if (Cld_Shifts_Shift_Date.SelectedDate != null)
                {
                    dateSelected = (DateTime)Cld_Shifts_Shift_Date.SelectedDate;
                }
                m_Shift.date = dateSelected.ToString("dd/MM/yyyy");
                m_Shift.end_time = Cmb_Shifts_Shift_EndHour_Hour.Text + ":" + Cmb_Shifts_Shift_EndHour_Min.Text;
                m_Shift.start_time = Cmb_Shifts_Shift_StartHour_Hour.Text + ":" + Cmb_Shifts_Shift_StartHour_Min.Text;
                m_Shift.hourly_rate = Txt_Shifts_Shift_HourlyRate.Text;
                if (Cmb_Shifts_Shift_HostOrHostess.SelectedIndex != -1)
                {
                    m_Shift.id_hostorhostess = m_List_IdHostsAndHostesses[Cmb_Shifts_Shift_HostOrHostess.SelectedIndex];
                }
                else
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("NoHostSelectedErrorText"),
                        m_Global_Handler.Resources_Handler.Get_Resources("NoHostSelectedErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                m_Shift.id_mission = m_Mission.id;
                m_Shift.pause = Txt_Shifts_Shift_Pause.Text;
                m_Shift.suit = (bool)Chk_Shifts_Shift_Suit.IsChecked;

                //Edit to internet database
                string res = m_Database_Handler.Edit_ShiftToDatabase(m_Shift.date, m_Shift.end_time, m_Shift.hourly_rate, m_Shift.id, m_Shift.id_hostorhostess, m_Shift.id_mission,
                    m_Shift.pause, m_Shift.start_time, m_Shift.suit);

                //Treat the result
                if (res.Contains("OK"))
                {
                    //Edit id
                    string newId = m_Mission.id + "_" + m_Shift.date + "_" + m_Shift.start_time + "_" + m_Shift.end_time + "_" + m_Shift.id_hostorhostess;
                    m_Database_Handler.Edit_ShiftIdToDatabase(m_Shift.id, newId);
                    //Edit the mission to include the id of the new shift
                    m_Mission.id_list_shifts = m_Mission.id_list_shifts.Replace(m_Shift.id, newId);
                    m_Database_Handler.Edit_MissionToDatabase(m_Mission.address, m_Mission.city,
                        m_Mission.client_name, m_Mission.country, m_Mission.description, m_Mission.end_date, m_Mission.id,
                        m_Mission.id_list_shifts, m_Mission.start_date, m_Mission.state, m_Mission.zipcode);

                    //Edit into the collection
                    Shift shift = SoftwareObjects.ShiftsCollection.Find(x => x.id.Equals(m_Shift.id));
                    shift.date = m_Shift.date;
                    shift.end_time = m_Shift.end_time;
                    shift.hourly_rate = m_Shift.hourly_rate;
                    shift.id = newId;
                    shift.id_hostorhostess = m_Shift.id_hostorhostess;
                    shift.id_mission = m_Shift.id_mission;
                    shift.pause = m_Shift.pause;
                    shift.start_time = m_Shift.start_time;
                    shift.suit = m_Shift.suit;

                    //Modify selected shift id
                    m_Shift.id = newId;

                    //Edit into the datagrid
                    m_Datagrid_Mission_Shifts datagridShift = (m_Datagrid_Mission_Shifts)Datagrid_Shifts.SelectedItem;
                    if (datagridShift != null)
                    {
                        datagridShift.id = m_Shift.id;
                        datagridShift.date = m_Shift.date;
                        datagridShift.end_time = m_Shift.end_time;
                        datagridShift.hostorhostess = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(m_Shift.id_hostorhostess)).firstname + " " +
                        SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(m_Shift.id_hostorhostess)).lastname;
                        datagridShift.start_time = m_Shift.start_time;
                    }
                    Datagrid_Shifts.Items.Refresh();

                    //Edit the mission to include the id of the new shift
                    m_Database_Handler.Edit_MissionToDatabase(m_Mission.address, m_Mission.city,
                        m_Mission.client_name, m_Mission.country, m_Mission.description, m_Mission.end_date, m_Mission.id,
                        m_Mission.id_list_shifts, m_Mission.start_date, m_Mission.state, m_Mission.zipcode);
                }
                else if (res.Contains("Error"))
                {
                    //Treatment of the error
                    MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                    m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, res);
                    return;
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
        /// Click on the save button
        /// </summary>
        private void Btn_Shifts_Quit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                m_ConfirmQuit = true;
                this.DialogResult = true;
                Close();
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Datagrid shifts autogenerating columns
        /// </summary>
        private void Datagrid_Shifts_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            string columnHeader = e.Column.Header.ToString();
            if (columnHeader == "id")
            {
                e.Column.Visibility = Visibility.Collapsed;
            }
            else if (columnHeader == "date")
            {
                e.Column.Header = m_Global_Handler.Resources_Handler.Get_Resources("Date");
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            else if (columnHeader == "hostorhostess")
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
        /// Datagrid shifts selcetion changed
        /// </summary>
        private void Datagrid_Shifts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                m_Datagrid_Mission_Shifts selectedShift = (m_Datagrid_Mission_Shifts)Datagrid_Shifts.SelectedItem;
                if (selectedShift != null)
                {
                    //Find the shift in the collection
                    Shift shift = SoftwareObjects.ShiftsCollection.Find(x => x.id.Equals(selectedShift.id));
                    //Fill boxes
                    Cmb_Shifts_Shift_StartHour_Hour.Text = shift.start_time.Split(':')[0];
                    Cmb_Shifts_Shift_StartHour_Min.Text = shift.start_time.Split(':')[1];
                    Cmb_Shifts_Shift_EndHour_Hour.Text = shift.end_time.Split(':')[0];
                    Cmb_Shifts_Shift_EndHour_Min.Text = shift.end_time.Split(':')[1];
                    Txt_Shifts_Shift_HourlyRate.Text = shift.hourly_rate;
                    Txt_Shifts_Shift_Pause.Text = shift.pause;
                    Hostess hostess = SoftwareObjects.HostsAndHotessesCollection.Find(y => y.id.Equals(shift.id_hostorhostess));
                    Cmb_Shifts_Shift_HostOrHostess.Text = hostess.zipcode + " \t " + hostess.firstname + " " + hostess.lastname;
                    Chk_Shifts_Shift_Suit.IsChecked = shift.suit;
                    Cld_Shifts_Shift_Date.SelectedDate = Convert.ToDateTime(shift.date);
                    Cld_Shifts_Shift_Date.DisplayDate = Convert.ToDateTime(shift.date);
                    m_Shift = shift;
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