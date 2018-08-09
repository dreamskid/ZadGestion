using GeneralClasses;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Xml;

namespace WindowMission
{
    /// <summary>
    /// Class MainWindow_Mission from the namespace WindowMission
    /// </summary>
    public partial class MainWindow_Mission : Window
    {
        #region Initialization

        #region Variables

        /// <summary>
        /// Initialization
        /// Variables
        /// Xml document containing all the cities in France with zipcode
        /// </summary>
        private XmlDocument m_Cities_DocFrance = new XmlDocument();

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
        /// Confirmation of window loaded
        /// </summary>
        private bool m_IsLoaded = false;

        /// <summary>
        /// Initialization
        /// Variables
        /// Boolean indiciating if it is a modification of contact (true) or a creation (false)
        /// </summary>
        private bool m_IsModification = false;

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

        #endregion

        #region Functions

        /// <summary>
        /// Initialization
        /// Functions
        /// Constructor for the mission main window
        /// </summary>
        public MainWindow_Mission(Handlers _Global_Handler, Database.Database _Database_Handler, XmlDocument _Cities_DocFrance,
            bool _IsModification, Mission _Mission)
        {
            try
            {
                //Intialize the components
                InitializeComponent();
                this.Closing += new CancelEventHandler(Window_Closing);

                //Initialize variables
                m_Global_Handler = _Global_Handler;
                m_Database_Handler = _Database_Handler;
                m_Cities_DocFrance = _Cities_DocFrance;
                m_IsModification = _IsModification;

                //Fill clients combo box
                for (int iClient = 0; iClient < SoftwareObjects.ClientsCollection.Count; ++iClient)
                {
                    Client client = SoftwareObjects.ClientsCollection[iClient];
                    if (client.archived == 0)
                    {
                        Cmb_Mission_Client.Items.Add(client.corporate_name);
                    }
                }

                //Define content
                if (m_Global_Handler != null)
                {
                    Define_Content();
                }

                //Select case
                if (m_IsModification == true && _Mission != null)
                {
                    m_Mission = _Mission;

                    //Fill the fields
                    Txt_Mission_Address.Text = m_Mission.address;
                    Txt_Mission_Description.Text = m_Mission.description;
                    Txt_Mission_Zipcode.Text = m_Mission.zipcode;

                    Cmb_Mission_City.Text = m_Mission.city;
                    Cmb_Mission_Client.Text = m_Mission.client_name;
                    Cmb_Mission_Country.Text = m_Mission.country;
                    if (Cmb_Mission_Country.Text == "United States")
                    {
                        Lbl_Mission_State.Visibility = Visibility.Visible;
                        Txt_Mission_State.Visibility = Visibility.Visible;
                    }

                    DateTime startDate = Convert.ToDateTime(m_Mission.start_date);
                    Cld_Mission_StartDate.SelectedDate = startDate;
                    Cld_Mission_StartDate.DisplayDate = startDate;
                    DateTime endDate = Convert.ToDateTime(m_Mission.end_date);
                    Cld_Mission_EndDate.SelectedDate = endDate;
                    Cld_Mission_EndDate.DisplayDate = endDate;

                    //Fill shifts
                    String[] listShifts = m_Mission.id_list_shifts.Split(';');
                    //TODO - Get shifts
                }
                else
                {
                    m_Mission = new Mission();
                }

                //Select the text box
                Cmb_Mission_Client.Focus();

                //Window loaded
                m_IsLoaded = true;
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
                //Labels
                Lbl_Mission_Address.Content = m_Global_Handler.Resources_Handler.Get_Resources("Address");
                Lbl_Mission_City.Content = m_Global_Handler.Resources_Handler.Get_Resources("City");
                Lbl_Mission_Client.Content = m_Global_Handler.Resources_Handler.Get_Resources("Customer") + "*";
                Lbl_Mission_Country.Content = m_Global_Handler.Resources_Handler.Get_Resources("Country");
                Lbl_Mission_Description.Content = m_Global_Handler.Resources_Handler.Get_Resources("Description") + "*";
                Lbl_Mission_EndDate.Content = m_Global_Handler.Resources_Handler.Get_Resources("EndDate");
                Lbl_Mission_Shifts.Content = m_Global_Handler.Resources_Handler.Get_Resources("Shifts");
                Lbl_Mission_StartDate.Content = m_Global_Handler.Resources_Handler.Get_Resources("StartDate");
                Lbl_Mission_State.Content = m_Global_Handler.Resources_Handler.Get_Resources("State");
                Lbl_Mission_Zipcode.Content = m_Global_Handler.Resources_Handler.Get_Resources("ZipCode");

                //Combo box
                Cmb_Mission_Country.ItemsSource = Get_CountryList();
                Cmb_Mission_Country.Items.SortDescriptions.Add(new SortDescription("", ListSortDirection.Ascending));
                if (m_Global_Handler.Language_Handler == "en-US")
                {
                    Lbl_Mission_State.Visibility = Visibility.Visible;
                    Txt_Mission_State.Visibility = Visibility.Visible;
                    Cmb_Mission_Country.SelectedItem = "United States";
                }
                else
                {
                    Lbl_Mission_State.Visibility = Visibility.Hidden;
                    Txt_Mission_State.Visibility = Visibility.Hidden;
                    if (m_Global_Handler.Language_Handler == "fr-FR")
                    {
                        Cmb_Mission_Country.SelectedItem = "France";
                    }
                }

                //Buttons
                if (m_IsModification == false)
                {
                    Btn_Mission_Add.Content = m_Global_Handler.Resources_Handler.Get_Resources("Create");
                    this.Title = m_Global_Handler.Resources_Handler.Get_Resources("CreateMission");
                }
                else
                {
                    Btn_Mission_Add.Content = m_Global_Handler.Resources_Handler.Get_Resources("Modify");
                    this.Title = m_Global_Handler.Resources_Handler.Get_Resources("EditMission");
                }
                Btn_Mission_QuitWithoutSaving.Content = m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSaving");
                Btn_Mission_Shifts_Edit.Content = m_Global_Handler.Resources_Handler.Get_Resources("Edit");

                //List of controls
                m_ListOfFields.Add(Cmb_Mission_Client);
                m_ListOfFields.Add(Txt_Mission_Description);
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
        /// Click on add button
        /// </summary>
        private void Btn_Mission_Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Verify the fields
                List<string> neededFieldsToVerify = new List<string>();
                neededFieldsToVerify.Add(m_Global_Handler.Resources_Handler.Get_Resources("Customer"));
                neededFieldsToVerify.Add(m_Global_Handler.Resources_Handler.Get_Resources("Description"));
                MessageBoxResult result = m_Global_Handler.Controls_Handler.Verify_BlankFields(m_ListOfFields, neededFieldsToVerify, m_Global_Handler.Resources_Handler);
                if (result == MessageBoxResult.OK || result == MessageBoxResult.Cancel)
                {
                    return;
                }
                if (Cld_Mission_EndDate.SelectedDate == null || Cld_Mission_StartDate.SelectedDate == null)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissingSelectedDateErrorText"),
                        m_Global_Handler.Resources_Handler.Get_Resources("MissingSelectedDateErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                //Creation
                if (m_IsModification == false)
                {
                    Add_MissionToDatabase();
                }
                //Modification
                else
                {
                    Edit_MissionToDatabase();
                }

                //Close the window
                m_ConfirmQuit = true;
                this.DialogResult = true;
                Close();
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, exception.StackTrace);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Click on quit wihout saving button
        /// </summary>
        private void Btn_Mission_QuitWithoutSaving_Click(object sender, RoutedEventArgs e)
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
        /// Click on edit shifts button
        /// </summary>
        private void Btn_Mission_Shifts_Edit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Verify the fields
                List<string> neededFieldsToVerify = new List<string>();
                neededFieldsToVerify.Add(m_Global_Handler.Resources_Handler.Get_Resources("Customer"));
                neededFieldsToVerify.Add(m_Global_Handler.Resources_Handler.Get_Resources("Description"));
                MessageBoxResult result = m_Global_Handler.Controls_Handler.Verify_BlankFields(m_ListOfFields, neededFieldsToVerify, m_Global_Handler.Resources_Handler);
                if (result == MessageBoxResult.OK || result == MessageBoxResult.Cancel)
                {
                    return;
                }
                if (Cld_Mission_EndDate.SelectedDate == null || Cld_Mission_StartDate.SelectedDate == null)
                {
                    MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("MissingSelectedDateErrorText"),
                        m_Global_Handler.Resources_Handler.Get_Resources("MissingSelectedDateErrorCaption"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                //Add or edit
                if (SoftwareObjects.MissionsCollection.Find(x => x.id.Equals(m_Mission.id)) == null)
                {
                    Add_MissionToDatabase();
                }
                else
                {
                    Edit_MissionToDatabase();
                }

                //Open the mission window
                MainWindowShift shiftsWindow = new MainWindowShift(m_Global_Handler, m_Database_Handler, m_Mission);
                Nullable<bool> resShow = shiftsWindow.ShowDialog();

                //Close
                if (resShow == true)
                {
                    m_ConfirmQuit = true;
                    this.DialogResult = true;
                    Close();
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, exception.StackTrace);
                return;
            }
        }

        /// <summary>
        /// Event
        /// Change in the country combo box
        /// </summary>
        private void Cmb_Mission_Country_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (m_IsLoaded == false)
            {
                return;
            }

            if (Cmb_Mission_Country.SelectedItem.ToString() == "United States")
            {
                Lbl_Mission_State.Visibility = Visibility.Visible;
                Txt_Mission_State.Visibility = Visibility.Visible;
            }
            else
            {
                Lbl_Mission_State.Visibility = Visibility.Hidden;
                Txt_Mission_State.Visibility = Visibility.Hidden;
            }
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
        /// Text changed in the zip code text box
        /// </summary>
        private void Txt_Mission_ZipCode_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                //French cities
                if (Cmb_Mission_Country.Text == "France")
                {
                    //Verification
                    if (Txt_Mission_Zipcode.Text.Length < 5)
                    {
                        Cmb_Mission_City.Items.Clear();
                        return;
                    }
                    else if (Txt_Mission_Zipcode.Text.Length == 5)
                    {
                        //Verification
                        if (m_Cities_DocFrance.DocumentElement == null)
                        {
                            return;
                        }

                        //Research in the file
                        foreach (XmlNode node in m_Cities_DocFrance.DocumentElement.ChildNodes)
                        {
                            if (node.Name == "database" && node.Attributes["name"].Value == "villes_tests")
                            {
                                foreach (XmlNode childNode in node.ChildNodes)
                                {
                                    if (childNode.Name == "table" && childNode.Attributes["name"].Value == "villes_france")
                                    {
                                        string zipCode = childNode.ChildNodes[7].InnerText;
                                        if (zipCode.Contains(Txt_Mission_Zipcode.Text))
                                        {
                                            string city = childNode.ChildNodes[4].InnerText;
                                            Cmb_Mission_City.Items.Add(city);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Cmb_Mission_City.Items.Clear();
                        return;
                    }
                }
                else
                {
                    //No treatment for the moment - TODO
                    Cmb_Mission_City.Items.Clear();
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

        #region Functions

        /// <summary>
        /// Functions
        /// Add a mission to the database
        /// </summary>
        private void Add_MissionToDatabase()
        {
            //Fill parameters
            m_Mission.address = Txt_Mission_Address.Text;
            m_Mission.city = Cmb_Mission_City.Text; ;
            m_Mission.client_name = Cmb_Mission_Client.Text;
            m_Mission.country = Cmb_Mission_Country.Text;
            m_Mission.description = Txt_Mission_Description.Text;
            m_Mission.end_date = Cld_Mission_EndDate.SelectedDate.ToString();
            m_Mission.start_date = Cld_Mission_StartDate.SelectedDate.ToString();
            m_Mission.state = Txt_Mission_State.Text;
            m_Mission.zipcode = Txt_Mission_Zipcode.Text;

            //Creation of the id
            m_Mission.id = m_Mission.Create_MissionId();
            m_Mission.date_creation = DateTime.Today.ToString();

            //Add to internet database
            string res = m_Database_Handler.Add_MissionToDatabase(m_Mission.address, m_Mission.city,
                m_Mission.client_name, m_Mission.country, m_Mission.description, m_Mission.end_date, m_Mission.id,
                m_Mission.start_date, m_Mission.state, m_Mission.zipcode);

            //Treat the result
            if (res.Contains("OK"))
            {
                //Add to collection
                SoftwareObjects.MissionsCollection.Add(m_Mission);
            }
            else if (res.Contains("Error"))
            {
                //Treatment of the error
                MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, res);
                return;
            }
        }

        /// <summary>
        /// Functions
        /// Edit a mission to the database
        /// </summary>
        private void Edit_MissionToDatabase()
        {
            //Fill parameters
            m_Mission.address = Txt_Mission_Address.Text;
            m_Mission.city = Cmb_Mission_City.Text; ;
            m_Mission.client_name = Cmb_Mission_Client.Text;
            m_Mission.country = Cmb_Mission_Country.Text;
            m_Mission.description = Txt_Mission_Description.Text;
            m_Mission.end_date = Cld_Mission_EndDate.SelectedDate.ToString();
            m_Mission.start_date = Cld_Mission_StartDate.SelectedDate.ToString();
            m_Mission.state = Txt_Mission_State.Text;
            m_Mission.zipcode = Txt_Mission_Zipcode.Text;

            //Edit in internet database
            string res = m_Database_Handler.Edit_MissionToDatabase(m_Mission.address, m_Mission.city,
                m_Mission.client_name, m_Mission.country, m_Mission.description, m_Mission.end_date, m_Mission.id,
                m_Mission.id_list_shifts, m_Mission.start_date, m_Mission.state, m_Mission.zipcode);

            //Treat the result
            if (res.Contains("OK"))
            {
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

        #endregion

    }
}