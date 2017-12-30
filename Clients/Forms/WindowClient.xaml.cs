using GeneralClasses;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Xml;

namespace WindowClient
{
    /// <summary>
    /// Class MainWindow_Clients from the namespace WindowHostOrHostess
    /// </summary>
    public partial class MainWindow_Client : Window
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
        public Client m_Client = null;

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

        #endregion

        #region Functions

        /// <summary>
        /// Initialization
        /// Functions
        /// Constructor for the host and hostess main window
        /// </summary>
        public MainWindow_Client(Handlers _Global_Handler, Database.Database _Database_Handler, XmlDocument _Cities_DocFrance,
            bool _IsModification, Client _Client)
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

                //Define content
                if (m_Global_Handler != null)
                {
                    Define_Content();
                }

                //Select case
                if (m_IsModification == true && _Client != null)
                {
                    m_Client = _Client;

                    //Fill the fields
                    Txt_Client_Address.Text = m_Client.address;
                    Cmb_Client_City.Text = m_Client.city;
                    Txt_Client_CorporateName.Text = m_Client.corporate_name;
                    Txt_Client_CorporateNumber.Text = m_Client.corporate_number;
                    Cmb_Client_Country.Text = m_Client.country;
                    Txt_Client_Email.Text = m_Client.email;
                    Txt_Client_Phone.Text = m_Client.phone;
                    Txt_Client_State.Text = m_Client.state;
                    Txt_Client_ZipCode.Text = m_Client.zipcode;
                    if (Cmb_Client_Country.SelectedItem.ToString() == "United States")
                    {
                        Lbl_Client_State.Visibility = Visibility.Visible;
                        Txt_Client_State.Visibility = Visibility.Visible;
                    }
                }
                else
                {
                    m_Client = new Client();
                }

                //Select the text box
                Txt_Client_CorporateName.Focus();

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
                Lbl_Client_Address.Content = m_Global_Handler.Resources_Handler.Get_Resources("Address");
                Lbl_Client_City.Content = m_Global_Handler.Resources_Handler.Get_Resources("City");
                Lbl_Client_CorporateName.Content = m_Global_Handler.Resources_Handler.Get_Resources("CorporateName") + " *";
                Lbl_Client_CorporateNumber.Content = m_Global_Handler.Resources_Handler.Get_Resources("CorporateNumber");
                Lbl_Client_Country.Content = m_Global_Handler.Resources_Handler.Get_Resources("Country");
                Lbl_Client_Email.Content = m_Global_Handler.Resources_Handler.Get_Resources("Email") + " *";
                Lbl_Client_Phone.Content = m_Global_Handler.Resources_Handler.Get_Resources("Phone");
                Lbl_Client_State.Content = m_Global_Handler.Resources_Handler.Get_Resources("State");
                Lbl_Client_ZipCode.Content = m_Global_Handler.Resources_Handler.Get_Resources("ZipCode");

                //Languages specificities
                Cmb_Client_Country.ItemsSource = Get_CountryList();
                if (m_Global_Handler.Language_Handler == "en-US")
                {
                    Lbl_Client_State.Visibility = Visibility.Visible;
                    Txt_Client_State.Visibility = Visibility.Visible;
                    Cmb_Client_Country.SelectedItem = "United States";
                }
                else
                {
                    Lbl_Client_State.Visibility = Visibility.Hidden;
                    Txt_Client_State.Visibility = Visibility.Hidden;
                    if (m_Global_Handler.Language_Handler == "fr-FR")
                    {
                        Cmb_Client_Country.SelectedItem = "France";
                    }
                }

                //Buttons
                if (m_IsModification == false)
                {
                    Btn_Client_Add.Content = m_Global_Handler.Resources_Handler.Get_Resources("Create");
                    this.Title = m_Global_Handler.Resources_Handler.Get_Resources("CreateClient");
                }
                else
                {
                    Btn_Client_Add.Content = m_Global_Handler.Resources_Handler.Get_Resources("Edit");
                    this.Title = m_Global_Handler.Resources_Handler.Get_Resources("EditClient");
                }
                Btn_Client_QuitWithoutSaving.Content = m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSaving");

                //List of controls
                m_ListOfFields.Add(Txt_Client_Address);
                m_ListOfFields.Add(Cmb_Client_City);
                m_ListOfFields.Add(Txt_Client_CorporateName);
                m_ListOfFields.Add(Txt_Client_CorporateNumber);
                m_ListOfFields.Add(Cmb_Client_Country);
                m_ListOfFields.Add(Txt_Client_Email);
                m_ListOfFields.Add(Txt_Client_Phone);
                if (m_Global_Handler.Language_Handler == "en-US")
                {
                    m_ListOfFields.Add(Txt_Client_State);
                }
                m_ListOfFields.Add(Txt_Client_ZipCode);
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
        private void Btn_Client_Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Verify the fields
                List<string> neededFieldsToVerify = new List<string>();
                neededFieldsToVerify.Add(m_Global_Handler.Resources_Handler.Get_Resources("CorporateName"));
                MessageBoxResult result = m_Global_Handler.Controls_Handler.Verify_BlankFields(m_ListOfFields, neededFieldsToVerify, m_Global_Handler.Resources_Handler);
                if (result == MessageBoxResult.OK || result == MessageBoxResult.Cancel)
                {
                    return;
                }
                Regex pattern = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
                if (pattern.IsMatch(Txt_Client_Email.Text) == false)
                {
                    MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("InvalidEmail"), m_Global_Handler.Resources_Handler.Get_Resources("InvalidEmailCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                //Fill parameters
                m_Client.address = Txt_Client_Address.Text;
                m_Client.city = Cmb_Client_City.Text;
                m_Client.country = Cmb_Client_Country.Text;
                m_Client.corporate_name = Txt_Client_CorporateName.Text;
                m_Client.corporate_number = Txt_Client_CorporateNumber.Text;
                m_Client.date_creation = DateTime.Now.ToString();
                m_Client.email = Txt_Client_Email.Text;
                m_Client.phone = Txt_Client_Phone.Text;
                m_Client.state = Txt_Client_State.Text;
                m_Client.zipcode = Txt_Client_ZipCode.Text;

                //Creation
                if (m_IsModification == false)
                {
                    //Creation of the id
                    m_Client.id = Create_ClientId(m_Client);

                    //Add to internet database
                    string res = m_Database_Handler.Add_ClientToDatabase(m_Client.address, m_Client.city, m_Client.corporate_name,
                        m_Client.corporate_number, m_Client.country, m_Client.email, m_Client.id, m_Client.phone, m_Client.state, m_Client.zipcode);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Add to collection
                        SoftwareObjects.ClientsCollection.Add(m_Client);

                        //Close the window
                        m_ConfirmQuit = true;
                        this.DialogResult = true;
                        Close();
                    }
                    else if (res.Contains("Error"))
                    {
                        //Treatment of the error
                        MessageBox.Show(this, res, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                        m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, res);
                        return;
                    }
                }
                //Modification
                else
                {
                    //Edit in internet database
                    string res = m_Database_Handler.Edit_ClientToDatabase(m_Client.address, m_Client.city, m_Client.corporate_name,
                        m_Client.corporate_number, m_Client.country, m_Client.email, m_Client.id, m_Client.phone, m_Client.state, m_Client.zipcode);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Close
                        m_ConfirmQuit = true;
                        this.DialogResult = true;

                        Close();
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
        private void Btn_Client_QuitWithoutSaving_Click(object sender, RoutedEventArgs e)
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
        /// Change in the country combo box
        /// </summary>
        private void Cmb_Client_Country_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (m_IsLoaded == false)
            {
                return;
            }

            if (Cmb_Client_Country.SelectedItem.ToString() == "United States")
            {
                Lbl_Client_State.Visibility = Visibility.Visible;
                Txt_Client_State.Visibility = Visibility.Visible;
            }
            else
            {
                Lbl_Client_State.Visibility = Visibility.Hidden;
                Txt_Client_State.Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        /// Event
        /// Text changed in the zip code text box
        /// </summary>
        private void Txt_Client_ZipCode_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                //French cities
                if (Cmb_Client_Country.Text == "France")
                {
                    //Verification
                    if (Txt_Client_ZipCode.Text.Length < 5)
                    {
                        Cmb_Client_City.Items.Clear();
                        return;
                    }
                    else if (Txt_Client_ZipCode.Text.Length == 5)
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
                                        if (zipCode.Contains(Txt_Client_ZipCode.Text))
                                        {
                                            string city = childNode.ChildNodes[4].InnerText;
                                            Cmb_Client_City.Items.Add(city);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Cmb_Client_City.Items.Clear();
                        return;
                    }
                }
                else
                {
                    //No treatment for the moment - TODO
                    Cmb_Client_City.Items.Clear();
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
        /// Creation of the Id of the client
        /// </summary>
        private string Create_ClientId(Client _Client)
        {
            string id = "";

            id += SoftwareObjects.ClientsCollection.Count;
            id += "_";
            if (_Client.corporate_name.Length > 2)
            {
                id += _Client.corporate_name.Substring(0, 2);
            }
            if (_Client.corporate_number.Length > 2)
            {
                id += "_" + _Client.corporate_number.Substring(0, 2);
            }
            if (_Client.zipcode != "")
            {
                id += "_" + _Client.zipcode;
            }

            return id;
        }

        #endregion

    }
}