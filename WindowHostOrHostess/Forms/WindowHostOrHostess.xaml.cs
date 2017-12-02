using GeneralClasses;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Xml;

namespace WindowHostOrHostess
{
    /// <summary>
    /// Class MainWindow_HostOrHostess from the namespace WindowHostOrHostess
    /// </summary>
    public partial class MainWindow_HostOrHostess : Window
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
        public Hostess m_HostOrHostess = null;

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
        /// List of photos
        /// </summary>
        public List<string> m_ListOfPhotos = new List<string>();

        #endregion

        #region Functions

        /// <summary>
        /// Initialization
        /// Functions
        /// Constructor for the host and hostess main window
        /// </summary>
        public MainWindow_HostOrHostess(Handlers _Global_Handler, Database.Database _Database_Handler, XmlDocument _Cities_DocFrance,
            bool _IsModification, Hostess _HostAndHostess)
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
                if (m_IsModification == true && _HostAndHostess != null)
                {
                    m_HostOrHostess = _HostAndHostess;

                    //Fill the fields
                    Txt_HostAndHostess_Address.Text = m_HostOrHostess.address;
                    Txt_HostAndHostess_BirthCity.Text = m_HostOrHostess.birth_city;
                    String[] birthDate = m_HostOrHostess.birth_date.Split(' ');
                    Cmb_HostAndHostess_BirthDate_Day.Text = birthDate[0];
                    Cmb_HostAndHostess_BirthDate_Month.Text = birthDate[1];
                    Cmb_HostAndHostess_BirthDate_Year.Text = birthDate[2];
                    Txt_HostAndHostess_CellPhone.Text = m_HostOrHostess.cellphone;
                    Cmb_HostAndHostess_City.Text = m_HostOrHostess.city;
                    Cmb_HostAndHostess_Country.Text = m_HostOrHostess.country;
                    Txt_HostAndHostess_Email.Text = m_HostOrHostess.email;
                    Txt_HostAndHostess_FirstName.Text = m_HostOrHostess.firstname;
                    if (m_HostOrHostess.has_car == true)
                    {
                        Cmb_HostAndHostess_Car.Text = m_Global_Handler.Resources_Handler.Get_Resources("Yes");
                    }
                    else
                    {
                        Cmb_HostAndHostess_Car.Text = m_Global_Handler.Resources_Handler.Get_Resources("No");
                    }
                    if (m_HostOrHostess.has_driver_licence == true)
                    {
                        Cmb_HostAndHostess_Licence.Text = m_Global_Handler.Resources_Handler.Get_Resources("Yes");
                    }
                    else
                    {
                        Cmb_HostAndHostess_Licence.Text = m_Global_Handler.Resources_Handler.Get_Resources("No");
                    }
                    Txt_HostAndHostess_IdPaycheck.Text = m_HostOrHostess.id_paycheck;
                    Chk_HostOrhostess_Language_English.IsChecked = m_HostOrHostess.language_english;
                    Chk_HostOrhostess_Language_German.IsChecked = m_HostOrHostess.language_german;
                    Chk_HostOrhostess_Language_Italian.IsChecked = m_HostOrHostess.language_italian;
                    Txt_HostAndHostess_Language_Others.Text = m_HostOrHostess.language_others;
                    Chk_HostOrhostess_Language_Spanish.IsChecked = m_HostOrHostess.language_spanish;
                    Txt_HostAndHostess_LastName.Text = m_HostOrHostess.lastname;
                    Chk_HostOrhostess_Profile_Event.IsChecked = m_HostOrHostess.profile_event;
                    Chk_HostOrhostess_Profile_Permanent.IsChecked = m_HostOrHostess.profile_permanent;
                    Chk_HostOrhostess_Profile_Street.IsChecked = m_HostOrHostess.profile_street;
                    Cmb_HostAndHostess_Sex.Text = m_HostOrHostess.sex;
                    Cmb_HostAndHostess_Size.Text = m_HostOrHostess.size;
                    Cmb_HostAndHostess_Sizes_Pants.Text = m_HostOrHostess.size_pants;
                    Cmb_HostAndHostess_Sizes_Shirt.Text = m_HostOrHostess.size_shirt;
                    Cmb_HostAndHostess_Sizes_Shoes.Text = m_HostOrHostess.size_shoes;
                    Txt_HostAndHostess_SocialSecurityNumber.Text = m_HostOrHostess.social_number;
                    Txt_HostAndHostess_State.Text = m_HostOrHostess.state;
                    Txt_HostAndHostess_ZipCode.Text = m_HostOrHostess.zipcode;
                    if (Cmb_HostAndHostess_Country.SelectedItem.ToString() == "United States")
                    {
                        Lbl_HostAndHostess_State.Visibility = Visibility.Visible;
                        Txt_HostAndHostess_State.Visibility = Visibility.Visible;
                    }

                    //Get repository
                    string repository = SoftwareObjects.GlobalSettings.photos_repository + "\\" + m_HostOrHostess.id;
                    if (Directory.Exists(repository))
                    {
                        //Get photos
                        string[] files = Directory.GetFiles(repository);
                        foreach (var fileName in files)
                        {
                            m_ListOfPhotos.Add(fileName);
                        }

                        //Display the first photo
                        if (m_ListOfPhotos.Count > 0)
                        {
                            string photoFilename = m_ListOfPhotos[0];
                            BitmapImage photoBmp = m_Global_Handler.Image_Handler.Load_BitmapImage(photoFilename);
                            Img_HostAndHostess_Photo.Source = photoBmp;
                            Img_HostAndHostess_Photo.Tag = 0;
                        }
                    }
                }
                else
                {
                    m_HostOrHostess = new Hostess();
                    Cmb_HostAndHostess_BirthDate_Day.Text = "1";
                    Cmb_HostAndHostess_BirthDate_Month.Text = m_Global_Handler.Resources_Handler.Get_Resources("January");
                    Cmb_HostAndHostess_BirthDate_Year.Text = "1995";
                    Cmb_HostAndHostess_Car.Text = "Yes";
                    Cmb_HostAndHostess_Licence.Text = "Yes";
                    Cmb_HostAndHostess_Size.Text = "170";
                    Cmb_HostAndHostess_Sizes_Pants.Text = "34";
                    Cmb_HostAndHostess_Sizes_Shirt.Text = "S";
                    Cmb_HostAndHostess_Sizes_Shoes.Text = "37";
                }

                //Select the text box
                Txt_HostAndHostess_FirstName.Focus();

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
                //Check boxes
                Chk_HostOrhostess_Language_English.Content = m_Global_Handler.Resources_Handler.Get_Resources("English");
                Chk_HostOrhostess_Language_German.Content = m_Global_Handler.Resources_Handler.Get_Resources("German");
                Chk_HostOrhostess_Language_Italian.Content = m_Global_Handler.Resources_Handler.Get_Resources("Italian");
                Chk_HostOrhostess_Language_Spanish.Content = m_Global_Handler.Resources_Handler.Get_Resources("Spanish");
                Chk_HostOrhostess_Profile_Event.Content = m_Global_Handler.Resources_Handler.Get_Resources("Event");
                Chk_HostOrhostess_Profile_Permanent.Content = m_Global_Handler.Resources_Handler.Get_Resources("Permanent");
                Chk_HostOrhostess_Profile_Street.Content = m_Global_Handler.Resources_Handler.Get_Resources("Street");

                //Labels
                Lbl_HostAndHostess_Address.Content = m_Global_Handler.Resources_Handler.Get_Resources("Address");
                Lbl_HostAndHostess_BirthCity.Content = m_Global_Handler.Resources_Handler.Get_Resources("BirthCity");
                Lbl_HostAndHostess_BirthDate.Content = m_Global_Handler.Resources_Handler.Get_Resources("BirthDate");
                Lbl_HostAndHostess_Car.Content = m_Global_Handler.Resources_Handler.Get_Resources("Car");
                Lbl_HostAndHostess_CellPhone.Content = m_Global_Handler.Resources_Handler.Get_Resources("CellPhone");
                Lbl_HostAndHostess_City.Content = m_Global_Handler.Resources_Handler.Get_Resources("City");
                Lbl_HostAndHostess_Country.Content = m_Global_Handler.Resources_Handler.Get_Resources("Country");
                Lbl_HostAndHostess_Email.Content = m_Global_Handler.Resources_Handler.Get_Resources("Email") + " *"; ;
                Lbl_HostAndHostess_FirstName.Content = m_Global_Handler.Resources_Handler.Get_Resources("FirstName");
                Lbl_HostAndHostess_IdPaycheck.Content = m_Global_Handler.Resources_Handler.Get_Resources("IdPaycheck");
                Lbl_HostAndHostess_Languages.Content = m_Global_Handler.Resources_Handler.Get_Resources("Languages");
                Lbl_HostAndHostess_Languages_Others.Content = m_Global_Handler.Resources_Handler.Get_Resources("Others");
                Lbl_HostAndHostess_LastName.Content = m_Global_Handler.Resources_Handler.Get_Resources("LastName") + " *";
                Lbl_HostAndHostess_Licence.Content = m_Global_Handler.Resources_Handler.Get_Resources("DriverLicence");
                Lbl_HostAndHostess_Profile.Content = m_Global_Handler.Resources_Handler.Get_Resources("Profile");
                Lbl_HostAndHostess_Sex.Content = m_Global_Handler.Resources_Handler.Get_Resources("Sex");
                Lbl_HostAndHostess_Size.Content = m_Global_Handler.Resources_Handler.Get_Resources("Size");
                Lbl_HostAndHostess_Sizes_Pants.Content = m_Global_Handler.Resources_Handler.Get_Resources("SizePants");
                Lbl_HostAndHostess_Sizes_Shirt.Content = m_Global_Handler.Resources_Handler.Get_Resources("SizeShirt");
                Lbl_HostAndHostess_Sizes_Shoes.Content = m_Global_Handler.Resources_Handler.Get_Resources("SizeShoes");
                Lbl_HostAndHostess_SocialSecurityNumber.Content = m_Global_Handler.Resources_Handler.Get_Resources("SocialSecurityNumber");
                Lbl_HostAndHostess_State.Content = m_Global_Handler.Resources_Handler.Get_Resources("State");
                Lbl_HostAndHostess_ZipCode.Content = m_Global_Handler.Resources_Handler.Get_Resources("ZipCode");

                //Combo boxes
                for (int iDay = 1; iDay <= 31; ++iDay)
                {
                    Cmb_HostAndHostess_BirthDate_Day.Items.Add(iDay);
                }
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("January"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("February"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("March"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("April"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("May"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("June"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("July"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("August"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("September"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("October"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("November"));
                Cmb_HostAndHostess_BirthDate_Month.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("December"));
                for (int iYear = 2017; iYear >= 1900; --iYear)
                {
                    Cmb_HostAndHostess_BirthDate_Year.Items.Add(iYear);
                }
                Cmb_HostAndHostess_Car.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("Yes"));
                Cmb_HostAndHostess_Car.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("No"));
                Cmb_HostAndHostess_Country.ItemsSource = Get_CountryList();
                Cmb_HostAndHostess_Country.Items.SortDescriptions.Add(new SortDescription("", ListSortDirection.Ascending));
                Cmb_HostAndHostess_Licence.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("Yes"));
                Cmb_HostAndHostess_Licence.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("No"));
                Cmb_HostAndHostess_Sex.Items.Clear();
                Cmb_HostAndHostess_Sex.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("Female"));
                Cmb_HostAndHostess_Sex.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("Male"));
                Cmb_HostAndHostess_Sex.SelectedIndex = 0;
                for (int iSize = 150; iSize <= 200; ++iSize)
                {
                    Cmb_HostAndHostess_Size.Items.Add(iSize);
                }
                for (int iPants = 30; iPants <= 50; iPants = iPants + 2)
                {
                    Cmb_HostAndHostess_Sizes_Pants.Items.Add(iPants);
                }
                Cmb_HostAndHostess_Sizes_Shirt.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("XXS"));
                Cmb_HostAndHostess_Sizes_Shirt.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("XS"));
                Cmb_HostAndHostess_Sizes_Shirt.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("S"));
                Cmb_HostAndHostess_Sizes_Shirt.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("M"));
                Cmb_HostAndHostess_Sizes_Shirt.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("L"));
                Cmb_HostAndHostess_Sizes_Shirt.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("XL"));
                Cmb_HostAndHostess_Sizes_Shirt.Items.Add(m_Global_Handler.Resources_Handler.Get_Resources("XXL"));
                for (int iShoes = 30; iShoes <= 50; ++iShoes)
                {
                    Cmb_HostAndHostess_Sizes_Shoes.Items.Add(iShoes);
                }

                //Languages specificities
                if (m_Global_Handler.Language_Handler == "en-US")
                {
                    Lbl_HostAndHostess_State.Visibility = Visibility.Visible;
                    Txt_HostAndHostess_State.Visibility = Visibility.Visible;
                    Cmb_HostAndHostess_Country.SelectedItem = "United States";
                }
                else
                {
                    Lbl_HostAndHostess_State.Visibility = Visibility.Hidden;
                    Txt_HostAndHostess_State.Visibility = Visibility.Hidden;
                    if (m_Global_Handler.Language_Handler == "fr-FR")
                    {
                        Cmb_HostAndHostess_Country.SelectedItem = "France";
                    }
                }

                //Buttons
                TextBlock addPhoto = new TextBlock();
                addPhoto.Text = m_Global_Handler.Resources_Handler.Get_Resources("AddPhotoHostOrHostess");
                addPhoto.TextAlignment = TextAlignment.Center;
                addPhoto.TextWrapping = TextWrapping.Wrap;
                Btn_HostAndHostess_AddPhoto.Content = addPhoto;
                if (m_IsModification == false)
                {
                    Btn_HostAndHostess_Add.Content = m_Global_Handler.Resources_Handler.Get_Resources("Create");
                    this.Title = m_Global_Handler.Resources_Handler.Get_Resources("CreateHostOrHostess");
                }
                else
                {
                    Btn_HostAndHostess_Add.Content = m_Global_Handler.Resources_Handler.Get_Resources("Edit");
                    this.Title = m_Global_Handler.Resources_Handler.Get_Resources("EditHostOrHostess");
                }
                Btn_HostAndHostess_QuitWithoutSaving.Content = m_Global_Handler.Resources_Handler.Get_Resources("QuitWithoutSaving");

                //List of controls
                m_ListOfFields.Add(Txt_HostAndHostess_Address);
                m_ListOfFields.Add(Cmb_HostAndHostess_BirthDate_Day);
                m_ListOfFields.Add(Cmb_HostAndHostess_BirthDate_Month);
                m_ListOfFields.Add(Cmb_HostAndHostess_BirthDate_Year);
                m_ListOfFields.Add(Txt_HostAndHostess_CellPhone);
                m_ListOfFields.Add(Cmb_HostAndHostess_City);
                m_ListOfFields.Add(Cmb_HostAndHostess_Country);
                m_ListOfFields.Add(Txt_HostAndHostess_Email);
                m_ListOfFields.Add(Txt_HostAndHostess_FirstName);
                m_ListOfFields.Add(Txt_HostAndHostess_IdPaycheck);
                m_ListOfFields.Add(Txt_HostAndHostess_LastName);
                m_ListOfFields.Add(Cmb_HostAndHostess_Sex);
                if (m_Global_Handler.Language_Handler == "en-US")
                {
                    m_ListOfFields.Add(Txt_HostAndHostess_State);
                }
                m_ListOfFields.Add(Txt_HostAndHostess_ZipCode);
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
        private void Btn_HostAndHostess_Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Verify the fields
                List<string> neededFieldsToVerify = new List<string>();
                neededFieldsToVerify.Add(m_Global_Handler.Resources_Handler.Get_Resources("LastName"));
                neededFieldsToVerify.Add(m_Global_Handler.Resources_Handler.Get_Resources("Email"));
                MessageBoxResult result = m_Global_Handler.Controls_Handler.Verify_BlankFields(m_ListOfFields, neededFieldsToVerify, m_Global_Handler.Resources_Handler);
                if (result == MessageBoxResult.OK || result == MessageBoxResult.Cancel)
                {
                    return;
                }
                Regex pattern = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
                if (pattern.IsMatch(Txt_HostAndHostess_Email.Text) == false)
                {
                    MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("InvalidEmail"), m_Global_Handler.Resources_Handler.Get_Resources("InvalidEmailCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                //Fill parameters
                m_HostOrHostess.address = Txt_HostAndHostess_Address.Text;
                m_HostOrHostess.birth_city = Txt_HostAndHostess_BirthCity.Text;
                if (Cmb_HostAndHostess_BirthDate_Day.Text == "" || Cmb_HostAndHostess_BirthDate_Month.Text == "" || Cmb_HostAndHostess_BirthDate_Year.Text == "")
                {
                    MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("InvalidBirthDate"), m_Global_Handler.Resources_Handler.Get_Resources("InvalidBirthDateCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                m_HostOrHostess.birth_date = Cmb_HostAndHostess_BirthDate_Day.Text + " " + Cmb_HostAndHostess_BirthDate_Month.Text + " " + Cmb_HostAndHostess_BirthDate_Year.Text;
                m_HostOrHostess.cellphone = Txt_HostAndHostess_CellPhone.Text;
                m_HostOrHostess.city = Cmb_HostAndHostess_City.Text;
                m_HostOrHostess.country = Cmb_HostAndHostess_Country.Text;
                m_HostOrHostess.date_creation = DateTime.Now.ToString();
                m_HostOrHostess.email = Txt_HostAndHostess_Email.Text;
                m_HostOrHostess.firstname = Txt_HostAndHostess_FirstName.Text;
                if (Cmb_HostAndHostess_Car.Text == m_Global_Handler.Resources_Handler.Get_Resources("Yes"))
                {
                    m_HostOrHostess.has_car = true;
                }
                else
                {
                    m_HostOrHostess.has_car = false;
                }
                if (Cmb_HostAndHostess_Licence.Text == m_Global_Handler.Resources_Handler.Get_Resources("Yes"))
                {
                    m_HostOrHostess.has_driver_licence = true;
                }
                else
                {
                    m_HostOrHostess.has_driver_licence = false;
                }
                m_HostOrHostess.id_paycheck = Txt_HostAndHostess_IdPaycheck.Text;
                m_HostOrHostess.language_english = (bool)Chk_HostOrhostess_Language_English.IsChecked;
                m_HostOrHostess.language_german = (bool)Chk_HostOrhostess_Language_German.IsChecked;
                m_HostOrHostess.language_italian = (bool)Chk_HostOrhostess_Language_Italian.IsChecked;
                m_HostOrHostess.language_others = Txt_HostAndHostess_Language_Others.Text;
                m_HostOrHostess.language_spanish = (bool)Chk_HostOrhostess_Language_Spanish.IsChecked;
                m_HostOrHostess.lastname = Txt_HostAndHostess_LastName.Text;
                m_HostOrHostess.profile_event = (bool)Chk_HostOrhostess_Profile_Event.IsChecked;
                m_HostOrHostess.profile_permanent = (bool)Chk_HostOrhostess_Profile_Permanent.IsChecked;
                m_HostOrHostess.profile_street = (bool)Chk_HostOrhostess_Profile_Street.IsChecked;
                m_HostOrHostess.sex = Cmb_HostAndHostess_Sex.Text;
                m_HostOrHostess.size = Cmb_HostAndHostess_Size.Text;
                m_HostOrHostess.size_pants = Cmb_HostAndHostess_Sizes_Pants.Text;
                m_HostOrHostess.size_shirt = Cmb_HostAndHostess_Sizes_Shirt.Text;
                m_HostOrHostess.size_shoes = Cmb_HostAndHostess_Sizes_Shoes.Text;
                m_HostOrHostess.social_number = Txt_HostAndHostess_SocialSecurityNumber.Text;
                m_HostOrHostess.state = Txt_HostAndHostess_State.Text;
                m_HostOrHostess.zipcode = Txt_HostAndHostess_ZipCode.Text;

                //Creation
                if (m_IsModification == false)
                {
                    //Creation of the id
                    m_HostOrHostess.id = m_HostOrHostess.Create_HostOrHostessId();

                    //Add to internet database
                    string res = m_Database_Handler.Add_HostAndHostessToDatabase(m_HostOrHostess.address, m_HostOrHostess.birth_city,
                        m_HostOrHostess.birth_date, m_HostOrHostess.cellphone, m_HostOrHostess.city, m_HostOrHostess.country, m_HostOrHostess.email,
                        m_HostOrHostess.firstname, m_HostOrHostess.has_car, m_HostOrHostess.has_driver_licence, m_HostOrHostess.id, m_HostOrHostess.id_paycheck,
                        m_HostOrHostess.language_english, m_HostOrHostess.language_german, m_HostOrHostess.language_italian, m_HostOrHostess.language_others,
                        m_HostOrHostess.language_spanish, m_HostOrHostess.lastname, m_HostOrHostess.profile_event, m_HostOrHostess.profile_permanent,
                        m_HostOrHostess.profile_street, m_HostOrHostess.sex, m_HostOrHostess.size, m_HostOrHostess.size_pants, m_HostOrHostess.size_shirt,
                        m_HostOrHostess.size_shoes, m_HostOrHostess.social_number, m_HostOrHostess.state, m_HostOrHostess.zipcode);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Add to collection
                        SoftwareObjects.HostsAndHotessesCollection.Add(m_HostOrHostess);

                        //Save photos
                        Save_PhotosToDisc(m_HostOrHostess);

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
                    string res = m_Database_Handler.Edit_HostAndHostessToDatabase(m_HostOrHostess.address, m_HostOrHostess.birth_city,
                        m_HostOrHostess.birth_date, m_HostOrHostess.cellphone, m_HostOrHostess.city, m_HostOrHostess.country, m_HostOrHostess.email,
                        m_HostOrHostess.firstname, m_HostOrHostess.has_car, m_HostOrHostess.has_driver_licence, m_HostOrHostess.id, m_HostOrHostess.id_paycheck,
                        m_HostOrHostess.language_english, m_HostOrHostess.language_german, m_HostOrHostess.language_italian, m_HostOrHostess.language_others,
                        m_HostOrHostess.language_spanish, m_HostOrHostess.lastname, m_HostOrHostess.profile_event, m_HostOrHostess.profile_permanent,
                        m_HostOrHostess.profile_street, m_HostOrHostess.sex, m_HostOrHostess.size, m_HostOrHostess.size_pants, m_HostOrHostess.size_shirt,
                        m_HostOrHostess.size_shoes, m_HostOrHostess.social_number, m_HostOrHostess.state, m_HostOrHostess.zipcode);

                    //Treat the result
                    if (res.Contains("OK"))
                    {
                        //Save photos
                        Save_PhotosToDisc(m_HostOrHostess);

                        //Close
                        m_ConfirmQuit = true;
                        this.DialogResult = true;

                        Close();
                    }
                    else if (res.Contains("error"))
                    {
                        //Treatment of the error
                        Log ClassError = m_Database_Handler.Deserialize_JSON<Log>(res);
                        string errorText = ClassError.error;
                        MessageBox.Show(this, errorText, m_Global_Handler.Resources_Handler.Get_Resources("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                        m_Global_Handler.Log_Handler.WriteMessage(MethodBase.GetCurrentMethod().Name, errorText);
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
        /// Click on add photo button
        /// </summary>
        private void Btn_HostAndHostess_AddPhoto_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Create OpenFileDialog 
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

                // Set filter for file extension and default file extension 
                dlg.DefaultExt = ".png";
                dlg.Filter = "Image Files (*.bmp;*.jpg;*.jpeg,*.png,*.gif)|*.bmp; *.jpeg; *.jpg; *.png; *.gif";
                dlg.Multiselect = true;

                // Display OpenFileDialog by calling ShowDialog method 
                Nullable<bool> result = dlg.ShowDialog();

                // Get the selected file name and display it
                if (result == true)
                {
                    // Display photo 
                    string photoFilename = dlg.FileNames[0];
                    BitmapImage photoBmp = m_Global_Handler.Image_Handler.Load_BitmapImage(photoFilename);
                    Img_HostAndHostess_Photo.Source = photoBmp;
                    Img_HostAndHostess_Photo.Tag = 0;

                    //Save list of photos
                    for (int iPhoto = 0; iPhoto < dlg.FileNames.Length; ++iPhoto)
                    {
                        string fileName = dlg.FileNames[iPhoto];
                        if (m_ListOfPhotos.Contains(fileName) == false)
                        {
                            m_ListOfPhotos.Add(fileName);
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
        /// Click on delete photo button
        /// </summary>
        private void Btn_HostAndHostess_DeletePhoto_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Verification
                if (Img_HostAndHostess_Photo.Tag == null)
                {
                    return;
                }

                //Confirm the delete
                MessageBoxResult result = MessageBox.Show(this, m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostessPhotoConfirmDelete"),
                    m_Global_Handler.Resources_Handler.Get_Resources("HostOrHostessPhotoConfirmDeleteCaption"), MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                //Get  and display next photo
                int nDisplayed = (int)Img_HostAndHostess_Photo.Tag;

                //Delete from list
                string fileToDelete = m_ListOfPhotos[nDisplayed];
                m_ListOfPhotos.Remove(fileToDelete);

                //Verify number of photos
                if (m_ListOfPhotos.Count == 0)
                {
                    Img_HostAndHostess_Photo.Tag = null;
                    Img_HostAndHostess_Photo.Source = null;
                }
                else
                {
                    //Get  and display previous photo
                    if (nDisplayed - 1 < 0)
                    {
                        BitmapImage photoBmp = m_Global_Handler.Image_Handler.Load_BitmapImage(m_ListOfPhotos[m_ListOfPhotos.Count - 1]);
                        Img_HostAndHostess_Photo.Source = photoBmp;
                        Img_HostAndHostess_Photo.Tag = m_ListOfPhotos.Count - 1;
                    }
                    else
                    {
                        BitmapImage photoBmp = m_Global_Handler.Image_Handler.Load_BitmapImage(m_ListOfPhotos[nDisplayed - 1]);
                        Img_HostAndHostess_Photo.Source = photoBmp;
                        Img_HostAndHostess_Photo.Tag = nDisplayed - 1;
                    }
                }

                //Delete o disk
                File.Delete(fileToDelete);
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return;
            }

        }

        /// <summary>
        /// Event
        /// Click on next photo button
        /// </summary>
        private void Btn_HostAndHostess_NextPhoto_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Verification
                if (Img_HostAndHostess_Photo.Tag == null)
                {
                    return;
                }

                //Get  and display next photo
                int nDisplayed = (int)Img_HostAndHostess_Photo.Tag;
                if (nDisplayed + 1 >= m_ListOfPhotos.Count)
                {
                    BitmapImage photoBmp = m_Global_Handler.Image_Handler.Load_BitmapImage(m_ListOfPhotos[0]);
                    Img_HostAndHostess_Photo.Source = photoBmp;
                    Img_HostAndHostess_Photo.Tag = 0;
                }
                else
                {
                    BitmapImage photoBmp = m_Global_Handler.Image_Handler.Load_BitmapImage(m_ListOfPhotos[nDisplayed + 1]);
                    Img_HostAndHostess_Photo.Source = photoBmp;
                    Img_HostAndHostess_Photo.Tag = nDisplayed + 1;
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
        /// Click on previous photo button
        /// </summary>
        private void Btn_HostAndHostess_PreviousPhoto_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Verification
                if (Img_HostAndHostess_Photo.Tag == null)
                {
                    return;
                }

                //Get  and display next photo
                int nDisplayed = (int)Img_HostAndHostess_Photo.Tag;
                if (nDisplayed - 1 < 0)
                {
                    BitmapImage photoBmp = m_Global_Handler.Image_Handler.Load_BitmapImage(m_ListOfPhotos[m_ListOfPhotos.Count - 1]);
                    Img_HostAndHostess_Photo.Source = photoBmp;
                    Img_HostAndHostess_Photo.Tag = m_ListOfPhotos.Count - 1;
                }
                else
                {
                    BitmapImage photoBmp = m_Global_Handler.Image_Handler.Load_BitmapImage(m_ListOfPhotos[nDisplayed - 1]);
                    Img_HostAndHostess_Photo.Source = photoBmp;
                    Img_HostAndHostess_Photo.Tag = nDisplayed - 1;
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
        /// Click on quit wihout saving button
        /// </summary>
        private void Btn_HostAndHostess_QuitWithoutSaving_Click(object sender, RoutedEventArgs e)
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
        private void Cmb_HostAndHostess_Country_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (m_IsLoaded == false)
            {
                return;
            }

            if (Cmb_HostAndHostess_Country.SelectedItem.ToString() == "United States")
            {
                Lbl_HostAndHostess_State.Visibility = Visibility.Visible;
                Txt_HostAndHostess_State.Visibility = Visibility.Visible;
            }
            else
            {
                Lbl_HostAndHostess_State.Visibility = Visibility.Hidden;
                Txt_HostAndHostess_State.Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        /// Event
        /// Text changed in the zip code text box
        /// </summary>
        private void Txt_HostAndHostess_ZipCode_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                //French cities
                if (Cmb_HostAndHostess_Country.Text == "France")
                {
                    //Verification
                    if (Txt_HostAndHostess_ZipCode.Text.Length < 5)
                    {
                        Cmb_HostAndHostess_City.Items.Clear();
                        return;
                    }
                    else if (Txt_HostAndHostess_ZipCode.Text.Length == 5)
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
                                        if (zipCode.Contains(Txt_HostAndHostess_ZipCode.Text))
                                        {
                                            string city = childNode.ChildNodes[4].InnerText;
                                            Cmb_HostAndHostess_City.Items.Add(city);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Cmb_HostAndHostess_City.Items.Clear();
                        return;
                    }
                }
                else
                {
                    //No treatment for the moment - TODO
                    Cmb_HostAndHostess_City.Items.Clear();
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
        /// Save the picture on the disc
        /// </summary>
        private void Save_PhotosToDisc(Hostess _HostOrHostess)
        {
            try
            {
                //Get repository
                string repository = SoftwareObjects.GlobalSettings.photos_repository + "\\" + _HostOrHostess.id;
                if (Directory.Exists(repository) == false)
                {
                    Directory.CreateDirectory(repository);
                }

                // Save files
                for (int iPhoto = 0; iPhoto < m_ListOfPhotos.Count; ++iPhoto)
                {
                    string fileName = Path.GetFileName(m_ListOfPhotos[iPhoto]);
                    string destFile = Path.Combine(repository, fileName);
                    if (m_ListOfPhotos[iPhoto] != destFile)
                    {
                        File.Copy(m_ListOfPhotos[iPhoto], destFile, true);
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

    }
}