using GeneralClasses;
using MySql.Data.MySqlClient;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;

namespace Database
{
    /// <summary>
    /// Class Database containing all the databse stuff
    /// </summary>
    public class Database
    {

        #region Variables

        /// <summary>
        /// Global handler
        /// </summary>
        public static Handlers m_Global_Handler = new Handlers();

        /// <summary>
        /// MySQL connection
        /// </summary>
        private MySqlConnection m_SQLConnection;

        #endregion

        #region Common

        /// <summary>
        /// Function which calculate the MD5
        /// <param name="_Input">Input key</param>
        /// <returns>The calculated key</returns>
        /// </summary>
        private string Calculate_MD5Hash()
        {
            try
            {
                // Creation of the datekey
                string dateKey = "S" + m_Global_Handler.DateAndTime_Handler.Get_PresentYear() + "L" + m_Global_Handler.DateAndTime_Handler.Get_PresentDay() + "C" + m_Global_Handler.DateAndTime_Handler.Get_PresentMonth();
                string input = dateKey;

                // step 1, calculate MD5 hash from input
                MD5 md5 = MD5.Create();
                byte[] inputBytes = Encoding.ASCII.GetBytes(input);
                byte[] hash = md5.ComputeHash(inputBytes);

                // step 2, convert byte array to hex string
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < hash.Length; ++i)
                {
                    sb.Append(hash[i].ToString("X2"));
                }
                return sb.ToString();
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return "";
            }
        }

        /// <summary>
        /// Template Deserialization function for JSON
        /// <param name="_JsonString">string containing the json</param>
        /// <returns>Returns the collection of objects red in the string</returns>
        /// </summary>
        public T Deserialize_JSON<T>(string _JsonString)
        {
            try
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(T));
                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_JsonString));
                T obj = (T)ser.ReadObject(ms);
                return obj;
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, exception);
                return default(T);
            }
        }

        #endregion

        #region Initialisation

        /// <summary>
        /// Constructor
        /// </summary>
        public Database(Handlers _GlobalHandler)
        {
            m_Global_Handler = _GlobalHandler;
            this.InitConnection();
        }


        /// <summary>
        /// Initialize the connection to the database
        /// </summary>
        private void InitConnection()
        {
            //Database repository
            string settingsFilePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Resources\\Settings.zad";
            string settingsFile = new Uri(settingsFilePath).LocalPath;
            bool is_filePresent = true;
            if (m_Global_Handler.File_Handler.File_Exists(settingsFile) == true)
            {
                //Read database file
                var listLines = new List<string>();
                var fileStream = new FileStream(settingsFile, FileMode.Open, FileAccess.Read);
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
                {
                    string line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        listLines.Add(line);
                    }
                }
                //Apply to login and licence
                if (listLines.Count > 0)
                {
                    SoftwareObjects.GlobalSettings.database_definition = listLines[0];
                }
                else
                {
                    //Corrupted file, create a new one
                    is_filePresent = false;
                }
            }
            else
            {
                is_filePresent = false;
            }

            //File not on the disk, create and save by default
            if (is_filePresent == false)
            {
                //Default settings
                SoftwareObjects.GlobalSettings.database_definition = "SERVER=127.0.0.1; DATABASE=zadgestion; UID=root; PASSWORD=";

                //Save infos
                string uriPath = settingsFilePath;
                string localPath = new Uri(uriPath).LocalPath;
                using (StreamWriter file = new StreamWriter(localPath, true))
                {
                    file.WriteLine(SoftwareObjects.GlobalSettings.database_definition);
                }

                //Exit function
                return;
            }

            // Creation of the connection chain
            string connectionString = SoftwareObjects.GlobalSettings.database_definition;
            this.m_SQLConnection = new MySqlConnection(connectionString);
        }

        #endregion

        #region Host and hostess

        /// <summary>
        /// Add a host or hostess to the ZadGestion's database
        /// <param name="_Address">Address of the host or hostess</param>
        /// <param name="_BirthCity">City of birth of the host or hostess</param>
        /// <param name="_BirthDate">Date of birth of the host or hostess</param>
        /// <param name="_CellPhone">Cell phone of the host or hostess</param>
        /// <param name="_City">City of the host or hostess</param>
        /// <param name="_Country">Country of the host or hostess</param>
        /// <param name="_Email">E-mail of the host or hostess</param>
        /// <param name="_FirstName">First name of the host or hostess</param>
        /// <param name="_HasCar">Boolean defining if the host or hostess has a car or not</param>
        /// <param name="_HasDriverLicence">Boolean defining if the host or hostess has a driver licence or not</param>
        /// <param name="_Id">Id of the host or hostess</param>
        /// <param name="_IdPaycheck">Id of the paycheck of the host or hostess</param>
        /// <param name="_LanguageEnglish">Boolean defining if the host or hostess speaks english</param>
        /// <param name="_LanguageGerman">Boolean defining if the host or hostess speaks german</param>
        /// <param name="_LanguageItalian">Boolean defining if the host or hostess speaks italian</param>
        /// <param name="_LanguageOthers">Others languages spoken by the host or hostess</param>
        /// <param name="_LanguageSpanish">Boolean defining if the host or hostess speaks spanish</param>
        /// <param name="_LastName">Last name of the host or hostess</param>
        /// <param name="_ProfileEvent">Boolean defining if the host or hostess has an event profile</param>
        /// <param name="_ProfilePermanent">Boolean defining if the host or hostess has a permanent profile</param>
        /// <param name="_ProfileStreet">Boolean defining if the host or hostess has a street profile</param>
        /// <param name="_Sex">Sex of the host or hostess</param>
        /// <param name="_Size">Size of the host or hostess</param>
        /// <param name="_SizePants">Pants size of the host or hostess</param>
        /// <param name="_SizeShirt">Shirt size of the host or hostess</param>
        /// <param name="_SizeShoes">Shoes size of the host or hostess</param>
        /// <param name="_SocialNumber">State of the host or hostess</param>
        /// <param name="_State">State of the host or hostess</param>
        /// <param name="_ZipCode">Zip code of the host or hostess</param>
        /// <returns>The string containing the result of the operation (Id of the contact created, error)</returns>
        /// </summary>
        public string Add_HostAndHostessToDatabase(string _Address, string _BirthCity, string _BirthDate, string _CellPhone, string _City, string _Country,
            string _Email, string _FirstName, bool _HasCar, bool _HasDriverLicence, string _Id, string _IdPaycheck, bool _LanguageEnglish,
            bool _LanguageGerman, bool _LanguageItalian, string _LanguageOthers, bool _LanguageSpanish, string _LastName, bool _ProfileEvent,
            bool _ProfilePermanent, bool _ProfileStreet, string _Sex, string _Size, string _SizePants,
            string _SizeShirt, string _SizeShoes, string _SocialNumber, string _State, string _ZipCode)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "INSERT INTO hostsandhostesses (address, birth_city, birth_date, cellphone, city, country, date_creation," +
                    " email, firstname, has_car, has_driver_licence, id, id_paycheck, language_english, language_german, language_italian," +
                    " language_others, language_spanish, lastname, profile_event, profile_permanent, profile_street, sex, size, size_pants," +
                    " size_shirt, size_shoes, social_number, state, zipcode)" +
                    " VALUES (@address, @birth_city, @birth_date, @cellphone, @city, @country, @date_creation, @email, @firstname, @has_car," +
                    " @has_driver_licence, @id, @id_paycheck, @language_english, @language_german, @language_italian, @language_others, @language_spanish," +
                    " @lastname, @profile_event, @profile_permanent, @profile_street, @sex, @size, @size_pants, @size_shirt, @size_shoes," +
                    " @social_number, @state, @zipcode)";

                //Fill SQL parameters
                cmd.Parameters.AddWithValue("@address", _Address);
                cmd.Parameters.AddWithValue("@birth_city", _BirthCity);
                cmd.Parameters.AddWithValue("@birth_date", _BirthDate);
                cmd.Parameters.AddWithValue("@cellphone", _CellPhone);
                cmd.Parameters.AddWithValue("@city", _City);
                cmd.Parameters.AddWithValue("@country", _Country);
                cmd.Parameters.AddWithValue("@date_creation", DateTime.Today);
                cmd.Parameters.AddWithValue("@email", _Email);
                cmd.Parameters.AddWithValue("@firstname", _FirstName);
                cmd.Parameters.AddWithValue("@has_car", _HasCar);
                cmd.Parameters.AddWithValue("@has_driver_licence", _HasDriverLicence);
                cmd.Parameters.AddWithValue("@id", _Id);
                cmd.Parameters.AddWithValue("@id_paycheck", _IdPaycheck);
                cmd.Parameters.AddWithValue("@language_english", _LanguageEnglish);
                cmd.Parameters.AddWithValue("@language_german", _LanguageGerman);
                cmd.Parameters.AddWithValue("@language_italian", _LanguageItalian);
                cmd.Parameters.AddWithValue("@language_others", _LanguageOthers);
                cmd.Parameters.AddWithValue("@language_spanish", _LanguageSpanish);
                cmd.Parameters.AddWithValue("@lastname", _LastName);
                cmd.Parameters.AddWithValue("@profile_event", _ProfileEvent);
                cmd.Parameters.AddWithValue("@profile_permanent", _ProfilePermanent);
                cmd.Parameters.AddWithValue("@profile_street", _ProfileStreet);
                cmd.Parameters.AddWithValue("@sex", _Sex);
                cmd.Parameters.AddWithValue("@size", _Size);
                cmd.Parameters.AddWithValue("@size_pants", _SizePants);
                cmd.Parameters.AddWithValue("@size_shirt", _SizeShirt);
                cmd.Parameters.AddWithValue("@size_shoes", _SizeShoes);
                cmd.Parameters.AddWithValue("@social_number", _SocialNumber);
                cmd.Parameters.AddWithValue("@state", _State);
                cmd.Parameters.AddWithValue("@zipcode", _ZipCode);

                //Execute request
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                //Close connection
                this.m_SQLConnection.Close();

                //Write error to log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                if (e.InnerException != null)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e.InnerException);
                }
                return "Error - " + e.Message;
            }
        }

        /// <summary>
        /// Delete a host or hostess to the ZadGestion's database
        /// <param name="_Id">Id of the host or hostess</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Delete_HostAndHostessToDatabase(string _Id)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "DELETE FROM hostsandhostesses WHERE id = '" + _Id + "'";

                //Execute request
                cmd.ExecuteNonQuery();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                //Close connection
                this.m_SQLConnection.Close();

                //Write error to log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return "Error - " + e.Message;
            }
        }

        /// <summary>
        /// Edit a host or hostess to the ZadGestion's database
        /// <param name="_Address">Address of the host or hostess</param>
        /// <param name="_BirthCity">City of birth of the host or hostess</param>
        /// <param name="_BirthDate">Date of birth of the host or hostess</param>
        /// <param name="_CellPhone">Cell phone of the host or hostess</param>
        /// <param name="_City">City of the host or hostess</param>
        /// <param name="_Country">Country of the host or hostess</param>
        /// <param name="_Email">E-mail of the host or hostess</param>
        /// <param name="_FirstName">First name of the host or hostess</param>
        /// <param name="_HasCar">Boolean defining if the host or hostess has a car or not</param>
        /// <param name="_HasDriverLicence">Boolean defining if the host or hostess has a driver licence or not</param>
        /// <param name="_Id">Id of the host or hostess</param>
        /// <param name="_IdPaycheck">Id of the paycheck of the host or hostess</param>
        /// <param name="_LanguageEnglish">Boolean defining if the host or hostess speaks english</param>
        /// <param name="_LanguageGerman">Boolean defining if the host or hostess speaks german</param>
        /// <param name="_LanguageItalian">Boolean defining if the host or hostess speaks italian</param>
        /// <param name="_LanguageOthers">Others languages spoken by the host or hostess</param>
        /// <param name="_LanguageSpanish">Boolean defining if the host or hostess speaks spanish</param>
        /// <param name="_LastName">Last name of the host or hostess</param>
        /// <param name="_ProfileEvent">Boolean defining if the host or hostess has an event profile</param>
        /// <param name="_ProfilePermanent">Boolean defining if the host or hostess has a permanent profile</param>
        /// <param name="_ProfileStreet">Boolean defining if the host or hostess has a street profile</param>
        /// <param name="_Sex">Sex of the host or hostess</param>
        /// <param name="_Size">Size of the host or hostess</param>
        /// <param name="_SizePants">Pants size of the host or hostess</param>
        /// <param name="_SizeShirt">Shirt size of the host or hostess</param>
        /// <param name="_SizeShoes">Shoes size of the host or hostess</param>
        /// <param name="_SocialNumber">State of the host or hostess</param>
        /// <param name="_State">State of the host or hostess</param>
        /// <param name="_ZipCode">Zip code of the host or hostess</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Edit_HostAndHostessToDatabase(string _Address, string _BirthCity, string _BirthDate, string _CellPhone, string _City, string _Country,
            string _Email, string _FirstName, bool _HasCar, bool _HasDriverLicence, string _Id, string _IdPaycheck, bool _LanguageEnglish,
            bool _LanguageGerman, bool _LanguageItalian, string _LanguageOthers, bool _LanguageSpanish, string _LastName, bool _ProfileEvent,
            bool _ProfilePermanent, bool _ProfileStreet, string _Sex, string _Size, string _SizePants,
            string _SizeShirt, string _SizeShoes, string _SocialNumber, string _State, string _ZipCode)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "UPDATE hostsandhostesses SET address = @address, birth_city = @birth_city, birth_date = @birth_date," +
                    " cellphone = @cellphone, city = @city, country = @country, email = @email, firstname = @firstname, id_paycheck = @id_paycheck," +
                    " language_english = @language_english, language_german = @language_german, language_italian = @language_italian," +
                    " language_others = @language_others, language_spanish = @language_spanish, lastname = @lastname, profile_event = @profile_event," +
                    " profile_permanent = @profile_permanent,profile_street = @profile_street, sex = @sex, size = @size, size_pants = @size_pants," +
                    " size_shirt = @size_shirt, size_shoes = @size_shoes, social_number= @social_number, state = @state, zipcode = @zipcode WHERE id = @id";


                //Fill SQL parameters
                cmd.Parameters.AddWithValue("@address", _Address);
                cmd.Parameters.AddWithValue("@birth_city", _BirthCity);
                cmd.Parameters.AddWithValue("@birth_date", _BirthDate);
                cmd.Parameters.AddWithValue("@cellphone", _CellPhone);
                cmd.Parameters.AddWithValue("@city", _City);
                cmd.Parameters.AddWithValue("@country", _Country);
                cmd.Parameters.AddWithValue("@email", _Email);
                cmd.Parameters.AddWithValue("@firstname", _FirstName);
                cmd.Parameters.AddWithValue("@has_car", _HasCar);
                cmd.Parameters.AddWithValue("@has_driver_licence", _HasDriverLicence);
                cmd.Parameters.AddWithValue("@id", _Id);
                cmd.Parameters.AddWithValue("@id_paycheck", _IdPaycheck);
                cmd.Parameters.AddWithValue("@language_english", _LanguageEnglish);
                cmd.Parameters.AddWithValue("@language_german", _LanguageGerman);
                cmd.Parameters.AddWithValue("@language_italian", _LanguageItalian);
                cmd.Parameters.AddWithValue("@language_others", _LanguageOthers);
                cmd.Parameters.AddWithValue("@language_spanish", _LanguageSpanish);
                cmd.Parameters.AddWithValue("@lastname", _LastName);
                cmd.Parameters.AddWithValue("@profile_event", _ProfileEvent);
                cmd.Parameters.AddWithValue("@profile_permanent", _ProfilePermanent);
                cmd.Parameters.AddWithValue("@profile_street", _ProfileStreet);
                cmd.Parameters.AddWithValue("@sex", _Sex);
                cmd.Parameters.AddWithValue("@size", _Size);
                cmd.Parameters.AddWithValue("@size_pants", _SizePants);
                cmd.Parameters.AddWithValue("@size_shirt", _SizeShirt);
                cmd.Parameters.AddWithValue("@size_shoes", _SizeShoes);
                cmd.Parameters.AddWithValue("@social_number", _SocialNumber);
                cmd.Parameters.AddWithValue("@state", _State);
                cmd.Parameters.AddWithValue("@zipcode", _ZipCode);

                //Execute request
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                //Close connection
                this.m_SQLConnection.Close();

                //Write error to log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return "Error - " + e.Message;
            }
        }

        /// <summary>
        /// Get all the hosts and hostesses from ZadGestion's database
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Get_HostsAndHostessesFromDatabase()
        {
            try
            {
                //Clear collection
                SoftwareObjects.HostsAndHotessesCollection.Clear();

                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "SELECT * FROM hostsandhostesses";

                //Execute request
                MySqlDataAdapter mySQLDatabaseAdapter = new MySqlDataAdapter(cmd);
                DataTable table = new DataTable();
                mySQLDatabaseAdapter.Fill(table);

                //Fill collection
                SoftwareObjects.HostsAndHotessesCollection = (from DataRow row in table.Rows.OfType<DataRow>()
                                                              select new Hostess
                                                              {
                                                                  address = row["address"].ToString(),
                                                                  archived = (int)row["archived"],
                                                                  birth_city = row["birth_city"].ToString(),
                                                                  birth_date = row["birth_date"].ToString(),
                                                                  cellphone = row["cellphone"].ToString(),
                                                                  city = row["city"].ToString(),
                                                                  country = row["country"].ToString(),
                                                                  date_creation = row["date_creation"].ToString(),
                                                                  email = row["email"].ToString(),
                                                                  firstname = row["firstname"].ToString(),
                                                                  has_car = (bool)row["has_car"],
                                                                  has_driver_licence = (bool)row["has_driver_licence"],
                                                                  id = row["id"].ToString(),
                                                                  id_paycheck = row["id_paycheck"].ToString(),
                                                                  language_english = (bool)row["language_english"],
                                                                  language_german = (bool)row["language_german"],
                                                                  language_italian = (bool)row["language_italian"],
                                                                  language_others = row["language_others"].ToString(),
                                                                  language_spanish = (bool)row["language_spanish"],
                                                                  lastname = row["lastname"].ToString(),
                                                                  profile_event = (bool)row["profile_event"],
                                                                  profile_permanent = (bool)row["profile_permanent"],
                                                                  profile_street = (bool)row["profile_street"],
                                                                  sex = row["sex"].ToString(),
                                                                  size = row["size"].ToString(),
                                                                  size_pants = row["size_pants"].ToString(),
                                                                  size_shirt = row["size_shirt"].ToString(),
                                                                  size_shoes = row["size_shoes"].ToString(),
                                                                  social_number = row["social_number"].ToString(),
                                                                  state = row["state"].ToString(),
                                                                  zipcode = row["zipcode"].ToString()
                                                              }).ToList();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                //Close connection
                this.m_SQLConnection.Close();

                //Write error to log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return "Error - " + e.Message;
            }
        }

        #endregion

        #region Client

        /// <summary>
        /// Archive a host or hostess to the ZadGestion's database
        /// <param name="_Id">Id of the host or hostess</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Archive_HostOrHostess(string _Id)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "UPDATE hostsandhostesses SET archived = @archived WHERE id = @id";

                //Fill SQL parameters
                cmd.Parameters.AddWithValue("@id", _Id);
                cmd.Parameters.AddWithValue("@archived", 1);

                //Execute request
                cmd.ExecuteNonQuery();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        /// <summary>
        /// Add a host or hostess to the ZadGestion's database
        /// <param name="_Address">Address of the client</param>
        /// <param name="_City">City of the client</param>
        /// <param name="_CorporateName">Corporate name of the client</param>
        /// <param name="_CorporateNumber">Corporate number of the client</param>
        /// <param name="_Country">Country of the client</param>
        /// <param name="_Email">E-mail of the client</param>
        /// <param name="_Id">Id of the client</param>
        /// <param name="_Phone">Phone of the client</param>
        /// <param name="_State">State of the client</param>
        /// <param name="_ZipCode">Zip code of the client</param>
        /// <returns>The string containing the result of the operation (Id of the contact created, error)</returns>
        /// </summary>
        public string Add_ClientToDatabase(string _Address, string _City, string _CorporateName, string _CorporateNumber, string _Country,
            string _Email, string _Id, string _Phone, string _State, string _ZipCode)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "INSERT INTO clients (address, city, corporate_name, corporate_number, country, date_creation, email, id, phone, state, zipcode)" +
                    " VALUES (@address, @city, @corporate_name, @corporate_number, @country, @date_creation, @email, @id, @phone, @state, @zipcode)";

                //Fill SQL parameters
                cmd.Parameters.AddWithValue("@address", _Address);
                cmd.Parameters.AddWithValue("@city", _City);
                cmd.Parameters.AddWithValue("@corporate_name", _CorporateName);
                cmd.Parameters.AddWithValue("@corporate_number", _CorporateNumber);
                cmd.Parameters.AddWithValue("@country", _Country);
                cmd.Parameters.AddWithValue("@date_creation", DateTime.Today);
                cmd.Parameters.AddWithValue("@email", _Email);
                cmd.Parameters.AddWithValue("@id", _Id);
                cmd.Parameters.AddWithValue("@phone", _Phone);
                cmd.Parameters.AddWithValue("@state", _State);
                cmd.Parameters.AddWithValue("@zipcode", _ZipCode);

                //Execute request
                cmd.ExecuteNonQuery();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                //Close connection
                this.m_SQLConnection.Close();

                //Write error to log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                if (e.InnerException != null)
                {
                    m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e.InnerException);
                }
                return "Error - " + e.Message;
            }
        }

        /// <summary>
        /// Delete a client to the ZadGestion's database
        /// <param name="_Id">Id of the client</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Delete_ClientToDatabase(string _Id)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "DELETE FROM clients WHERE id = '" + _Id + "'";

                //Execute request
                cmd.ExecuteNonQuery();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                //Close connection
                this.m_SQLConnection.Close();

                //Write error to log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return "Error - " + e.Message;
            }
        }

        /// <summary>
        /// Edit a client to the ZadGestion's database
        /// <param name="_Address">Address of the client</param>
        /// <param name="_City">City of the client</param>
        /// <param name="_CorporateName">Corporate name of the client</param>
        /// <param name="_CorporateNumber">Corporate number of the client</param>
        /// <param name="_Country">Country of the client</param>
        /// <param name="_Email">E-mail of the client</param>
        /// <param name="_Id">Id of the client</param>
        /// <param name="_Phone">Phone of the client</param>
        /// <param name="_State">State of the client</param>
        /// <param name="_ZipCode">Zip code of the client</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Edit_ClientToDatabase(string _Address, string _City, string _CorporateName, string _CorporateNumber, string _Country,
            string _Email, string _Id, string _Phone, string _State, string _ZipCode)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "UPDATE clients SET address = @address, city = @city, corporate_name = @corporate_name, corporate_number = @corporate_number, " +
                    "country = @country, email = @email, phone = @phone, state = @state, zipcode = @zipcode WHERE id = @id";

                //Fill SQL parameters
                cmd.Parameters.AddWithValue("@address", _Address);
                cmd.Parameters.AddWithValue("@city", _City);
                cmd.Parameters.AddWithValue("@corporate_name", _CorporateName);
                cmd.Parameters.AddWithValue("@corporate_number", _CorporateNumber);
                cmd.Parameters.AddWithValue("@country", _Country);
                cmd.Parameters.AddWithValue("@email", _Email);
                cmd.Parameters.AddWithValue("@id", _Id);
                cmd.Parameters.AddWithValue("@phone", _Phone);
                cmd.Parameters.AddWithValue("@state", _State);
                cmd.Parameters.AddWithValue("@zipcode", _ZipCode);

                //Execute request
                cmd.ExecuteNonQuery();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                //Close connection
                this.m_SQLConnection.Close();

                //Write error to log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return "Error - " + e.Message;
            }
        }

        /// <summary>
        /// Get all the clients from ZadGestion's database
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Get_ClientsFromDatabase()
        {
            try
            {
                //Clear collection
                SoftwareObjects.ClientsCollection.Clear();

                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "SELECT * FROM clients";

                //Execute request
                MySqlDataAdapter mySQLDatabaseAdapter = new MySqlDataAdapter(cmd);
                DataTable table = new DataTable();
                mySQLDatabaseAdapter.Fill(table);

                //Fill collection
                SoftwareObjects.ClientsCollection = (from DataRow row in table.Rows.OfType<DataRow>()
                                                     select new Client
                                                     {
                                                         address = row["address"].ToString(),
                                                         archived = (bool)row["archived"],
                                                         city = row["city"].ToString(),
                                                         corporate_name = row["corporate_name"].ToString(),
                                                         corporate_number = row["corporate_number"].ToString(),
                                                         country = row["country"].ToString(),
                                                         date_creation = row["date_creation"].ToString(),
                                                         email = row["email"].ToString(),
                                                         id = row["id"].ToString(),
                                                         phone = row["phone"].ToString(),
                                                         state = row["state"].ToString(),
                                                         zipcode = row["zipcode"].ToString()
                                                     }).ToList();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                //Close connection
                this.m_SQLConnection.Close();

                //Write error to log
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return "Error - " + e.Message;
            }
        }

        /// <summary>
        /// Restore a host or hostess to the ZadGestion's database
        /// <param name="_Id">Id of the host or hostess</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Restore_HostOrHostess(string _Id)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = "UPDATE hostsandhostesses SET archived = @archived WHERE id = @id";

                //Fill SQL parameters
                cmd.Parameters.AddWithValue("@id", _Id);
                cmd.Parameters.AddWithValue("@archived", 0);

                //Execute request
                cmd.ExecuteNonQuery();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        #endregion

        #region Mission

        /// <summary>
        /// Get all the missions from ZadGestion's database
        /// <param name="_GetClosed">Boolean indicating if closed missions should be pulled</param>
        /// <returns>The string containing all the missions</returns>
        /// </summary>
        public string Get_MissionsFromDatabaseDatabase(bool _GetClosed)
        {
            try
            {
                return "";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        /// <summary>
        /// Get the mission from its Id from ZadGestion's database
        /// <param name="_Id">Id of the mission asked</param>
        /// <returns>The string containing the mission</returns>
        /// </summary>
        public string Get_MissionByIdFromDatabaseDatabase(int _Id)
        {
            try
            {
                return "";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        /// <summary>
        /// Edit an mission to the ZadGestion's database
        /// <param name="_Id">Id of the mission</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Edit_MissionToDatabase(int _Id)
        {
            try
            {
                return "";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        /// <summary>
        /// Add a mission to the ZadGestion's database
        /// <returns>The string containing the result of the operation (Id of the mission created, error)</returns>
        /// </summary>
        public string Add_MissionToDatabase()
        {
            try
            {
                return "";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        /// <summary>
        /// Delete an mission to the ZadGestion's database
        /// <param name="_Id">Id of the mission</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Delete_MissionToDatabase(int _Id)
        {
            try
            {
                return "";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        /// <summary>
        /// Close a mission to the ZadGestion's database
        /// <param name="_Id">Id of the mission</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Close_Mission(int _Id)
        {
            try
            {
                return "";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        /// <summary>
        /// Restore a mission to the ZadGestion's database
        /// <param name="_Id">Id of the mission</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Reopen_Mission(int _Id)
        {
            try
            {
                return "";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);
                return e.Message;
            }
        }

        #endregion

        #region Settings

        #region General

        /// <summary>
        /// Edit the directory of the photo in the database
        /// <param name="_PhotosDirectory">Photo directory</param>
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Edit_Settings_General_PhotosDirectory(string _PhotosDirectory)
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                _PhotosDirectory = _PhotosDirectory.Replace(@"\", @"\\");
                cmd.CommandText = "UPDATE settings SET photos_repository = '" + _PhotosDirectory + "' WHERE id = 0";

                //Execute request
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);

                //Close connection
                this.m_SQLConnection.Close();

                return "Error - " + e.Message;
            }
        }

        /// <summary>
        /// Get the setting form the database
        /// <returns>The string containing the result of the operation (OK, error)</returns>
        /// </summary>
        public string Get_SettingsFromDatabase()
        {
            try
            {
                //Open SQL connection
                this.m_SQLConnection.Open();

                //Create SQL command
                MySqlCommand cmd = this.m_SQLConnection.CreateCommand();

                //SQL request
                cmd.CommandText = " SELECT * FROM settings  ";

                //Execute request
                MySqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    SoftwareObjects.GlobalSettings.photos_repository = reader.GetString("photos_repository");
                }

                //Close connection
                this.m_SQLConnection.Close();

                return "OK";
            }
            catch (Exception e)
            {
                m_Global_Handler.Log_Handler.WriteException(MethodBase.GetCurrentMethod().Name, e);

                //Close connection
                this.m_SQLConnection.Close();

                return e.Message;
            }
        }

        #endregion

        #endregion


    }
}
