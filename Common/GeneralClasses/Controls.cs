using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace GeneralClasses
{

    /// <summary>
    /// Class Controls containing all the functions for handling WPF controls
    /// </summary>
    public class Controls
    {

        #region Variables

        /// <summary>
        /// Error handler 
        /// </summary>
        private static Log m_Error_Handler = new Log();

        /// <summary>
        /// Boolean for invalid or valid email
        /// </summary>
        bool m_IsInvalid = false;

        #endregion

        #region Functions

        /// <summary>
        /// Cleans the fields
        /// <param name="_ListFields">List of fields to clean</param>
        /// <returns>True if the cleaning has been correctly done, false otherwise</returns>
        /// </summary>
        public bool Clean_Fields(List<System.Windows.Controls.Control> _ListFields)
        {
            try
            {
                //Verification
                if (_ListFields == null)
                {
                    throw new ArgumentException(System.Reflection.MethodBase.GetCurrentMethod().Name + " - Parameter null");
                }

                //Cleaning
                for (int i = 0; i < _ListFields.Count; ++i)
                {
                    if (_ListFields.GetType() == typeof(System.Windows.Controls.TextBox))
                    {
                        System.Windows.Controls.TextBox textbox = (System.Windows.Controls.TextBox)_ListFields[i];
                        textbox.Text = "";
                    }
                    else if (_ListFields.GetType() == typeof(System.Windows.Controls.ComboBox))
                    {
                        System.Windows.Controls.ComboBox combobox = (System.Windows.Controls.ComboBox)_ListFields[i];
                        combobox.Text = "";
                    }
                }
                return true;
            }
            catch (ArgumentException e)
            {
                Console.WriteLine("Parameter null exception.", e);
                return false;
            }
            catch (Exception e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                return false;
            }
        }

        /// <summary>
        /// Verifies which fields are blank in a list of fields
        /// <param name="_ListFields">List of all the fields</param>
        /// <param name="_NeededFields">Fields needed to be verify</param>
        /// <param name="_RessourceHandler">Ressource handler</param>
        /// <returns>A list of fields containing the blank fields</returns>
        /// </summary>
        public System.Windows.MessageBoxResult Verify_BlankFields(List<System.Windows.Controls.Control> _ListFields, List<string> _NeededFields, ResourcesManager _RessourceHandler)
        {
            //Initialization of the lists
            List<System.Windows.Controls.Control> ListBlankFields = new List<System.Windows.Controls.Control>();
            List<string> FieldsName = new List<string>();

            try
            {
                //Verification
                if (_ListFields == null)
                {
                    throw new ArgumentException(System.Reflection.MethodBase.GetCurrentMethod().Name + " - Parameter null");
                }

                //Get blanks fields
                for (int i = 0; i < _ListFields.Count; ++i)
                {
                    if (_ListFields[i].GetType() == typeof(System.Windows.Controls.TextBox))
                    {
                        System.Windows.Controls.TextBox textbox = (System.Windows.Controls.TextBox)_ListFields[i];
                        if (textbox.Text == "")
                        {
                            ListBlankFields.Add(textbox);
                            string[] FieldName = textbox.Name.Split('_');
                            if (FieldName.Length > 0)
                            {
                                FieldsName.Add(_RessourceHandler.Get_Resources(FieldName[FieldName.Length - 1]));
                            }
                        }
                    }
                    else if (_ListFields[i].GetType() == typeof(System.Windows.Controls.ComboBox))
                    {
                        System.Windows.Controls.ComboBox combobox = (System.Windows.Controls.ComboBox)_ListFields[i];
                        if (combobox.Text == "")
                        {
                            ListBlankFields.Add(combobox);
                            string[] FieldName = combobox.Name.Split('_');
                            if (FieldName.Length > 0)
                            {
                                FieldsName.Add(_RessourceHandler.Get_Resources(FieldName[FieldName.Length - 1]));
                            }
                        }
                    }
                }
                //Show message
                if (FieldsName.Count > 0)
                {
                    string blankFieldsStr = "";
                    for (int i = 0; i < FieldsName.Count; ++i)
                    {
                        if (_NeededFields != null)
                        {
                            if (!_NeededFields.Contains(FieldsName[i]))
                            {
                                continue;
                            }
                        }
                        if (blankFieldsStr != "")
                        {
                            blankFieldsStr = blankFieldsStr + ", ";
                        }
                        blankFieldsStr = blankFieldsStr + FieldsName[i];
                    }
                    if (blankFieldsStr != "")
                    {
                        string question = _RessourceHandler.Get_Resources("BlankFields") + " " + blankFieldsStr + "\n";
                        System.Windows.MessageBoxResult result = System.Windows.MessageBox.Show(question,
                            _RessourceHandler.Get_Resources("BlankFieldsTitle"),
                            System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        return result;
                    }
                }
                return System.Windows.MessageBoxResult.None;
            }
            catch (ArgumentException e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                return System.Windows.MessageBoxResult.None;
            }
            catch (Exception e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                return System.Windows.MessageBoxResult.None;
            }
        }

        /// <summary>
        /// Boolean text box changing
        /// </summary>
        private bool m_textBoxChanging = false;
        /// <summary>
        /// Validates or not a numeric text box
        /// <param name="_TextBox">the numeric text box</param>
        /// </summary>
        public void Validate_NumericTextBox(System.Windows.Controls.TextBox _TextBox)
        {
            try
            {
                // stop multiple changes;
                if (m_textBoxChanging)
                {
                    return;
                }
                m_textBoxChanging = true;

                string text = _TextBox.Text;
                if (text == "")
                {
                    return;
                }
                string validText = "";
                bool hasPeriod = false;
                int pos = _TextBox.SelectionStart;
                for (int iChar = 0; iChar < text.Length; ++iChar)
                {
                    bool badChar = false;
                    char character = text[iChar];
                    if (character == '.')
                    {
                        if (hasPeriod)
                            badChar = true;
                        else
                            hasPeriod = true;
                    }
                    else if (character < '0' || character > '9')
                        badChar = true;

                    if (!badChar)
                        validText += character;
                    else
                    {
                        if (iChar <= pos)
                            pos--;
                    }
                }

                // trim starting 00s
                while (validText.Length >= 2 && validText[0] == '0')
                {
                    if (validText[1] != '.')
                    {
                        validText = validText.Substring(1);
                        if (pos < 2)
                            pos--;
                    }
                    else
                    {
                        break;
                    }
                }

                if (pos > validText.Length)
                {
                    pos = validText.Length;
                }
                _TextBox.Text = validText;
                _TextBox.SelectionStart = pos;
                m_textBoxChanging = false;
            }
            catch (Exception exception)
            {
                m_Error_Handler.WriteMessage(System.Reflection.MethodBase.GetCurrentMethod().Name, exception.StackTrace);
                return;
            }
        }

        /// <summary>
        /// Validates if an email string is in an email format
        /// <param name="_Email">Email to test</param>
        /// <returns>True if the the email is valid, false otherwise</returns>
        /// </summary>
        public bool Is_ValidEmail(string _Email)
        {
            m_IsInvalid = false;
            if (string.IsNullOrEmpty(_Email))
                return false;

            // Use IdnMapping class to convert Unicode domain names.
            try
            {
                _Email = Regex.Replace(_Email, @"(@)(.+)$", this.DomainMapper,
                                      RegexOptions.None, TimeSpan.FromMilliseconds(200));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }

            if (m_IsInvalid)
                return false;

            // Return true if strIn is in valid e-mail format.
            try
            {
                return Regex.IsMatch(_Email,
                      @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                      @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                      RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }

        /// <summary>
        /// Wrap the domain
        /// <param name="_Match">Email to test</param>
        /// <returns>The domain wrapped/returns>
        /// </summary>
        private string DomainMapper(Match _Match)
        {
            // IdnMapping class with default property values.
            IdnMapping idn = new IdnMapping();

            string domainName = _Match.Groups[2].Value;
            try
            {
                domainName = idn.GetAscii(domainName);
            }
            catch (ArgumentException)
            {
                m_IsInvalid = true;
            }
            return _Match.Groups[1].Value + domainName;
        }

        #endregion

    }
}