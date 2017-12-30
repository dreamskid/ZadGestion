using Microsoft.Office.Interop.Word;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace GeneralClasses
{

    /// <summary>
    /// Class WordGeneration containing all the tools to generate a bill to word
    /// </summary>
    public class WordGeneration
    {

        #region Variables

        /// <summary>
        /// Global handler
        /// </summary>
        private static Handlers m_Global_Handler = new Handlers();

        /// <summary>
        /// Bill to treat
        /// </summary>
        private static Mission m_Mission = new Mission();

        /// <summary>
        /// Word application
        /// </summary>
        private Microsoft.Office.Interop.Word.Application m_AppWord = null;

        /// <summary>
        /// Word document to generate
        /// </summary>
        private Document m_NewWordDoc;

        /// <summary>
        /// Missions generated filename
        /// </summary>
        private string m_GeneratedBillFileName = "";

        /// <summary>
        /// Error handler
        /// </summary>
        private static Log m_Error_Handler = new Log();

        /// <summary>
        /// Application directory
        /// </summary>
        private static string m_AppDir = "";

        /// <summary>
        /// Object missing definition
        /// </summary>
        private static object m_Missing = System.Reflection.Missing.Value;

        /// <summary>
        /// Object for adding row in tables
        /// </summary>
        private Object m_BeforeRow = Type.Missing;

        #endregion

        #region Getter/Setter

        /// <summary>
        /// Getter for the generated bill
        /// </summary>
        public string Get_GeneratedBillFileName()
        {
            return m_GeneratedBillFileName;
        }

        #endregion

        #region Functions

        /// <summary>
        /// Replace a text in the word template in the new one
        /// <param name="_TextToReplace">Text in the template</param>
        /// <param name="_NewText">New text</param>
        /// <param name="_MsWord">Document where to replace the text</param>
        /// <returns>True if everything went well, false otherwise</returns>
        /// </summary>
        private bool Replace_Text(string _TextToReplace, string _NewText, Microsoft.Office.Interop.Word.Application _MsWord)
        {
            try
            {
                //Test lenght
                if (_NewText.Length > 255)
                {
                    return Replace_LargeText(_TextToReplace, _NewText, _MsWord);
                }

                //Object missing definition
                m_Missing = System.Reflection.Missing.Value;

                //Find the text to replace
                Find findObject = _MsWord.Selection.Find;
                findObject.ClearFormatting();
                findObject.Text = _TextToReplace;
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = _NewText;

                //Execute the operation
                object replaceAll = WdReplace.wdReplaceAll;
                findObject.Execute(ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing,
                    ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing,
                    ref replaceAll, ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing);

                //Everything ok
                return true;
            }
            catch (Exception exception)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, exception);
                return false;
            }
        }

        /// <summary>
        /// Replace a large text in the word template in the new one
        /// <param name="_TextToReplace">Text in the template</param>
        /// <param name="_NewText">New text</param>
        /// <param name="_MsWord">Document where to replace the text</param>
        /// <returns>True if everything went well, false otherwise</returns>
        /// </summary>
        private bool Replace_LargeText(string _TextToReplace, string _NewText, Microsoft.Office.Interop.Word.Application _MsWord)
        {
            try
            {
                // Find text 
                Range range = _MsWord.ActiveDocument.Content;
                range.Find.Execute(_TextToReplace);
                //Clear text
                Replace_Text(_TextToReplace, "", _MsWord);
                // Define new range 
                range = _MsWord.ActiveDocument.Range(range.End, range.End);
                range.Text = _NewText;

                return true;
            }
            catch (Exception exception)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, exception);
                return false;
            }
        }

        /// <summary>
        /// Generate the bill
        /// <param name="_Global_Handler">Global handler to get the ressources</param>
        /// <param name="_Bill">Bill to generate</param>
        /// <param name="_IsMission">Generation of an Mission or a bill</param>
        /// <param name="_Is_Visible">Word visible during generation or not</param>
        /// <param name="_FinalFilename">PDF output file name</param>
        /// <returns>0 if everything went well, 1 otherwise</returns>
        /// </summary>
        public int Generate_Bill(Handlers _Global_Handler, Mission _Bill, bool _Is_Visible, string _FinalFilename)
        {
            try
            {
                //Initialize
                m_GeneratedBillFileName = _FinalFilename;
                m_Global_Handler = _Global_Handler;
                m_Mission = _Bill;

                //Connection to word
                m_AppWord = new Microsoft.Office.Interop.Word.Application();
                if (m_AppWord == null)
                {
                    MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("MicrosoftWordNotFound"),
                        m_Global_Handler.Resources_Handler.Get_Resources("MicrosoftWordNotFoundCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return -1;
                }
                m_AppWord.Visible = _Is_Visible;

                //Choose template
                m_AppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
                string fullFilePath = Path.Combine(m_AppDir, "Template\\Template.docx");
                object templateName = fullFilePath;

                //Create document
                m_NewWordDoc = m_AppWord.Documents.Add(ref templateName, ref m_Missing, ref m_Missing, ref m_Missing);

                //Logo
                bool resLogo = Generate_Logo();
                if (resLogo == false)
                {
                    return -1;
                }

                //Replace fields
                bool resFields = Replace_Fields();
                if (resFields == false)
                {
                    return -1;
                }

                //Finalization
                bool resFinalization = Finalize();
                if (resFinalization == false)
                {
                    return -1;
                }

                return 1;
            }
            catch (Exception error)
            {
                m_Global_Handler.Log_Handler.WriteMessage(System.Reflection.MethodBase.GetCurrentMethod().Name, error.StackTrace);
                MessageBox.Show(error.Message, m_Global_Handler.Resources_Handler.Get_Resources("Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);

                //Close document and word
                if (m_NewWordDoc != null)
                {
                    ((_Document)m_NewWordDoc).Close();
                }
                if (m_AppWord != null)
                {
                    ((_Application)m_AppWord).Quit();
                }

                return -1;
            }

        }

        /// <summary>
        /// Generate the bill
        /// <param name="_Global_Handler">Global handler to get the ressources</param>
        /// <param name="_CreditNote">Credit note to generate</param>
        /// <param name="_Is_Visible">Word visible during generation or not</param>
        /// <param name="_FinalFilename">PDF output file name</param>
        /// <returns>0 if everything went well, 1 otherwise</returns>
        /// </summary>
        public int Generate_CreditNote(Handlers _Global_Handler, Mission _CreditNote, bool _Is_Visible, string _FinalFilename)
        {
            try
            {
                //Initialize
                m_GeneratedBillFileName = _FinalFilename;
                m_Global_Handler = _Global_Handler;
                m_Mission = _CreditNote;

                //Connection to word
                m_AppWord = new Microsoft.Office.Interop.Word.Application();
                if (m_AppWord == null)
                {
                    System.Windows.MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("MicrosoftWordNotFound"),
                        m_Global_Handler.Resources_Handler.Get_Resources("MicrosoftWordNotFoundCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return -1;
                }
                m_AppWord.Visible = _Is_Visible;

                //Choose template
                m_AppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
                string fullFilePath = Path.Combine(m_AppDir, "Template\\TemplateCreditNote.docx");
                object templateName = fullFilePath;

                //Create document
                m_NewWordDoc = m_AppWord.Documents.Add(ref templateName, ref m_Missing, ref m_Missing, ref m_Missing);

                //Logo
                bool resLogo = Generate_Logo();
                if (resLogo == false)
                {
                    return -1;
                }

                //Replace fields
                bool resFields = Replace_Fields();
                if (resFields == false)
                {
                    return -1;
                }

                //Finalization
                bool resFinalization = Finalize();
                if (resFinalization == false)
                {
                    return -1;
                }

                return 1;
            }
            catch (Exception error)
            {
                m_Global_Handler.Log_Handler.WriteMessage(System.Reflection.MethodBase.GetCurrentMethod().Name, error.StackTrace);
                System.Windows.MessageBox.Show(error.Message,
                    m_Global_Handler.Resources_Handler.Get_Resources("Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);

                //Close document and word
                if (m_NewWordDoc != null)
                {
                    ((_Document)m_NewWordDoc).Close();
                }
                if (m_AppWord != null)
                {
                    ((_Application)m_AppWord).Quit();
                }

                return -1;
            }

        }

        /// <summary>
        /// Generate the logo on the word document
        /// <returns>True if everything went well, false otherwise</returns>
        /// </summary>
        private bool Generate_Logo()
        {
            try
            {
                //Logo
                string logoFilename = Path.Combine(m_AppDir, "Resources\\Logo.png");
                if (m_Global_Handler.File_Handler.File_Exists(logoFilename) == true)
                {
                    //Look for the range of the logo picture
                    List<Range> ranges = new List<Range>();
                    foreach (InlineShape shape in m_NewWordDoc.InlineShapes)
                    {
                        if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                        {
                            //Logo is the first shape of type picture
                            ranges.Add(shape.Range);
                            shape.Delete();
                            break;
                        }
                    }
                    Range range = ranges[0];

                    //Replace the picture
                    range.InlineShapes.AddPicture(logoFilename, ref m_Missing, ref m_Missing, ref m_Missing);
                }

                return true;
            }
            catch (Exception error)
            {
                m_Global_Handler.Log_Handler.WriteMessage(System.Reflection.MethodBase.GetCurrentMethod().Name, error.StackTrace);

                //Close document and word
                Close_Word();

                return false;
            }

        }

        /// <summary>
        /// Replace the fields on the word document
        /// <returns>True if everything went well, false otherwise</returns>
        /// </summary>
        private bool Replace_Fields()
        {
            try
            {/*
                
                //Contact informations
                Hostess contactSel = SoftwareObjects.HostsAndHotessesCollection.Find(x => x.id.Equals(m_Mission.id_hostorhostess));
                if (contactSel == null)
                {
                    string contactNotFound = "";
                    string contactNotFoundCaption = "";
                    if (m_Type == BillType.Mission)
                    {
                        contactNotFound = m_Global_Handler.Resources_Handler.Get_Resources("MissionContactNotFound");
                        contactNotFoundCaption = m_Global_Handler.Resources_Handler.Get_Resources("MissionContactNotFoundCaption");
                    }
                    else if (m_Type == BillType.INVOICE)
                    {
                        contactNotFound = m_Global_Handler.Resources_Handler.Get_Resources("InvoiceContactNotFound");
                        contactNotFoundCaption = m_Global_Handler.Resources_Handler.Get_Resources("InvoiceContactNotFoundCaption");
                    }
                    else if (m_Type == BillType.CREDITNOTE)
                    {
                        contactNotFound = m_Global_Handler.Resources_Handler.Get_Resources("CreditNoteContactNotFound");
                        contactNotFoundCaption = m_Global_Handler.Resources_Handler.Get_Resources("CreditNoteContactNotFoundCaption");
                    }
                    MessageBox.Show(contactNotFound, contactNotFoundCaption, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    contactSel = new Hostess();
                }
                Replace_Text("SLC_HostAndHostess_FirstName", contactSel.firstname, m_AppWord);
                Replace_Text("SLC_HostAndHostess_LastName", contactSel.lastname, m_AppWord);
                string contactAddress = contactSel.address + ((char)13) + contactSel.zipcode + " " + contactSel.city;
                Replace_Text("SLC_HostAndHostess_Address", contactAddress, m_AppWord);
                string infos = "";
                if (contactSel.cellphone != "")
                {
                    if (infos != "")
                    {
                        infos = infos + ((char)13);
                    }
                    infos = infos + contactSel.cellphone;
                }
                if (contactSel.email != "")
                {
                    if (infos != "")
                    {
                        infos = infos + ((char)13);
                    }
                    infos = infos + contactSel.email;
                }
                //Fix - for now no infos are needed to be displayed
                infos = "";
                Replace_Text("SLC_HostAndHostess_Infos", infos, m_AppWord);

                //Signature
                Replace_Text("SLC_Bill_Sign", m_Global_Handler.Resources_Handler.Get_Resources("TemplateSignature"), m_AppWord);

                //Bill informations
                Replace_Text("SLC_Bill_Subject", m_Mission.subject, m_AppWord);
                string dateGenerated = m_Global_Handler.Resources_Handler.Get_Resources("GenerationDate") + " : ";
                dateGenerated = dateGenerated + m_Global_Handler.DateAndTime_Handler.Treat_Date(DateTime.Today.ToString(), m_Global_Handler.Language_Handler);
                Replace_Text("SLC_Bill_DateGenerated", dateGenerated, m_AppWord);
                if (m_Type == BillType.Mission)
                {
                    Replace_Text("SLC_Bill_Validity", m_Global_Handler.Resources_Handler.Get_Resources("MissionValidity"), m_AppWord);
                }
                else if (m_Type == BillType.INVOICE)
                {
                    string MissionDateTermStr = m_Global_Handler.Resources_Handler.Get_Resources("TermDate");
                    MissionDateTermStr = MissionDateTermStr + " : " + m_Global_Handler.DateAndTime_Handler.Treat_Date(m_Mission.date_invoice_due, m_Global_Handler.Language_Handler);
                    Replace_Text("SLC_Bill_Validity", MissionDateTermStr, m_AppWord);
                }
                string globalDicount = "";
                if (m_Mission.discount != 0)
                {
                    globalDicount = m_Global_Handler.Resources_Handler.Get_Resources("GlobalDiscountGeneration1") + " " + m_Mission.discount + m_Global_Handler.Resources_Handler.Get_Resources("GlobalDiscountGeneration2");
                }
                Replace_Text("SLC_Bill_GlobalDiscount", globalDicount, m_AppWord);
                */
                return true;

            }
            catch (Exception error)
            {
                m_Global_Handler.Log_Handler.WriteMessage(System.Reflection.MethodBase.GetCurrentMethod().Name, error.StackTrace);

                //Close document and word
                Close_Word();

                return false;
            }
        }

        /// <summary>
        /// Finalize the word document
        /// <returns>True if everything went well, false otherwise</returns>
        /// </summary>
        private bool Finalize()
        {
            try
            {
                /*
                //Save doc
                string date = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "_" +
                    DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "-" + DateTime.Now.Second.ToString();
                object fileName = m_AppDir + "\\" + m_Mission.num_Mission + "_" + m_Mission.subject + "__" + date + ".docx";
                m_NewWordDoc.SaveAs(ref fileName, ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing,
                            ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing,
                            ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing, ref m_Missing,
                            ref m_Missing);

                //Save to pdf and open the pdf
                string prefix = "";
                if (m_Type == BillType.Mission)
                {
                    prefix = m_Global_Handler.Resources_Handler.Get_Resources("Mission");
                }
                else if (m_Type == BillType.INVOICE)
                {
                    prefix = m_Global_Handler.Resources_Handler.Get_Resources("Invoice");
                }
                else if (m_Type == BillType.CREDITNOTE)
                {
                    prefix = m_Global_Handler.Resources_Handler.Get_Resources("Credit");
                }
                string fileNamePDF = m_AppDir + "\\" + prefix + " - " + m_Mission.subject + ".pdf";
                if (m_GeneratedBillFileName != "")
                {
                    fileNamePDF = m_GeneratedBillFileName;
                }
                else
                {
                    m_GeneratedBillFileName = fileNamePDF;
                }
                m_NewWordDoc.ExportAsFixedFormat(fileNamePDF, WdExportFormat.wdExportFormatPDF);

                //Close document and word
                ((_Document)m_NewWordDoc).Close();
                ((_Application)m_AppWord).Quit();

                //Delete the word doc
                if (m_Global_Handler.File_Handler.File_Exists(fileName.ToString()) == true)
                {
                    m_Global_Handler.File_Handler.Delete_File(fileName.ToString());
                }
                */
                return true;
            }
            catch (Exception error)
            {
                m_Global_Handler.Log_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, error);

                //Close document and word
                Close_Word();

                return false;
            }
        }

        /// <summary>
        /// Close the word document
        /// </summary>
        private void Close_Word()
        {
            if (m_NewWordDoc != null)
            {
                ((_Document)m_NewWordDoc).Close();
            }
            if (m_AppWord != null)
            {
                ((_Application)m_AppWord).Quit();
            }
        }

        #endregion

    }
}