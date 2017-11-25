using Microsoft.Office.Interop.Excel;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.Windows;

namespace GeneralClasses
{

    /// <summary>
    /// Class ExcelGeneration containing all the tools to generate a excel statement from a list of bills
    /// </summary>
    public class ExcelGeneration
    {

        #region Variables

        /// <summary>
        /// Global handler
        /// </summary>
        private static Handlers m_Global_Handler = new Handlers();

        /// <summary>
        /// Error handler
        /// </summary>
        private static Log m_Error_Handler = new Log();

        /// <summary>
        /// Bill to treat
        /// </summary>
        private static Billing m_Bill = new Billing();

        #endregion

        #region Functions

        /// <summary>
        /// Transform a list to an excel file
        /// <param name="_Filename">Excel filename</param>
        /// <param name="_Lines">List of bills</param>
        /// <returns>True if everything went well, false otherwise</returns>
        /// </summary>
        public int Generate_ExcelStatement(string _Filename, List<string> _Lines)
        {
            try
            {
                //Start excel
                Microsoft.Office.Interop.Excel.Application excapp = new Microsoft.Office.Interop.Excel.Application();
                if (excapp == null)
                {
                    MessageBox.Show(m_Global_Handler.Resources_Handler.Get_Resources("MicrosoftExcelNotFound"),
                        m_Global_Handler.Resources_Handler.Get_Resources("MicrosoftExcelNotFoundCaption"),
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return -1;
                }
                excapp.Visible = false;

                //Create a blank workbook
                var workbook = excapp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                // Pull in all the cells of the worksheet
                Range cells = workbook.Worksheets[1].Cells;
                // set each cell's format to Text
                cells.NumberFormat = "@";
                // reset horizontal alignment to the right
                cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                //Get sheet
                var sheet = (Worksheet)workbook.Sheets[1]; //indexing starts from 1

                //Write the list
                int counterLine = 1;
                foreach (string item in _Lines)
                {
                    string[] values = item.Split('\t');
                    int counterColumn = 1;
                    foreach (string value in values)
                    {
                        sheet.Cells[counterLine, counterColumn] = value;
                        ++counterColumn;
                    }
                    ++counterLine;
                }
                //AutoFit columns
                sheet.Columns.AutoFit();

                //Save
                workbook.SaveAs(_Filename);

                //Close
                workbook.Close(0);
                excapp.Quit();

                //Everything OK
                return 0;
            }
            catch (Exception exception)
            {
                m_Error_Handler.WriteMessage(System.Reflection.MethodBase.GetCurrentMethod().Name, exception.StackTrace);
                return -1;
            }
        }

        #endregion
    }
}