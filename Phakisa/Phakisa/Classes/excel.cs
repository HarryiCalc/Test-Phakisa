using System;
using System.IO;
using System.Collections;
using System.Threading; 
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Analysis = clsAnalysis;
using TB = clsTable;
using DB = clsDBase;
using Base = clsMain;
using General = clsGeneral;
using System.Data.OleDb;
using MetaReportRuntime;
using System.Windows.Forms;

namespace ATPMain
{
    /// <summary>
    /// Project:	Assay Trending Process - Controlled Spreadsheets project
    /// Author:		Vahe Karamian
    /// Date:		03/01/2005
    /// Version:	0.0
    /// </summary>
    public class VkExcel
    {
        clsBL.clsBL BusinessLanguage = new clsBL.clsBL();
        clsTable.clsTable TB = new clsTable.clsTable();
        clsGeneral.clsGeneral General = new clsGeneral.clsGeneral();
        clsMain.clsMain Base = new clsMain.clsMain();
        clsAnalysis.clsAnalysis Analysis = new clsAnalysis.clsAnalysis();

        private Excel.Application excelApp = null;
        private Excel.Workbook excelWorkbook = null;
        private Excel.Sheets excelSheets = null;
        private Excel.Worksheet excelWorksheet = null;

        private static object vk_missing = System.Reflection.Missing.Value;

        private static object vk_visible = true;
        private static object vk_false = false;
        private static object vk_true = true;

        private bool vk_app_visible = false;

        private object vk_filename;

        #region OPEN WORKBOOK VARIABLES
        private object vk_update_links = 0;
        private object vk_read_only = vk_true;
        private object vk_format = 1;
        private object vk_password = vk_missing;
        private object vk_write_res_password = vk_missing;
        private object vk_ignore_read_only_recommend = vk_true;
        private object vk_origin = vk_missing;
        private object vk_delimiter = vk_missing;
        private object vk_editable = vk_false;
        private object vk_notify = vk_false;
        private object vk_converter = vk_missing;
        private object vk_add_to_mru = vk_false;
        private object vk_local = vk_false;
        private object vk_corrupt_load = vk_false;
        #endregion

        #region CLOSE WORKBOOK VARIABLES
        private object vk_save_changes = vk_false;
        private object vk_route_workbook = vk_false;
        #endregion

        /// <summary>
        /// Vahe Karamian - 03/04/2005 - Excel Object Constructor.
        /// </summary>
        public VkExcel()
        {
            this.startExcel();
        }

        /// <summary>
        /// Vahe Karamian - 03/04/2005 - Excel Object Constructor
        /// visible is a parameter, either TRUE or FALSE, of type object.
        /// </summary>
        /// <param name="visible">Visible parameter, true for visible, false for non-visible</param>
        public VkExcel(bool visible)
        {
            this.vk_app_visible = visible;
            this.startExcel();
        }

        /// <summary>
        /// Vahe Karamian - 03/04/2005 - Start Excel Application
        /// </summary>
        #region START EXCEL
        private void startExcel()
        {
            if (this.excelApp == null)
            {
                this.excelApp = new Excel.Application();
            }

            // Make Excel Visible
            this.excelApp.Visible = this.vk_app_visible;
        }
        #endregion

        /// <summary>
        /// Vahe Karamian - 03/23/2005 - Kill the current Excel Process
        /// </summary>
        #region STOP EXCEL
        public void stopExcel()
        {
            if (this.excelApp != null)
            {
                Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");
                pProcess[0].Kill();
            }
        }
        #endregion

        /// <summary>
        /// Vahe Karamian - 03/09/2005 - Open File function for Excel 2003
        /// The following function will take in a filename, and a password
        /// associated, if needed, to open the file.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="password"></param>
        #region OPEN FILE FOR EXCEL
        public string OpenFile(string fileName, string password)
        {
            vk_filename = fileName;

            if (password.Length > 0)
            {
                vk_password = password;
            }

            try
            {
                // Open a workbook in Excel
                this.excelWorkbook = this.excelApp.Workbooks.Open(
                    fileName, vk_update_links, vk_read_only, vk_format, vk_password,
                    vk_write_res_password, vk_ignore_read_only_recommend, vk_origin,
                    vk_delimiter, vk_editable, vk_notify, vk_converter, vk_add_to_mru,
                    vk_local, vk_corrupt_load);
            }
            catch (Exception e)
            {
                this.CloseFile();
                return e.Message;
            }
            return "OK";
        }
        #endregion

        public void CloseFile()
        {
            excelWorkbook.Close(vk_save_changes, vk_filename, vk_route_workbook);
        }

        /// <summary>
        /// Vahe Karamian - 03/20/2005 - Get Excel Sheets
        /// Get the collection of sheets in the workbook
        /// </summary>
        #region GET EXCEL SHEETS
        public void GetExcelSheets()
        {
            if (this.excelWorkbook != null)
            {
                excelSheets = excelWorkbook.Worksheets;
            }
        }
        #endregion

        /// <summary>
        /// Vahe Karamian - 03/21/2005 - Find Excel ATP Worksheet
        /// Search for ATP worksheet, if found return TRUE
        /// </summary>
        /// <returns>bool</returns>
        #region FIND EXCEL ATP WORKSHEET
        public bool FindExcelWorksheet(string worksheetName)
        {
            bool ATP_SHEET_FOUND = false;

            if (this.excelSheets != null)
            {
                // Step thru the worksheet collection and see if ATP sheet is
                // available. If found return true;
                for (int i = 1; i <= this.excelSheets.Count; i++)
                {
                    this.excelWorksheet = (Excel.Worksheet)excelSheets.get_Item((object)i);
                    if (this.excelWorksheet.Name.Equals(worksheetName))
                    {
                        this.excelWorksheet.Activate();
                        ATP_SHEET_FOUND = true;
                        return ATP_SHEET_FOUND;
                    }
                }
            }
            return ATP_SHEET_FOUND;
        }
        #endregion

        /// <summary>
        /// Vahe Karamian - 03/22/2005 - Get Range from Worksheet
        /// Return content of range from the selected range
        /// </summary>
        /// <param name="range">Range parameter: Example, GetRange("A1:D10")</param>
        #region GET RANGE
        public string[] GetRange(string range)
        {
            Excel.Range workingRangeCells = excelWorksheet.get_Range(range, Type.Missing);
            //workingRangeCells.Select();
            System.Array array = (System.Array)workingRangeCells.Cells.Value2;
            string[] arrayS = this.ConvertToStringArray(array);

            return arrayS;
        }
        #endregion

        /// <summary>
        /// Vahe Karamian - 03/22/2005 - Convert To String Array
        /// Convert System.Array into string[]
        /// </summary>
        /// <param name="values">Values from range object</param>
        /// <returns>String[]</returns>
        #region CONVERT TO STRING ARRAY
        private string[] ConvertToStringArray(System.Array values)
        {
            string[] newArray = new string[values.Length];

            int index = 0;
            for (int i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++)
            {
                for (int j = values.GetLowerBound(1); j <= values.GetUpperBound(1); j++)
                {
                    if (values.GetValue(i, j) == null)
                    {
                        newArray[index] = "";
                    }
                    else
                    {
                        newArray[index] = (string)values.GetValue(i, j).ToString();
                    }
                    index++;
                }
            }
            return newArray;
        }
        #endregion

        public void SaveFile(string period,string environment)
        {
           // MessageBox.Show("in save file");

            this.excelWorkbook.SaveAs(

                "C:\\iCalc\\Harmony\\Phakisa\\" + environment + "\\Data\\adteam_" + period + ".xls",

                Excel.XlFileFormat.xlWorkbookNormal,

                "",

                vk_write_res_password,

                vk_read_only,

                null,

                 Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,

                 null,

                 vk_add_to_mru,

                 null, null, vk_local);
        }
        public string OpenFileShow(string fileName, string password)
        {

            vk_filename = fileName;
            //MessageBox.Show(fileName);


            if (password.Length > 0)
            {
                vk_password = password;

            }
            // MessageBox.Show("net voor try");
            try
            {
                //MessageBox.Show("gaan nou oop maak");
                // Open a workbook in Excel
                // MessageBox.Show(fileName);
                // MessageBox.Show(vk_format.ToString());



                this.excelWorkbook = this.excelApp.Workbooks.Open(
                    fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                excelApp.Visible = true;

            }
            catch (Exception e)
            {
                // MessageBox.Show(e.ToString());
                this.CloseFile();
                return e.Message;
            }
            return "OK";

        }
    }
}
