using System;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading;
using System.Windows.Forms;

namespace Phakisa
{
    class clsShared
    {
        #region Declarations

        #endregion

        #region Properties

        #endregion

        #region Methods

        //===============================================================================
        public void createViews(clsMain.clsMain main)
        {
            //A threat is started to extract views on calculated tables.  
            Thread t = new Thread(ExtractViews);   // Kick off a new thread
            t.Start(main);
        }

        static void ExtractViews(object main)
        {

            clsMain.clsMain M = (clsMain.clsMain)main;
            M.extractViews();

            //Base.extractViews();

        }
        //===============================================================================
        public void copyFormulas(clsMain.clsMain main)
        {
            //A threat that will copy the formulas if not exist on Analysis.  
            Thread t = new Thread(extractFormulas);   // Kick off a new thread
            t.Start(main);

            // main.CopyFormulas();
        }

        static void extractFormulas(object main)
        {

            clsMain.clsMain M = (clsMain.clsMain)main;
            M.getHistoryAndCopy();

        }
        //===============================================================================
        public void extractPrimaryKeys(clsMain.clsMain main)
        {
            //A threat is started to extract the primary keys of selected tables.
            //The primary keys are stored in clsMain.
            //When the user select one of the selected tables tab on the front-end,
            //the list are passed from clsMain into the primary keys list.
            //No extracts are done to the databases and that makes the audit table fast.
            //ExtractKeys(main);
            Thread t = new Thread(ExtractKeys);   // Kick off a new thread
            t.Start(main);
        }

        static void ExtractKeys(object main)
        {

            clsMain.clsMain M = (clsMain.clsMain)main;
            M.extractPrimaryKey();

        }

        //===============================================================================
        public void extractListOfTableNames(clsMain.clsMain main)
        {
            Thread t = new Thread(ExtractTableNames);   // Kick off a new thread
            t.Start(main);
        }

        static void ExtractTableNames(object main)
        {

            clsMain.clsMain M = (clsMain.clsMain)main;
            M.extractDBTableNames();

        }

        //===============================================================================
        public void evaluateDataTable(clsMain.clsMain main, string Tablename)
        {
            main.DataTableToExtract = Tablename;
            ExtractDataTable(main);
            Thread t = new Thread(ExtractDataTable);   // Kick off a new thread
            t.Start(main);
        }

        static void ExtractDataTable(object main)
        {

            clsMain.clsMain M = (clsMain.clsMain)main;
            M.ExtractDataTable();

        }

        //===============================================================================

        public Boolean CreateUDLFile(string FileName, OleDbConnectionStringBuilder builder)
        {
            try
            {

                string conn = Convert.ToString(builder);
                MSDASC.DataLinksClass aaa = new MSDASC.DataLinksClass();
                aaa.WriteStringToStorage(FileName, conn, 1);
                return true;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("UDL error message: " + ex.Message);
                return false;
            }
        }


        #endregion





        public void metareportAutoWParameter(string parameterSet, string parm1, string filename, string reportName,
                                                string MetaReportCode, string ReportPath)
        {

            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(MetaReportCode);
            mm.ProjectsPath = ReportPath;
            int repNo = mm.get_OpenReport(reportName);
            mm.SetStringParameter(parameterSet, parm1, repNo);

            //mm.PrintToFile(repNo, filename, MetaReportRuntime.MRFileFormat.ffPDF);
            //mm.CloseReport(repNo);
            //System.Diagnostics.Process.Start(filename);

            mm.PrintToFile(repNo, ReportPath.Replace(ReportPath.Substring(0, 1).ToString().Trim(), "C") + "\\PDFS\\" + filename, MetaReportRuntime.MRFileFormat.ffPDF);
            mm.CloseReport(repNo);
            System.Diagnostics.Process.Start(ReportPath.Replace(ReportPath.Substring(0, 1).ToString().Trim(), "C") + "\\PDFS\\" + filename);

            //om in Metareport oop te maak comment
            //mm.PrintToFile(repNo, filename, MetaReportRuntime.MRFileFormat.ffPDF); en
            //System.Diagnostics.Process.Start(filename);

            //uit en uncomment 

            //mm.StartReportWithParameters(repNo, true);
        }

        

        public void metareportAutoW3Parameter(string parameterSet1, string parameterSet2, string parameterSet3,
                                              string parm1, string parm2, string parm3, string filename,
                                              string reportName, string MetaReportCode, string ReportPath)
        {

            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(MetaReportCode);
            mm.ProjectsPath = ReportPath;
            int repNo = mm.get_OpenReport(reportName);
            mm.SetStringParameter(parameterSet1, parm1.ToString().Trim(), repNo);
            mm.SetStringParameter(parameterSet2, parm2.ToString().Trim(), repNo);
            mm.SetStringParameter(parameterSet3, parm3.ToString().Trim(), repNo);
            mm.StartReportWithParameters(repNo, true);

            mm.PrintToFile(repNo, ReportPath + "\\PDFS\\" + filename, MetaReportRuntime.MRFileFormat.ffPDF);
            mm.CloseReport(repNo);
            System.Diagnostics.Process.Start(ReportPath + "\\PDFS\\" + filename);



            //om in Metareport oop te maak comment
            //mm.PrintToFile(repNo, filename, MetaReportRuntime.MRFileFormat.ffPDF); en
            //System.Diagnostics.Process.Start(filename);

            //uit en uncomment 

            //mm.StartReportWithParameters(repNo, true);
        }

        public void metareportAutoW4Parameter(string parameterSet1, string parameterSet2, string parameterSet3, string parameterSet4,
                                             string parm1, string parm2, string parm3, string parm4, string filename,
                                             string reportName, string MetaReportCode, string ReportPath)
        {

            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(MetaReportCode);
            mm.ProjectsPath = ReportPath;
            int repNo = mm.get_OpenReport(reportName);
            mm.SetStringParameter(parameterSet1, parm1.ToString().Trim(), repNo);
            mm.SetStringParameter(parameterSet2, parm2.ToString().Trim(), repNo);
            mm.SetStringParameter(parameterSet3, parm3.ToString().Trim(), repNo);
            mm.SetStringParameter(parameterSet4, parm4.ToString().Trim(), repNo);
            mm.PrintToFile(repNo, ReportPath + "\\PDFS\\" + filename, MetaReportRuntime.MRFileFormat.ffPDF);
            mm.CloseReport(repNo);
            System.Diagnostics.Process.Start(ReportPath + "\\PDFS\\" + filename);


            //om in Metareport oop te maak comment
            //mm.PrintToFile(repNo, filename, MetaReportRuntime.MRFileFormat.ffPDF); en
            //System.Diagnostics.Process.Start(filename);

            //uit en uncomment 

            //mm.StartReportWithParameters(repNo, true);

        }

        public int checkLockCalendarProcesses(DataTable Status, string Period)
        {

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "N"
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          where locks.Field<string>("PROCESS").TrimEnd() == "tabCalendar"
                                          where locks.Field<string>("PERIOD").TrimEnd() == Period
                                          select locks;

            try
            {
                int intcount = query1.Count<DataRow>();

                return intcount;
            }
            catch
            {
                MessageBox.Show("Error in checkLockCalendarProcess.");
                return 0;
            }


        }

    }
}
