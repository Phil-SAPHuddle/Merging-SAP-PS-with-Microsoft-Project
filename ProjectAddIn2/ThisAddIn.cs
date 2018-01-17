using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Runtime.InteropServices;
using System.Data;


namespace ProjectAddIn2
{
    //  MS VSTO AddIn for MS Project
    //    Provides SAP OpenPS functionality as shareware
    //      Notes:
    //        1. Only one SAP Project can be open at a time.
    //        2. The referenced CSAPData.dll is compiled in VS 2003
    //


    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface AddInInterface
    {
    // Published interfaces for use in MSP VBA
    //

        void SelectProject();
        void RefreshProject();
        Boolean SaveToSAP();
        void ShowSAPData_Debug();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public partial class ThisAddIn : AddInInterface
    {
        private frmSAPData frmSAP;
        static MSProject.Application oApplication = new Microsoft.Office.Interop.MSProject.Application();
        static CProject oProject;  // Populated when an SAP project is opened
        static CSAPInterface oSAPInterface;
        private ThisAddIn objAddIn;
        private frmProjectSelection frmPrjSelection;
        
        protected override object RequestComAddInAutomationService()
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                if (objAddIn == null)
                    objAddIn = new ThisAddIn();

                Cursor.Current = Cursors.Default;
                return objAddIn;
            }
            catch (SystemException ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show(ex.Message, "Error- InvalidCastException.");
                return "Error in thsAddIn:RequstComAddInAutomationService";
            }

        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            oApplication = Application;
            // Set oProject as the current ActiveProject
            oProject = new CProject(ref Application);
            // Pass the project in for use in CSAPInterface
            oSAPInterface = new CSAPInterface(ref oProject);

            oApplication.StatusBar = "SAP AddIn loaded successfully.";

            Cursor.Current = Cursors.Default;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

            // Clean up
            if (frmSAP != null)
            {
                frmSAP.Close();
                frmSAP = null;
            }
            //Disconnect from our SAP DB
            if ( oSAPInterface != null )
            {
                oSAPInterface.Logout();
                oSAPInterface = null;
            }

            oApplication = null;

            // Release the AddIn reference
            if (objAddIn != null)
            {
                objAddIn.Dispose();
                objAddIn = null;
            }
        }


         public void SelectProject()  
        // Provides a form with a grid populated with
        //   the collection of SAP PS Projects.
        //   This is a starting point for working with
        //   a PS Project.
        //      - On the form OK button event the Project data is loaded.
        //      - Only one SAP Project can be opened at a time.
        //--------------------------------------------------------------------------------------------
        {
            Cursor.Current = Cursors.WaitCursor;

            oApplication.StatusBar = oSAPInterface.Login();
            // Close any open project ( only one SAP project is opened at a time )
            oProject.DeleteTasks();
            if(oProject.Name != "SAPProject")  // No project has been saved. This is the global template.
                oApplication.FileClose(Microsoft.Office.Interop.MSProject.PjSaveType.pjPromptSave, false);
            frmPrjSelection = new frmProjectSelection(ref oApplication, ref oSAPInterface);
            frmPrjSelection.Show();

            Cursor.Current = Cursors.Default;
        }

        public void RefreshProject()
        {
            oApplication.StatusBar = oSAPInterface.RefreshPrjData();
        }

        public Boolean SaveToSAP()
        {
            //Save MSP changes to SAP
            oApplication.StatusBar = oSAPInterface.MaintainSAPProject();
            return true;
        }

        public void ShowSAPData_Debug()
        {
            frmSAP = new frmSAPData(oSAPInterface, oProject.ProjectID);
            frmSAP.Show();

        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        private double DateDiff(System.DateTime startdate, System.DateTime enddate)
        {
            double diff = 0;
            System.TimeSpan ts = new System.TimeSpan(startdate.Ticks - enddate.Ticks);
            diff = Convert.ToDouble(ts.TotalDays);
            return diff;
        }
        #endregion
    }
}




