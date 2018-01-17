using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace ProjectAddIn2
{
    public partial class frmSAPData : Form
    {
        // Provides SAP data analysis and debugging support by displaying 
        //   the specific SAP BAPI returned data
        //   Notes:
        //     1. SAP BAPI object data is loaded into MS DataTable objects
        //
        //


        private DataTable oProjectDetail;
        private DataTable oNtwkActy;
        private DataTable oMsgs;
        private DataTable oWBSELEMEXP;
        private DataTable oWBSHRCY;
        private DataTable oWBSMLST;
        private DataTable oTaskHier;
        private CSAPInterface rSAPInterface;
        public frmSAPData(CSAPInterface oSAPInterface, String sProjectID)
        {
            // Set our reference to CSAPInterface
            rSAPInterface = oSAPInterface;
            InitializeComponent();
            InitializeDataTables();
            // Set the selected project ID
            sProjectDef.Text = sProjectID;
            DisplayData();
        }

        private void InitializeDataTables()
        {
            oProjectDetail     = new DataTable();
            oNtwkActy          = new DataTable();
            oMsgs              = new DataTable();
            oWBSELEMEXP        = new DataTable();
            oWBSHRCY           = new DataTable();
            oWBSMLST           = new DataTable(); 
            oTaskHier          = new DataTable();
        }

        public void SetStatusText(String sInfo)
        {
            lblStatustext.Text = sInfo;
        }

        public Boolean DisplayData()
        {
            Cursor.Current = Cursors.WaitCursor;

            lblStatustext.Text = rSAPInterface.ShowSAPData_ProjectDef(ref lstBoxProjectDef);
            lblStatustext.Text = GetWBSData();
            lblStatustext.Text = GetMSPTaskHierarchy();
            Cursor.Current = Cursors.Default;
            return true;

        }

        public String GetWBSData()
        {
            String sMsg;

            sMsg = rSAPInterface.ShowSAPData_WBSData(ref oNtwkActy,
                                                     ref oMsgs,
                                                     ref oWBSELEMEXP,
                                                     ref oWBSHRCY,
                                                     ref oWBSMLST);

            dataGrid1.DataSource = oWBSELEMEXP;
            dataGrid2.DataSource = oNtwkActy;
            dataGrid3.DataSource = oWBSHRCY;
            dataGrid4.DataSource = oWBSMLST;
            dataGrid5.DataSource = oMsgs;

            return sMsg;
        }       

         public String GetMSPTaskHierarchy()
         {
             String sMsg;

             oTaskHier.Clear();

             sMsg = rSAPInterface.ShowSAPData_TaskHierarchy(ref oTaskHier);

             dataGrid13.DataSource = oTaskHier;
             return sMsg;
         }

         public void SetTab(int iTab)
         {
             tabCtrl.SelectTab(iTab);
         }

     }
}
