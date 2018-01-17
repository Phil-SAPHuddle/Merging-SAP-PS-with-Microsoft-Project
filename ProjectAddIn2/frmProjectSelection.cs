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
    public partial class frmProjectSelection : Form
    {
        // Provides users a selection grid for opening an SAP Project
        //
        private MSProject.Application rApplication;
        private CSAPInterface rSAPInterface;
        private DataTable oProjDefs = new DataTable();

        //public frmProjectSelection(ref MSProject.Application oApplication, ref CSAPInterface oSAPInterface)
        public frmProjectSelection(ref MSProject.Application oApplication, ref CSAPInterface oSAPInterface)
        {
            InitializeComponent();
            // Set reference to MS Project Application object
            rApplication = oApplication;  //Used to set status bar from this form
            // Set reference to the CSAPInterface
            rSAPInterface = oSAPInterface;
            // Get the collection of SAP Project IDs to display
            oSAPInterface.GetProjDefList(ref oProjDefs);
            dataGridPrjDefs.DataSource = oProjDefs;
        }

         private void btnOK_Click(object sender, EventArgs e)
        {
            if (dataGridPrjDefs.CurrentRow.Selected == true)
            {
                String sProjectID; 
                Cursor.Current = Cursors.WaitCursor;

                // Retrieve the selected project ID 
                sProjectID = dataGridPrjDefs.CurrentRow.Cells[0].Value.ToString();
                rApplication.StatusBar = rSAPInterface.DisplayProject(sProjectID);

                Cursor.Current = Cursors.Default;
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            rSAPInterface = null;
            this.Close();
        }


    }
}
