using System;
using COMException = System.Runtime.InteropServices.COMException;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MSProject = Microsoft.Office.Interop.MSProject;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;


namespace ProjectAddIn2
{
    public class CProject 
    {
        // Represents the encapsulated integrated SAP <-> MSP project object
        private MSProject.Application   rApplication;
        private MSProject.Project       rProject;
        private MSProject.Tasks         rTasks;
        private String                  sProjectID;

        // Constructor
        public CProject(ref MSProject.Application Application)
        {
            // sGlobalTemplate is our global MS Project objects
            string sGlobalTemplate = Properties.Settings.Default.TestDir + "SAPProject.mpt";

            rApplication = Application;

            try
            {
                //Load our GlobalTemplate
                //rApplication.FileNew(false, @sGlobalTemplate, false, false);
                rApplication.FileNew(false, @sGlobalTemplate, false, false);
            }
            catch (IOException e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //Assign reference from standard MSProject to our own
            rProject = rApplication.ActiveProject;
            rTasks = rProject.Tasks;
        }

        public MSProject.Task Find(String sType, String sID, String sSubID)
        { // This is not optimized.  Preferably we would extend the rProject.Tasks
            // collection to include IEnumerable with lamba Find support

            foreach (MSProject.Task oMSPTask in rTasks)
            {
                switch (sType)
                {
                    case TaskType.WBS:
                        if(sID == oMSPTask.WBS)
                            return oMSPTask;
                        break;
                    case TaskType.Network:
                        if(sID == oMSPTask.Text30)
                            return oMSPTask;
                        break;
                    case TaskType.Activity:
                        if (sID == oMSPTask.Text30 && sSubID == oMSPTask.Text16)
                            return oMSPTask;
                        break;
                } // switch
            } // foreach
            // default
            return null;
          
        }

        public Boolean OpenMSPProject(String sProjectID)
        {  // Open an MS Project .mpp file for merging with SAP Data,
           //   If the file open fails then SAP data is loaded signifying
           //   a new project.

            // TestDir stores the merger working .mpp file 
            String sTemp = Properties.Settings.Default.TestDir + sProjectID;

            try
            {
                //Load the existing .mpp file 
                rApplication.FileOpen(sTemp, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                       Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                       Missing.Value, Microsoft.Office.Interop.MSProject.PjPoolOpen.pjDoNotOpenPool,
                                       Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                //Assign reference from opened MSProject to our own
                rProject = rApplication.ActiveProject;
                rTasks = rProject.Tasks;
            }
            catch (COMException ex)
            {
                //FileOpen failed
                return false;
            }

            return true;
        }

        public Boolean DeleteTasks()
        {
            foreach (MSProject.Task oMSPTask in rProject.Tasks)
                oMSPTask.Delete();

            return true;
        }

/////////////////////////////////////////////////////////////////////////////////////////
        // Accessor functions for Project attributes follow.
        public MSProject.Project Project
        {
            // We set the project instance only in the constructor
            get { return rProject; }
        }

        public MSProject.Tasks Tasks
        {
            get { return rTasks; }
        }

        public String Title
        {
            get { return rProject.Title.ToString(); }
            set { rProject.Title = value.ToString(); }
        }

        public string ProjectID
        { // MSP Project Name can not be changed so we 
            //   use a local variable and the CreateProjectFromTemplate
            //   sub to create a project with the SAP ID.
            // Note there is redunduncy with the CSAPInterface:ProjectID assesor 

            get { return sProjectID; }
            set { sProjectID = value; }
        }

        public String Name
        {
            get { return rProject.Name; }
            set { rProject.Name = value; }
        }


        public DateTime StartDate
        {
            get { return Convert.ToDateTime(rProject.ProjectStart); }
            set { rProject.ProjectStart = value; }
        }

        public DateTime FinishDate
        {
            get { return Convert.ToDateTime(rProject.ProjectFinish); }
            set { rProject.ProjectFinish = value; }
        }

        public String ProjectProfile
        {
            get { return rProject.Text1; }
            set { rProject.Text1 = value; }
        }

        public Int16 Priority   // To Do load from SAP
        {
            get { return Convert.ToInt16(rProject.Priority); }
            set { rProject.Priority = value; }
        }

        public void AddSummaryTask(String Description, Int16 Outlinelevel)
        {
            rTasks.Add(Description, Outlinelevel);
        }

    }
}


