using System;
using System.Data;
using COMException = System.Runtime.InteropServices.COMException;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MSProject = Microsoft.Office.Interop.MSProject;


namespace ProjectAddIn2
{
    public class CSAPInterface
    {
        // Encapsulates the data interfacing functionality between SAP and MSP
        //    Notes:
        //      1. Builds on the default SAP data proxy generated object (ref CSAPData)
        //      2. SAP BAPI related data structures are defined and populated here
        //      3. Includes the logic for converting an SAP project to an MSP project in the BuildSAPProject sub
        //      4.  
        //               
        //


        // const in function calls
        public const int iMaxRows = 1000;
        //////private string gProjectID;
        private static Boolean bSaveMethodAdded = false;      //Used in sub MaintainSAPActivity(..)
        // Reference to our custom version of the SAPProxy class
        private CSAPData.CSAPPrj oSAP = new CSAPData.CSAPPrj();

        // BAPI returns
        private CSAPData.BAPIRETURN1                       stBAPIReturn1 = new CSAPData.BAPIRETURN1();
        private CSAPData.BAPI_METH_MESSAGETable                   arMsgs = new CSAPData.BAPI_METH_MESSAGETable();
        // GetPrjDefList
        private CSAPData.BAPIPREXPTable                        arProjDef = new CSAPData.BAPIPREXPTable();
        private CSAPData.BAPI_2002_DESCR_RANGETable      arSAPDescrRange = new CSAPData.BAPI_2002_DESCR_RANGETable();
        private CSAPData.BAPI_2002_PD_RANGETable        arsAPProjIDRange = new CSAPData.BAPI_2002_PD_RANGETable();
        // GetPrjDetail
        private CSAPData.BAPI_BUS2001_DETAIL             stProjectDetail = new CSAPData.BAPI_BUS2001_DETAIL();
        // GetWBSData
        private CSAPData.BAPI_PROJECT_DEFINITION_EX         stProjectDef = new CSAPData.BAPI_PROJECT_DEFINITION_EX();
        private CSAPData.BAPI_NETWORK_ACTIVITY_EXPTable       arNtwkActy = new CSAPData.BAPI_NETWORK_ACTIVITY_EXPTable();
        private CSAPData.BAPI_WBS_ELEMENT_EXPTable          arWBSELEMEXP = new CSAPData.BAPI_WBS_ELEMENT_EXPTable();
        private CSAPData.BAPI_WBS_HIERARCHIETable              arWBSHRCY = new CSAPData.BAPI_WBS_HIERARCHIETable();
        private CSAPData.BAPI_WBS_MILESTONE_EXPTable           arWBSMLST = new CSAPData.BAPI_WBS_MILESTONE_EXPTable();
        private CSAPData.BAPI_WBS_ELEMENTSTable          arWBSELEMSELECT = new CSAPData.BAPI_WBS_ELEMENTSTable();

        private CTaskHierarchy oTaskCollection = new CTaskHierarchy();
        private CProject rProject;  // reference to our CProject object
        private Boolean bLoggedIn = new Boolean();

        public CSAPInterface(ref CProject oProject) 
        {
            rProject = oProject;
        }

        public String Login()
        {
            String sConn = "CLIENT="  + Properties.Settings.Default.Client +
                           " USER="   + Properties.Settings.Default.User +
                           " PASSWD=" + Properties.Settings.Default.Password +
                           " LANG="   + Properties.Settings.Default.Lang +
                           " ASHOST=" + Properties.Settings.Default.ASHOST;

            try
            {
                if (bLoggedIn != true)
                    bLoggedIn = oSAP.LogIn(sConn);
                if (bLoggedIn == true)
                    return "Login successfull";
                else
                    return "Login failed";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SAP connect error.");
                return "Error in CSAPInterface:Login";
            }
        }

        public String Logout()
        {
            if (bLoggedIn == true)
            {
                oSAP.LogOut();
            }
            bLoggedIn = false;
            return "Logged out.";
        }

        public String GetProjectDef(string sProjectDef)
        {
          
            try
            {
                oSAP.GetProjectDefData("EN", 
                                        sProjectDef,
                                        ref stProjectDetail);

                // Load the CProject object rProject with SAP Project Def attributes.
                rProject.Title =  stProjectDetail.Description;
                rProject.FinishDate = FromSAPDateConversion(stProjectDetail.Finish.ToString());
                rProject.Name = stProjectDetail.Project_Definition;  // Sets MSP Project # also
                rProject.ProjectProfile = stProjectDetail.Project_Profile;
                rProject.StartDate = FromSAPDateConversion(stProjectDetail.Start.ToString());

                return "Project Detail data loaded.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in GetPrjDefData.");
                return "Error in CSAPInterface:GetPrjDef";
            }
        }

        public String GetProjDefList(ref DataTable oProjDefs)
        // Returns a DataTable object loaded with all SAP PS Projects
        //   for displaying in a selection grid.
        //   Calls CSAPData.Bapi_Projectdef_Getlist
        {
            try
            {
                oSAP.ProjectDef_GetList(iMaxRows, ref arProjDef, ref arSAPDescrRange, ref arsAPProjIDRange);
                oProjDefs = arProjDef.ToADODataTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in GetProjDefList.");
                return "Failed project listing load.";
            }

            return "Listing of Projects loaded.";
        }

        public String GetWBSData(string sProjectDef,
                                 string sWBS)
        {
            Boolean bReturn = new Boolean();
            //   Get the full project			
            String sWithActivities = "X";
            String sWithMilestones = "X";
            String sWithSubtree = "";

            //	 Get a specific WBS and subtree	
            if (sWBS.Length > 0)
            {
                if (sProjectDef.Length > 0)
                {
                    MessageBox.Show("Error; Project Def and WBS can not be entered at the same time.");
                    return "Error; Project Def and WBS can not be entered at the same time.";
                }

                sWithSubtree = "X";
                CSAPData.BAPI_WBS_ELEMENTS sWBSSelect = new CSAPData.BAPI_WBS_ELEMENTS();
                sWBSSelect.Wbs_Element = sWBS;
                arWBSELEMSELECT.Add(sWBSSelect);
            }

            try
            {
                bReturn = oSAP.GetPrjWBSData(sProjectDef,
                                             sWithActivities,
                                             sWithMilestones,
                                             sWithSubtree,
                                             out stProjectDef,
                                             out stBAPIReturn1,
                                             ref arNtwkActy,
                                             ref arMsgs,
                                             ref arWBSELEMEXP,
                                             ref arWBSHRCY,
                                             ref arWBSMLST,
                                             ref arWBSELEMSELECT);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in GetPrjWBSData.");
                return "Error in CSAPInterface:GetWBSData.";
            }

            if (bReturn == true)
                return "WBS & children data loaded.";
            else
            {
                if (arMsgs != null)
                {

                    MessageBox.Show("Error in CSAPInterface:GetWBSData");
                    frmSAPData frmSAP = new frmSAPData(this, rProject.ProjectID);
                    frmSAP.Show();
                    frmSAP.SetTab(4);  // set to error message tab
                }
            }


            return "WBS & children data loaded.";
        }

        public String GetMSPTaskHierarchy()
        {

            try
            {
                oTaskCollection.GetTaskHierarchy(arWBSHRCY, arNtwkActy);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in GetMSPTaskHierarchy.");
                return "Error in CSAPInterface:GetMSPTaskHierarchy.";
            }

            return "MSP Task Hierarchy loaded.";

        }

        public void BuildMSPWBSElement(ref MSProject.Task oTask, int i, int j, short nLevel)
        { 
            // Many of the attributes of a summary task are derived by MS Project
            //   from the subordinate tasks

            // Use this to discuss creating and setting custom fields, see the notes, 
            // !Use intellisense to get the ID
            oTask.SetField(Microsoft.Office.Interop.MSProject.PjField.pjTaskOutlineCode1, "Yes");
            //oTask.Summary = true;
            oTask.Name = arWBSELEMEXP[j].Description;
            //oTask. = arWBSELEMEXP[j].Project_Definition;
            oTask.Text15 = arWBSELEMEXP[j].Description;
            oTask.Text16 = TaskType.WBS;
            oTask.Text29 = arWBSELEMEXP[j].Wbs_Element;
            try
            {  // Outline level of summary tasks are specified by
               //   their parent relationship
                oTask.OutlineLevel = nLevel;
            }
            catch (COMException ex)
            { ; }

            try
            {
                //If this is a summary task we cannot change the date value
                //  MS calculates; ex. A top level start date is the earliest start date
                //  of its subordinate tasks
                oTask.Start = (DateTime)FromSAPDateConversion(arWBSELEMEXP[j].Wbs_Basic_Start_Date);
                oTask.Finish = FromSAPDateConversion(arWBSELEMEXP[j].Wbs_Basic_Finish_Date);
                oTask.ActualStart = FromSAPDateConversion(arWBSELEMEXP[j].Wbs_Actual_Start_Date);
                oTask.ActualFinish = FromSAPDateConversion(arWBSELEMEXP[j].Wbs_Actual_Finish_Date);
                //oTask.OutlineLevel = nLevel;
            }
            catch (COMException ex)
            {
                //MessageBox.Show(ex.Message, "Error in BuildMSPWBSElement:oTaskStart");
                ;
            }

        }

        public void BuildMSPNetwork(ref MSProject.Task oTask, int i, int j, short nLevel)
        {  // On an initial project build we create a Network with the first Activity
            //   With the MergeSAPMSPProject sub we refresh the Network only


            //Parent object
            oTask.SetField(Microsoft.Office.Interop.MSProject.PjField.pjTaskOutlineCode1, "Yes");
            oTask.Name = TaskType.Network;  // arNtwkActy[j].Description;  need to get Network description
            try
            {  // Outline level of parent tasks are specified by
                //   their parent relationship
                oTask.OutlineLevel = nLevel;
            }
            catch (COMException ex)
            { ; }
            try
            {
                //If this is a parent task we cannot change the date value
                //  MS calculates; ex. A top level start date is the earliest start date
                //  of its subordinate tasks
                oTask.ActualStart = FromSAPDateConversion(arNtwkActy[j].Actual_Start_Date);
                oTask.ActualFinish = FromSAPDateConversion(arNtwkActy[j].Actual_Finish_Date);
            }
            catch (COMException ex)
            { ; }

            oTask.Text15 = arNtwkActy[j].Description;
            oTask.Text16 = TaskType.Network;
            //oTask.Text27 = arNtwkActy[j]. ... MRP Controller
            // Sched_type
            // Profile
            oTask.Text29 = arNtwkActy[j].Wbs_Element;  //My addition for checking
            oTask.Text30 = arNtwkActy[j].Network;

        }

        public void BuildMSPTask(ref MSProject.Task oTask, int i, int j, short nLevel)
        {
            oTask.SetField(Microsoft.Office.Interop.MSProject.PjField.pjTaskOutlineCode1, "Yes");
            oTask.Name = arNtwkActy[j].Description;
            try
            {  // ??????
                oTask.OutlineLevel = (short)(nLevel);  
            }
            catch (COMException ex)
            { ; }
            try
            {
                //?????
                oTask.ActualStart = FromSAPDateConversion(arNtwkActy[j].Actual_Start_Date);
                oTask.ActualFinish = FromSAPDateConversion(arNtwkActy[j].Actual_Finish_Date);
            }
            catch (COMException ex)
            { ; }

            oTask.Text1 = arNtwkActy[j].Activity;
            oTask.Text15 = arNtwkActy[j].Description;
            oTask.Text16 = TaskType.Activity;
            //oTask.Text21 = arNtwkActy[j]... Control Key  ..
            // Resource Text21 = Work Center
            // Resource Text20 = Plant
            oTask.Text29 = arNtwkActy[j].Wbs_Element;
            oTask.Text30 = arNtwkActy[j].Network;

        }

        public void MaintainSAPActivity(MSProject.Task oMSPTask, String sMethod, int iRefNum, 
                                        CSAPData.BAPI_METHOD_PROJECTTable arPrjMethod,
                                        CSAPData.BAPI_NETWORKTable arNtwks,
                                        CSAPData.BAPI_NETWORK_ACTIVITYTable arNtwkActivities,
                                        CSAPData.BAPI_NETWORK_ACTIVITY_UPTable arNtwkActivitiesUpDate )                   
        {
            // Set the method control parameters
            CSAPData.BAPI_METHOD_PROJECT stMethod = new CSAPData.BAPI_METHOD_PROJECT();
            stMethod.Refnumber = "00000" + iRefNum.ToString();
            stMethod.Objecttype = "NETWORKACTIVITY";
            stMethod.Method = sMethod;
            stMethod.Objectkey = oMSPTask.Text30 + oMSPTask.Text1;
            arPrjMethod.Add(stMethod);

            // Only add a single entry for the SAVE 
            if (bSaveMethodAdded == false)  // Save method call has not been added
            {
                CSAPData.BAPI_METHOD_PROJECT stMethodSave = new CSAPData.BAPI_METHOD_PROJECT();
                stMethodSave.Method = "SAVE";
                arPrjMethod.Add(stMethodSave);
                bSaveMethodAdded = true;
            }

            // Define the Network associated with the Activity
            CSAPData.BAPI_NETWORK oNtwk = new CSAPData.BAPI_NETWORK();
            oNtwk.Network = oMSPTask.Text30;
            arNtwks.Add(oNtwk);

            // Set the Activity attributes
            CSAPData.BAPI_NETWORK_ACTIVITY oNtwkActivity = new CSAPData.BAPI_NETWORK_ACTIVITY();
            CSAPData.BAPI_NETWORK_ACTIVITY_UP oNtwkActivityUp = new CSAPData.BAPI_NETWORK_ACTIVITY_UP();
            oNtwkActivity.Network = oMSPTask.Text30;
            oNtwkActivity.Activity = oMSPTask.Text1;
            oNtwkActivity.Control_Key = "PS03";    //Required by SAP PS configuration   
            oNtwkActivity.Description = oMSPTask.Name;
            arNtwkActivities.Add(oNtwkActivity);
            // Fill their update structure
            oNtwkActivityUp.Description = "X";
            arNtwkActivitiesUpDate.Add(oNtwkActivityUp);

        }

        public String MergeSAPMSPProject()
        {  // Intergrates the SAP and MSP project data
            //
            //   Loop over the MSP tasks
            //     if the task is an SAP task
            //       Find the SAP task in the CTaskHierarchy:oTaskCollection 
            //          if we have a match, update the MSP task
            //          else 
            //            after looping over the MSP tasks 
            //                 add all the new task
            //                 below it's parent
            //                 adjusted for level
            //     else (this is not an SAP task)
            //        Continue
            //
            CTaskHier oTaskH = new CTaskHier();
            CTaskHier oTaskHBuff = new CTaskHier();
            MSProject.Task oTask;
            String sCurrNetwork = "";
            int i = 1;      // iterative index for the TaskCollection
            int j = 0;      // pointer into the current array structure

            //Need to process oTaskCollection is sequence with rProject.Tasks
            oTaskH = oTaskCollection.MoveFirst();

            //Main Updating loop;  Updating existing MSP Tasks with SAP Task data
            foreach (MSProject.Task oMSPTask in rProject.Tasks)
            {
                if (oMSPTask.GetField(Microsoft.Office.Interop.MSProject.PjField.pjTaskOutlineCode1) == "No")
                {
                    // We have an MSP (non-SAP task)
                    i++;  // increment our pointer into the MSP Task structure
                    continue;
                }
                else
                {
                        switch (oMSPTask.Text16)  
                        {
                            case TaskType.WBS:
                                oTaskHBuff = oTaskCollection.Find(oMSPTask.Text16, oMSPTask.Text29, oMSPTask.Text1);
                                // If we don't find it in the SAP task collection we process at the end of sub
                                if (oTaskHBuff != null)
                                {
                                    oTask = oMSPTask;
                                    BuildMSPWBSElement(ref oTask, i++, oTaskHBuff.ArrayPtr, oMSPTask.OutlineLevel);
                                    oTaskHBuff.Flag = true;
                                }
                                break;
                            case TaskType.Network:
                                // Network & first Activity with an associated Network
                                //  are packaged by SAP into the first Activity
                                oTaskHBuff = oTaskCollection.Find("Activtiy", oMSPTask.Text30, oMSPTask.Text1);
                                // If we don't find it in the SAP task collection we process at the end of sub
                                if (oTaskHBuff != null)
                                {
                                    oTask = oMSPTask;
                                    BuildMSPNetwork(ref oTask, i++, oTaskHBuff.ArrayPtr, oMSPTask.OutlineLevel);
                                    oTaskHBuff.Flag = true;
                                }
                                break;
                            case TaskType.Activity:
                                // Find the correct Activity
                                oTaskHBuff = oTaskCollection.Find(oMSPTask.Text16, oMSPTask.Text30, oMSPTask.Text1);
                                // If we don't find it in the SAP task collection we process at the end of sub
                                if (oTaskHBuff != null)
                                {
                                    oTask = oMSPTask;
                                    BuildMSPTask(ref oTask, i++, oTaskHBuff.ArrayPtr, oMSPTask.OutlineLevel);
                                    oTaskHBuff.Flag = true;
                                }
                               break;
                             default:
                                oTask = rProject.Tasks.Add("Error", i++);
                                oTask.SetField(Microsoft.Office.Interop.MSProject.PjField.pjTaskOutlineCode1, "No");
                                oTask.Name = "Default condition";
                                break;
                        }  //switch (oTaskH.TaskType)
                }  // do we have an SAP Object
            }  //for each MSP task

            // Add any new SAP objects to the MSP task collection
            //
            // Logic:  Iterate over the oTaskCollection
            //           Use the oTaskH.ArrayPtr to find the objects parent in the 
            //              oTaskCollection.  
            //           Get the parents ID and lookup the corresponding MSP Task
            //           Insert this new task under the found MSP Task
            //

            foreach(CTaskHier oTaskHier in oTaskCollection)
            {
                //Is this an unprocessed SAP task?
                if (oTaskHier.Flag == false)
                {
                    oTaskHBuff = oTaskCollection.MoveBack();

                    switch (oTaskHier.TaskType)
                    {
                        case TaskType.WBS:
                            //Get the parent MSP Task
                            oTask = rProject.Find(oTaskHBuff.TaskType, oTaskHBuff.TaskID.ToString(), oTaskHBuff.SubID);
                             if(oTask != null)
                               i = oTask.Index;
                             else
                               i = rProject.Tasks.Count; //add to the bottom on the collection
                             oTask = rProject.Tasks.Add(arWBSELEMEXP[oTaskHier.ArrayPtr].Wbs_Element, ++i);
                             BuildMSPWBSElement(ref oTask, i, oTaskHier.ArrayPtr, (short)oTaskHier.Level);
                             break;
                        case TaskType.Activity:
                            //Get the parent MSP Task
                             oTask = rProject.Find(TaskType.Network, oTaskHBuff.TaskID.ToString(), oTaskHBuff.SubID);
                             if (oTask != null) //Found the Network so we add the Activity
                             {
                                     i = oTask.Index;
                                     oTask = rProject.Tasks.Add(arNtwkActy[oTaskHier.ArrayPtr].Activity, ++i);
                                     BuildMSPTask(ref oTask, i, oTaskHier.ArrayPtr, (short)(oTaskHier.Level));
                                 //BuildMSPNetwork(ref oTask, i, oTaskHier.ArrayPtr, (short)oTaskHier.Level);
                             }
                             else
                             { // Add the new Network & Activity under the parent
                                 // Have we added the new Network already
                                 oTask = rProject.Find(TaskType.Network, arNtwkActy[oTaskHier.ArrayPtr].Network, "");
                                 if (oTask == null)
                                 { //Add new Network task
                                     i = rProject.Tasks.Count; //add to the bottom on the collection
                                     oTask = rProject.Tasks.Add(arNtwkActy[oTaskHier.ArrayPtr].Network, ++i);
                                     BuildMSPNetwork(ref oTask, i, oTaskHBuff.ArrayPtr, (short)oTaskHier.Level);
                                     //BuildMSPTask(ref oTask, i, oTaskHier.ArrayPtr, (short)oTaskHier.Level);
                                 }
                                 oTask = rProject.Tasks.Add(arNtwkActy[oTaskHier.ArrayPtr].Activity, ++i);
                                 BuildMSPTask(ref oTask, i, oTaskHier.ArrayPtr, (short)(oTaskHier.Level+1));
                             } 
                            break;
                        default:
                            break;
                    }  //switch (oTaskH.TaskType)
                }// oTaskH.Flag == false
            } //foreach (CTaskHier oTaskH in oTaskCollection)

            return "SAP/MSP project data merged successfully";
        }

        public String DisplayProject(String sProjectID)
        {
            Boolean bFileExists = false;

            // Try opening an existing .mpp file for this project
            bFileExists = rProject.OpenMSPProject(sProjectID);
            rProject.ProjectID = sProjectID;
            // Get SAP ProjDef level data
            GetProjectDef(sProjectID);
            // Get most of the SAP PS Project data
            GetWBSData(sProjectID, "");
            // Build the PS Project hierarchy from SAP (populates CSAPInterface:oTaskCollection)
            GetMSPTaskHierarchy();
            // Merge SAP data with existing .mpp file
            return MergeSAPMSPProject();

        }


        public String MaintainSAPProject()
        //   Assumption: Projects are created in SAP
        //     - Provide support for creating/updating Activity & Network
        //     - Can add support for updating WBSs, 
        //       and updating/creating WBS Milestones, WBS Hierarchies
        {

            // MaintainSAPProject
            CSAPData.BAPI_NETWORK_ACTIVITYTable arNtwkActivities = new CSAPData.BAPI_NETWORK_ACTIVITYTable();
            CSAPData.BAPI_ACT_ELEMENTTable arActyElements = new CSAPData.BAPI_ACT_ELEMENTTable();
            CSAPData.BAPI_ACT_ELEMENT_UPDTable arActyElementsUpDate = new CSAPData.BAPI_ACT_ELEMENT_UPDTable();
            CSAPData.BAPI_ACT_MILESTONETable arActyMilestones = new CSAPData.BAPI_ACT_MILESTONETable();
            CSAPData.BAPI_ACT_MILESTONE_UPDTable arActyMilestonesUpDate = new CSAPData.BAPI_ACT_MILESTONE_UPDTable();
            CSAPData.BAPI_NETWORK_ACTIVITY_UPTable arNtwkActivitiesUpDate = new CSAPData.BAPI_NETWORK_ACTIVITY_UPTable();
            CSAPData.BAPI_METHOD_PROJECTTable arPrjMethod = new CSAPData.BAPI_METHOD_PROJECTTable();
            CSAPData.BAPI_NETWORKTable arNtwks = new CSAPData.BAPI_NETWORKTable();
            CSAPData.BAPI_NETWORK_UPDATETable arNtwksUpDate = new CSAPData.BAPI_NETWORK_UPDATETable();
            CSAPData.BAPI_NETWORK_RELATIONTable arRltns = new CSAPData.BAPI_NETWORK_RELATIONTable();
            CSAPData.BAPI_NETWORK_RELATION_UPTable arRltnsUpDate = new CSAPData.BAPI_NETWORK_RELATION_UPTable();
            CTaskHier oTaskH = new CTaskHier();
            CTaskHier oTaskHBuff = new CTaskHier();
            Boolean bReturn = new Boolean();

            // We save Tasks from MSP to SAP as Network Activities
            //   No WBSs or Networks are added.
            int iRefNum = new int();  //Reference parameter required by SAP
            foreach (MSProject.Task oMSPTask in rProject.Tasks)
            {
                // Maintain Activity
                // We are only adding MSP tasks flagged for SAP and of type Activity
                if (oMSPTask.GetField(Microsoft.Office.Interop.MSProject.PjField.pjTaskOutlineCode1) != "No"
                      && (oMSPTask.Text16 == TaskType.Activity))
                {
                    // Search for the Activity in the SAP Collection.  Is it a new addition?
                    oTaskHBuff = oTaskCollection.Find(oMSPTask.Text16, oMSPTask.Text30, oMSPTask.Text1);
                    if (oTaskHBuff != null) // Found the related Activity
                    {
                        iRefNum++;
                        MaintainSAPActivity(oMSPTask, "UPDATE", iRefNum, arPrjMethod, arNtwks, arNtwkActivities, arNtwkActivitiesUpDate); 
                    }
                    else  // New Activity so we add it
                    {
                        // Assign Activity characteristics
                        iRefNum++;
                        MaintainSAPActivity(oMSPTask, "CREATE", iRefNum, arPrjMethod, arNtwks, arNtwkActivities, arNtwkActivitiesUpDate);
                    }

                } //if (oMSPTask.GetField(Microsoft....PjField.pjTaskOutlineCode1) == "No")
            } //foreach (MSProject.Task oMSPTask in rProject.Tasks)


            try
            {
                bReturn = oSAP.MaintainNetworkActivities(out stBAPIReturn1, ref arMsgs, ref arNtwkActivities,
                                                         ref arActyElements, ref arActyElementsUpDate, ref arActyMilestones,
                                                         ref arActyMilestonesUpDate, ref arNtwkActivitiesUpDate, ref arPrjMethod,
                                                         ref arNtwks, ref arNtwksUpDate, ref arRltns, ref arRltnsUpDate);
                bSaveMethodAdded = false;  //flag so we only add one "SAVE" method to this fcn call

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in MaintainSAPProject.");
                return "Error in CSAPInterface:MaintainSAPProject";
            }

            if (bReturn == true)
                return "Project saved to SAP successfully.";
            else
            {
                if (arMsgs != null)
                {

                    MessageBox.Show("Error in MaintainSAPProject.");
                    frmSAPData frmSAP = new frmSAPData(this, rProject.ProjectID);
                    frmSAP.Show();
                    frmSAP.SetTab(4);  // set to error message tab
                }
            }

            return "Error,MaintainSAPProject;" + stBAPIReturn1.Message;
        }

               
        public String RefreshPrjData()
        {
            // Get SAP ProjDef level data
            GetProjectDef(rProject.ProjectID);
            // Get most of the SAP PS Project data
            GetWBSData(rProject.ProjectID, "");
            // Build the PS Project hierarchy from SAP (populates CSAPInterface:oTaskCollection)
            GetMSPTaskHierarchy();
            // Merge the existing MSP and SAP data
            MergeSAPMSPProject();
            return "Project Refreshed";



        }


        public DateTime FromSAPDateConversion(String sDate)
        {
            DateTime dParseDate = new DateTime();

            //                       MM             /            DD               /        YYYY
            String sBuff = sDate.Substring(4, 2) + "/" + sDate.Substring(6, 2) + "/" + sDate.Substring(0, 4);
            try
            {
                dParseDate = DateTime.Parse(sBuff);
            }
            catch
            {
                dParseDate = DateTime.Today;   
            }

            return dParseDate;
        }

        public String ToSAPDateConversion(DateTime dDate)
        {
            //           YYYY                     MM                      DD  ex. 20090908
            return dDate.Year.ToString() + dDate.Month.ToString() + dDate.Day.ToString();
        }


        // The following calls populate frmSAP with already loaded data (Debugging)

        public String ShowSAPData_ProjectDef(ref ListBox lstBoxProjectDef)
        {

            try
            {
                lstBoxProjectDef.Items.Add("App No:       " + stProjectDetail.Applicant_No);
                lstBoxProjectDef.Items.Add("Bdgt Profile: " + stProjectDetail.Budget_Profile);
                lstBoxProjectDef.Items.Add("Bus Area:     " + stProjectDetail.Business_Area);
                lstBoxProjectDef.Items.Add("Calendar:     " + stProjectDetail.Calendar);
                lstBoxProjectDef.Items.Add("Company CO:   " + stProjectDetail.Company_Code);
                lstBoxProjectDef.Items.Add("Control Area: " + stProjectDetail.Controlling_Area);
                lstBoxProjectDef.Items.Add("Description:  " + stProjectDetail.Description);
                lstBoxProjectDef.Items.Add("Distr Chan:   " + stProjectDetail.Distr_Chan);
                lstBoxProjectDef.Items.Add("Division:     " + stProjectDetail.Division);
                lstBoxProjectDef.Items.Add("Dli Profile:  " + stProjectDetail.Dli_Profile);
                lstBoxProjectDef.Items.Add("Equity Type:  " + stProjectDetail.Equity_Typ);
                lstBoxProjectDef.Items.Add("FCst Fin Date:" + stProjectDetail.Fcst_Finish);
                lstBoxProjectDef.Items.Add("FCst Srt Date:" + stProjectDetail.Fcst_Start);
                lstBoxProjectDef.Items.Add("Finish:       " + stProjectDetail.Finish);
                lstBoxProjectDef.Items.Add("Func Area:    " + stProjectDetail.Func_Area);
                lstBoxProjectDef.Items.Add("Grping Indic: " + stProjectDetail.Grouping_Indicator);
                lstBoxProjectDef.Items.Add("Interest Prof:" + stProjectDetail.Interest_Prof);
                lstBoxProjectDef.Items.Add("Inverst Prof: " + stProjectDetail.Invest_Profile);
                lstBoxProjectDef.Items.Add("JV JIBCL:     " + stProjectDetail.Jv_Jibcl);
                lstBoxProjectDef.Items.Add("JV JIBSA:     " + stProjectDetail.Jv_Jibsa);
                lstBoxProjectDef.Items.Add("JV OType:     " + stProjectDetail.Jv_Otype);
                lstBoxProjectDef.Items.Add("Lang:         " + stProjectDetail.Langu);
                lstBoxProjectDef.Items.Add("Lang ISO:     " + stProjectDetail.Langu_Iso);
                lstBoxProjectDef.Items.Add("Locattion:    " + stProjectDetail.Location);
                lstBoxProjectDef.Items.Add("Mask ID:      " + stProjectDetail.Mask_Id);
                lstBoxProjectDef.Items.Add("Ntwrk Asgmt:  " + stProjectDetail.Network_Assignment);
                lstBoxProjectDef.Items.Add("Ntwrk Profile:" + stProjectDetail.Network_Profile);
                lstBoxProjectDef.Items.Add("Obj Class:    " + stProjectDetail.Objectclass);
                lstBoxProjectDef.Items.Add("Partner Prof: " + stProjectDetail.Partner_Profile);
                lstBoxProjectDef.Items.Add("Plan Basic:   " + stProjectDetail.Plan_Basic);
                lstBoxProjectDef.Items.Add("Plan Fcst:    " + stProjectDetail.Plan_Fcst);
                lstBoxProjectDef.Items.Add("Plan Prof:    " + stProjectDetail.Plan_Profile);
                lstBoxProjectDef.Items.Add("Plan Integ:   " + stProjectDetail.Planintegrated);
                lstBoxProjectDef.Items.Add("Plant:        " + stProjectDetail.Plant);
                lstBoxProjectDef.Items.Add("Profit Ctr:   " + stProjectDetail.Profit_Ctr);
                lstBoxProjectDef.Items.Add("Prj Curr:     " + stProjectDetail.Project_Currency);
                lstBoxProjectDef.Items.Add("Prj Curr ISO: " + stProjectDetail.Project_Currency_Iso);
                lstBoxProjectDef.Items.Add("Prj Def:      " + stProjectDetail.Project_Definition);
                lstBoxProjectDef.Items.Add("Prj Profile:  " + stProjectDetail.Project_Profile);
                lstBoxProjectDef.Items.Add("Prj Stock:    " + stProjectDetail.Project_Stock);
                lstBoxProjectDef.Items.Add("Rec Ind:      " + stProjectDetail.Rec_Ind);
                lstBoxProjectDef.Items.Add("Res Anal Key: " + stProjectDetail.Res_Anal_Key);
                lstBoxProjectDef.Items.Add("Responsible #:" + stProjectDetail.Responsible_No);
                lstBoxProjectDef.Items.Add("Sales Org:    " + stProjectDetail.Salesorg);
                lstBoxProjectDef.Items.Add("Sched Scenrio:" + stProjectDetail.Sched_Scenario);
                lstBoxProjectDef.Items.Add("Sim Prof:     " + stProjectDetail.Simulation_Profile);
                lstBoxProjectDef.Items.Add("Start:        " + stProjectDetail.Start);
                lstBoxProjectDef.Items.Add("Stat Prof:    " + stProjectDetail.Stat_Prof);
                lstBoxProjectDef.Items.Add("Statistical:  " + stProjectDetail.Statistical);
                lstBoxProjectDef.Items.Add("System Status:" + stProjectDetail.System_Status);
                lstBoxProjectDef.Items.Add("Tax Jur Cde:  " + stProjectDetail.Taxjurcode);
                lstBoxProjectDef.Items.Add("Time Unit:    " + stProjectDetail.Time_Unit);
                lstBoxProjectDef.Items.Add("Time Unit ISO:" + stProjectDetail.Time_Unit_Iso);
                lstBoxProjectDef.Items.Add("Val Spec Stck:" + stProjectDetail.Valuation_Spec_Stock);
                lstBoxProjectDef.Items.Add("Venture:      " + stProjectDetail.Venture);
                lstBoxProjectDef.Items.Add("WBS Sch Prof: " + stProjectDetail.Wbs_Sched_Profile);
                lstBoxProjectDef.Items.Add("WBS Stat Prof:" + stProjectDetail.Wbs_Status_Profile);

                return "Project Detail data loaded.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in ShowSAPData_PrjDefData.");
                return "Error in CSAPInterface:ShowSAPData_PrjDefData";
            }
        }



        public String ShowSAPData_WBSData(ref DataTable oNtwkActy,
                                          ref DataTable oMsgs,
                                          ref DataTable oWBSELEMEXP,
                                          ref DataTable oWBSHRCY,
                                          ref DataTable oWBSMLST)
        {
            oWBSELEMEXP = arWBSELEMEXP.ToADODataTable();
            oNtwkActy = arNtwkActy.ToADODataTable();
            oWBSHRCY = arWBSHRCY.ToADODataTable();
            oWBSMLST = arWBSMLST.ToADODataTable();
            oMsgs = arMsgs.ToADODataTable();

            return "WBS data loaded.";
        }

         public String ShowSAPData_TaskHierarchy(ref DataTable oTaskHierarchy)
        // Display the data captured in the CSAPInterface:GetMSPTaskHierarchy call
        //    Assumption:  CSAPInterface:GetMSPTaskHierarchy was already called.
        {
            oTaskHierarchy = oTaskCollection.TaskHierarchy;
            return "MSP Task Hierarchy loaded.";

        }

    }
}
