using System;
using System.Data;
using System.Reflection;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ProjectAddIn2
{
    public class TaskType
    {
        public const string ProjectSummary = "ProjectSummary";
        public const string WBS = "WBS";
        public const string Network = "Network";
        public const string Task = "Task";
        public const string Activity = "Activity";
        public const string Component = "Component";
        public TaskType() { }
    }

    public class CTaskHier
    {
        private int nID;
        private int nArrayPtr;
        private String sElementID;
        private String sSubID;
        private String sTaskType;
        private int nLevel;
        private String sDescription;
        private Boolean bFlag;             // Flags an item as processed

        public CTaskHier() { }

        public int TaskID
        {
            get { return nID; }
            set { nID = value; }
        }

        public int ArrayPtr
        {
            get { return nArrayPtr; }
            set { nArrayPtr = value; }
        }

        public String ElementID
        {
            get { return sElementID; }
            set { sElementID = value; }
        }

        public String SubID
        {
            get { return sSubID; }
            set { sSubID = value; }
        }

        public String TaskType
        {
            get { return sTaskType; }
            set { sTaskType = value; }
        }

        public int Level
        {
            get { return nLevel; }
            set { nLevel = value; }
        }

        public String TaskDesc
        {
            get
            {
                if (sDescription != null) { return sDescription; }
                else { return ""; }
            }
            set { sDescription = value; }
        }

        public Boolean Flag
        {
            get { return bFlag; }
            set { bFlag = value; }
        }
    }

    public class CTaskHierarchy : IEnumerable<CTaskHier>
    {// Build our fully iterable collection class to store the
     //   SAP project objects

        public CTaskHierarchy() { }

        // Define a Typed Collection  
        private List<CTaskHier> arTaskHierarchy = new List<CTaskHier>();
        private int iPosition = -1;  //Position in the collection

        // Required, implement the IEnumerable.GetEnumerator to provide the
        //   ICollection ForEach iterator.  The generic IEnumerable<T> extends IEnumerable, so we need to 
        //   implement both versions of GetEnumerator().
        IEnumerator<CTaskHier> IEnumerable<CTaskHier>.GetEnumerator()
        { return arTaskHierarchy.GetEnumerator(); }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        { return arTaskHierarchy.GetEnumerator(); }


        public CTaskHier MoveFirst()
        {
            //Check we have a populated collection
            if (arTaskHierarchy.Count <= 0)
                return null;
            //Set to the first element
           iPosition = 0;

            //Check collection end point
            if (iPosition >= arTaskHierarchy.Count)
                return null;

            return arTaskHierarchy[iPosition];
        }

        public CTaskHier MoveNext()
        {
            //Check we have a populated collection
            if (arTaskHierarchy.Count <= 0)
                return null;
            //Increment the position pointer
            if (iPosition == -1) iPosition = 0;
            else iPosition++;

            //Check collection end point
            if (iPosition >= arTaskHierarchy.Count)
                return null;

            return arTaskHierarchy[iPosition];
        }

        public CTaskHier MoveBack()
        {
            //Check we have a populated collection
            if (arTaskHierarchy.Count <= 0)
                return null;
            //Increment the position pointer
            if (iPosition <= 0) iPosition = 0;
            else iPosition--;

            //Check collection end point
            if (iPosition >= arTaskHierarchy.Count)
                return null;

            return arTaskHierarchy[iPosition];
        }
        public CTaskHier Find(String sType, String sID, String sSubID)
        {
            return arTaskHierarchy.Find(delegate(CTaskHier obj)
            { return (obj.TaskType == sType && obj.ElementID == sID && obj.SubID == sSubID); });
        }

        public void GetTaskHierarchy(CSAPData.BAPI_WBS_HIERARCHIETable arWBSHRCY,
                                     CSAPData.BAPI_NETWORK_ACTIVITY_EXPTable arNtwkActy)
        {
            CTaskHier buffTask = new CTaskHier();
            int iCount = 0;
            ////// Remove all values in the Task Hierarchy "List" collection
            ////// The lamba expression with the => syntax removes all values that are not null
            arTaskHierarchy.RemoveAll(Item => Item != null);
            //Process WBS hierarchy
            for (int i = 0; i < arWBSHRCY.Count; i++)
            {
                CTaskHier oWBStask = new CTaskHier();

                // Start at the root
                if (i == 0)
                {
                    oWBStask.Level = 1;
                }
                else
                {  // Check if the "Up" member is different
                    if (arWBSHRCY[i].Wbs_Element != arWBSHRCY[i].Up)
                    {
                        //Get the level of the "Up" ElementID
                        foreach (CTaskHier t in arTaskHierarchy)
                        {
                            if (t.ElementID == arWBSHRCY[i].Up)
                            {
                                oWBStask.Level = t.Level + 1;
                                break;
                            }
                        }
                    }
                }
                oWBStask.TaskID = i;
                oWBStask.ArrayPtr = i;
                oWBStask.TaskType = TaskType.WBS;    
                oWBStask.ElementID = arWBSHRCY[i].Wbs_Element;
                oWBStask.SubID = "";
                arTaskHierarchy.Add(oWBStask);
                oWBStask = null;
            }

            //Process Networks and Activities
            // Note: Networks with no Activities are not processed
            iCount = 1;
            for (int i = 0; i < arNtwkActy.Count; i++)
            {
                // Get the level of the associated WBS
                foreach (CTaskHier t in arTaskHierarchy)
                {
                    if (t.ElementID == arNtwkActy[i].Wbs_Element)
                    {
                        buffTask = t;
                        break;
                    }
                }
                CTaskHier otask = new CTaskHier();
                otask.TaskID = buffTask.TaskID + iCount++;  //We insert after the found WBS
                otask.ArrayPtr = i;
                otask.TaskType = TaskType.Activity;
                otask.ElementID = arNtwkActy[i].Network;
                otask.SubID = arNtwkActy[i].Activity;
                otask.Level = buffTask.Level + 1;
                otask.TaskDesc = arNtwkActy[i].Description;
                arTaskHierarchy.Insert(otask.TaskID, otask);
                otask = null;
            }

            // Reorder the task id
            for (int i = 0; i <= arTaskHierarchy.Count - 1; i++)
            {
                arTaskHierarchy[i].TaskID = i;
            }

        }

        public DataTable TaskHierarchy
        {  //returns the calculated MS Project Hierarchy as a DataTable object
            get
            {  // Thank you StevenMcD.net 
               //     http://www.stevenmcd.net/2008/12/convert-linq-resultset-datatable/

            // Create DataTable to Fill
            DataTable _newDataTable = new DataTable();
            // Retrieve the Type passed into the Method
            Type _impliedType = typeof(CTaskHier);
            //Get an array of the Type’s properties
            PropertyInfo[] _propInfo = _impliedType.GetProperties();
            //Create the columns in the DataTable
            foreach (PropertyInfo pi in _propInfo)
            {
                _newDataTable.Columns.Add(pi.Name, pi.PropertyType);
            }

            //Populate the table
            foreach (CTaskHier item in arTaskHierarchy)
            {
                DataRow _newDataRow = _newDataTable.NewRow();
                _newDataRow.BeginEdit();

                foreach (PropertyInfo pi in _propInfo)
                {
                    _newDataRow[pi.Name] = pi.GetValue(item, null);
                }

                _newDataRow.EndEdit();
                _newDataTable.Rows.Add(_newDataRow);
            }

            return _newDataTable;
            }
        }

    }

}
