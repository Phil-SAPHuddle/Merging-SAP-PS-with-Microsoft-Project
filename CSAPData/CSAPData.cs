using System;
using System.Windows.Forms;
using System.Data;
using SAP.Connector;

namespace CSAPData
{
	/// <summary>
	/// BAPI developed interface for interacting with SAP
    ///   This class(DLL) is developed in MS VS 2003
	/// </summary>
	/// 

	public class CSAPPrj
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>

		// Main handle to the SAP Proxy
		private SAPProxy1 oSAPProxy = new SAPProxy1();

		public CSAPPrj(){}

		public Boolean LogIn(String sConn)
		{
			try
			{
                oSAPProxy.Connection = SAP.Connector.SAPConnection.GetConnection(sConn);
                oSAPProxy.Connection.Open();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "CSAPPrj:LogIn.");
			}

            if (oSAPProxy.Connection.IsOpen == true)
				return true;
			else
				return false;
		}

		public void LogOut()
		{
            try
            {
                oSAPProxy.Connection.Close();
                oSAPProxy.Dispose();
            }
            catch (Exception ex)
            {}
        }

        public Boolean GetProjectDefData(String sLang,
			                             string sProjName,
			                             ref BAPI_BUS2001_DETAIL stProjectDetail)
		{
			BAPIRET2Table ETReturn = new BAPIRET2Table();	
			BAPIPAREXTable Extensionin = new BAPIPAREXTable();
			BAPIPAREXTable Extensionout= new BAPIPAREXTable();

			try
			{
                oSAPProxy.Bapi_Bus2001_Getdata(sLang, sProjName,
                                               out stProjectDetail, 
                                               ref ETReturn, 
                                               ref Extensionin, 
                                               ref Extensionout);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "CSAPPrj:GetProjectDefData.");
				return false;
			}

			return true;
		}

		public Boolean GetPrjWBSData(string sProjName,
									 String sWithActivities,
									 String sWithMilestones,
									 String sWithSubtree,
									 out BAPI_PROJECT_DEFINITION_EX stPrjDef,
									 out BAPIRETURN1	stBAPIReturn1,
									 ref BAPI_NETWORK_ACTIVITY_EXPTable arNtwkActy,
									 ref BAPI_METH_MESSAGETable arMsgs,
									 ref BAPI_WBS_ELEMENT_EXPTable arWBSELEMEXP,
									 ref BAPI_WBS_HIERARCHIETable arWBSHRCY,
									 ref BAPI_WBS_MILESTONE_EXPTable arWBSMLST,
									 ref BAPI_WBS_ELEMENTSTable arWBSELEMSELECT)
		{

			try
			{
                oSAPProxy.Bapi_Project_Getinfo(sProjName,
					sWithActivities , 
					sWithMilestones,
					sWithSubtree,
					out stPrjDef, 
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
				MessageBox.Show(ex.Message, "CSAPPrj:GetPrjWBSData.");
				stPrjDef = null;
				stBAPIReturn1 = null;
				return false;
			}

			return true;
		}

		public Boolean MaintainNetworkActivities(
			out BAPIRETURN1 stBAPIReturn1,
			ref BAPI_METH_MESSAGETable arMsgs,
			ref BAPI_NETWORK_ACTIVITYTable arNtwkActivities,
			ref BAPI_ACT_ELEMENTTable arActyElements,
			ref BAPI_ACT_ELEMENT_UPDTable arActyElementsUpDate,
			ref BAPI_ACT_MILESTONETable arActyMilestones,
			ref BAPI_ACT_MILESTONE_UPDTable arActyMilestonesUpDate,
			ref BAPI_NETWORK_ACTIVITY_UPTable arNtwkActivitiesUpDate,
			ref BAPI_METHOD_PROJECTTable arPrjDef,
			ref BAPI_NETWORKTable arNtwks,
			ref BAPI_NETWORK_UPDATETable arNtwksUpDate,
			ref BAPI_NETWORK_RELATIONTable arRltns,
			ref BAPI_NETWORK_RELATION_UPTable arRltnsUpDate)

		{
			try
			{
				oSAPProxy.Bapi_Network_Maintain(out stBAPIReturn1,
												ref arMsgs,
												ref arNtwkActivities,
												ref arActyElements,
												ref arActyElementsUpDate,
												ref arActyMilestones,
												ref arActyMilestonesUpDate,
												ref arNtwkActivitiesUpDate,
												ref arPrjDef,
												ref arNtwks,
												ref arNtwksUpDate,
												ref arRltns,
												ref arRltnsUpDate);
			}
			catch (Exception ex)
			{
				stBAPIReturn1 = new BAPIRETURN1();
				stBAPIReturn1.Type = "E";
				stBAPIReturn1.Message = "MaintainNetworkActivities call failed.";
				stBAPIReturn1.Message_V1 = ex.Message;
				return false;
			}

			if(stBAPIReturn1.Type == "E") //Error
	          return false;
			else
			return true;

		}

		public Boolean ProjectDef_GetList(int iMaxRows,
			                              ref BAPIPREXPTable arProjDef,
										  ref BAPI_2002_DESCR_RANGETable arSAPDescrRange,
										  ref BAPI_2002_PD_RANGETable arSAPProjIDRange)
		{
			BAPIRET2 ETReturn = new BAPIRET2();	
			try
			{
				oSAPProxy.Bapi_Projectdef_Getlist(iMaxRows, 
					out ETReturn, ref arSAPDescrRange, ref arProjDef, ref arSAPProjIDRange);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "CSAPPrj:ProjectDef_GetList");
				return false;
			}

				return true;
		}			                           

    }

}  //namespace CSAPData
