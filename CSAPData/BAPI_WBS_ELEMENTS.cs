
//------------------------------------------------------------------------------
// 
//     This code was generated by a SAP. NET Connector Proxy Generator Version 2.0
//     Created at 8/29/2009
//     Created from Windows
//
//     Changes to this file may cause incorrect behavior and will be lost if 
//     the code is regenerated.
// 
//------------------------------------------------------------------------------
using System;
using System.Text;
using System.Collections;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using System.Xml.Schema;
using System.Web.Services;
using System.Web.Services.Description;
using System.Web.Services.Protocols;
using SAP.Connector;

namespace CSAPData
{

  /// <summary>
  /// List: WBS elements
  /// </summary>
  [RfcStructure(AbapName ="BAPI_WBS_ELEMENTS" , Length = 24, Length2 = 48)]
  [Serializable]
  public class BAPI_WBS_ELEMENTS : SAPStructure
  {
   

    /// <summary>
    /// Work Breakdown Structure Element (WBS Element)
    /// </summary>
 
    [RfcField(AbapName = "WBS_ELEMENT", RfcType = RFCTYPE.RFCTYPE_CHAR, Length = 24, Length2 = 48, Offset = 0, Offset2 = 0)]
    [XmlElement("WBS_ELEMENT", Form=XmlSchemaForm.Unqualified)]
    public string Wbs_Element
    { 
       get
       {
          return _Wbs_Element;
       }
       set
       {
          _Wbs_Element = value;
       }
    }
    private string _Wbs_Element;

  }

}
