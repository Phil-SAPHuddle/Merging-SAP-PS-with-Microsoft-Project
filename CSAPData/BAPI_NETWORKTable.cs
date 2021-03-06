
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
using System.Web.Services;
using System.Web.Services.Description;
using System.Web.Services.Protocols;
using SAP.Connector;

namespace CSAPData
{
  /// <summary>
  /// A typed collection of BAPI_NETWORK elements.
  /// </summary>
  [Serializable]
  public class BAPI_NETWORKTable : SAPTable 
  {
  
    /// <summary>
    /// Returns the element type BAPI_NETWORK.
    /// </summary>
    /// <returns>The type BAPI_NETWORK.</returns>
    public override Type GetElementType() 
    {
        return (typeof(BAPI_NETWORK));
    }

    /// <summary>
    /// Creates an empty new row of type BAPI_NETWORK.
    /// </summary>
    /// <returns>The newBAPI_NETWORK.</returns>
    public override object CreateNewRow()
    { 
        return new BAPI_NETWORK();
    }
     
    /// <summary>
    /// The indexer of the collection.
    /// </summary>
    public BAPI_NETWORK this[int index] 
    {
        get 
        {
            return ((BAPI_NETWORK)(List[index]));
        }
        set 
        {
            List[index] = value;
        }
    }
        
    /// <summary>
    /// Adds a BAPI_NETWORK to the end of the collection.
    /// </summary>
    /// <param name="value">The BAPI_NETWORK to be added to the end of the collection.</param>
    /// <returns>The index of the newBAPI_NETWORK.</returns>
    public int Add(BAPI_NETWORK value) 
    {
        return List.Add(value);
    }
        
    /// <summary>
    /// Inserts a BAPI_NETWORK into the collection at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index at which value should be inserted.</param>
    /// <param name="value">The BAPI_NETWORK to insert.</param>
    public void Insert(int index, BAPI_NETWORK value) 
    {
        List.Insert(index, value);
    }
        
    /// <summary>
    /// Searches for the specified BAPI_NETWORK and returnes the zero-based index of the first occurrence in the collection.
    /// </summary>
    /// <param name="value">The BAPI_NETWORK to locate in the collection.</param>
    /// <returns>The index of the object found or -1.</returns>
    public int IndexOf(BAPI_NETWORK value) 
    {
        return List.IndexOf(value);
    }
        
    /// <summary>
    /// Determines wheter an element is in the collection.
    /// </summary>
    /// <param name="value">The BAPI_NETWORK to locate in the collection.</param>
    /// <returns>True if found; else false.</returns>
    public bool Contains(BAPI_NETWORK value) 
    {
        return List.Contains(value);
    }
        
    /// <summary>
    /// Removes the first occurrence of the specified BAPI_NETWORK from the collection.
    /// </summary>
    /// <param name="value">The BAPI_NETWORK to remove from the collection.</param>
    public void Remove(BAPI_NETWORK value) 
    {
        List.Remove(value);
    }

    /// <summary>
    /// Copies the contents of the BAPI_NETWORKTable to the specified one-dimensional array starting at the specified index in the target array.
    /// </summary>
    /// <param name="array">The one-dimensional destination array.</param>           
    /// <param name="index">The zero-based index in array at which copying begins.</param>           
    public void CopyTo(BAPI_NETWORK[] array, int index) 
    {
        List.CopyTo(array, index);
	}
  }
}
