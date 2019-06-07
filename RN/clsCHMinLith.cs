using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsCHMinLith
{
    public string sOpcion;
    public string sSample;
    public string sMineral;
    public string sType;
    public int iSKMinLith;
    public string sChid;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string CHMinLith_Add()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(6);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            arr[2].ParameterName = "@Mineral";
            arr[2].Value = sMineral;
            arr[3].ParameterName = "@TypeMat_Phe";
            arr[3].Value = sType;
            arr[4].ParameterName = "@SKMinLith";
            arr[4].Value = iSKMinLith;
            arr[5].ParameterName = "@Chid";
            arr[5].Value = sChid;

            oRes = oData.ExecuteScalar("usp_CHMinLith_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error CHMinLith. " + eX.Message); ;
        }
    }

    public string CHMinLithLith_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKMinLith";
            arr[0].Value = iSKMinLith;

            oRes = oData.ExecuteScalar("usp_CHMinLith_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error CHMinLithLith. " + eX.Message); ;
        }
    }

    public DataTable getGCSamplesRockLithList()
    {
        try
        {
            DataSet dtCHMin = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtCHMin = oData.ExecuteDataset("usp_CHMinLith_List", arr, CommandType.StoredProcedure);
            return dtCHMin.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHMinLith: " + eX.Message);
        }
    }


}

