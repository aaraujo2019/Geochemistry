using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsCHAlterations
{
    public string sOpcion;
    public string sChid;
    public string sSample;
    public string sAltType;
    public string sAltInt;
    public string sAltStyle;
    public string sAltMin;
    public string sAltMin2;
    public string sAltMin3;
    public string sObservations;
    public int iSKAlteration;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string CHAlterations_Add()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(11);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Chid";
            arr[1].Value = sChid;
            arr[2].ParameterName = "@Sample";
            arr[2].Value = sSample;

            arr[3].ParameterName = "@ALTType";
            if (sAltType == null)
                arr[3].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[3].Value = sAltType;

            arr[4].ParameterName = "@ALTInt";
            if (sAltInt == null)
                arr[4].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[4].Value = sAltInt;

            arr[5].ParameterName = "@ALTStyle";
            if (sAltStyle == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = sAltStyle;

            arr[6].ParameterName = "@ALTMin";
            if (sAltMin == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sAltMin;

            arr[7].ParameterName = "@ALTMin2";
            if (sAltMin2 == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sAltMin2;

            arr[8].ParameterName = "@ALTMin3";
            if (sAltMin3 == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = sAltMin3;

            arr[9].ParameterName = "@Obsevations";
            if (sObservations == null)
                arr[9].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[9].Value = sObservations;

            arr[10].ParameterName = "@SKAlteration";
            arr[10].Value = iSKAlteration;

            oRes = oData.ExecuteScalar("usp_CH_Alterations_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error CHAlterations. " + eX.Message); ;
        }
    }

    public string CHAlterations_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKAlteration";
            arr[0].Value = iSKAlteration;

            oRes = oData.ExecuteScalar("usp_CH_Alterations_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error CHAlterations. " + eX.Message); ;
        }
    }

    public DataTable getCHAlterationsList()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(3);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Chid";
            arr[1].Value = sChid;
            arr[2].ParameterName = "@Sample";
            arr[2].Value = sSample;
            dtData = oData.ExecuteDataset("usp_CH_Alterations_List", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHAlterations: " + eX.Message);
        }
    }

    public DataTable getCHAlteration_ListReport()
    {
        try
        {
            DataSet dtAlt = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Chid";
            arr[0].Value = sChid;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtAlt = oData.ExecuteDataset("usp_CH_Alterations_ListReport", arr, CommandType.StoredProcedure);
            return dtAlt.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

}

