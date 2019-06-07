using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsGCSamplesRockAlterations
{
    public string sOpcion;
    public string sSample;
    public string sAltType;
    public string sAltInt;
    public string sAltStyle;
    public string sAltMin;
    public string sAltMin2;
    public string sAltMin3;
    public string sObservations;
    public int iSkSampleRAlt;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();


    public string GCSamplesRockAlt_Add()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(10);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            
            arr[2].ParameterName = "@ALTType";
            if (sAltType == null)
                arr[2].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[2].Value = sAltType;
            
            arr[3].ParameterName = "@ALTInt";
            if (sAltInt == null)
                arr[3].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[3].Value = sAltInt;
            
            arr[4].ParameterName = "@ALTStyle";
            if (sAltStyle == null)
                arr[4].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[4].Value = sAltStyle;

            arr[5].ParameterName = "@ALTMin";
            if (sAltMin == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = sAltMin;
            
            arr[6].ParameterName = "@ALTMin2";
            if (sAltMin2 == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sAltMin2;

            arr[7].ParameterName = "@ALTMin3";
            if (sAltMin3 == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sAltMin3;

            arr[8].ParameterName = "@Obsevations";
            if (sObservations == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = sObservations;
            

            arr[9].ParameterName = "@SKSampleRAlt";
            arr[9].Value = iSkSampleRAlt;

            oRes = oData.ExecuteScalar("usp_GC_SampleRockAlt_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error GCSamplesRockAlt. " + eX.Message); ;
        }
    }

    public string GCSamplesRockAlt_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKSampleRAlt";
            arr[0].Value = iSkSampleRAlt;

            oRes = oData.ExecuteScalar("usp_GC_SampleRockAlt_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error GCSamplesRockAlt. " + eX.Message); ;
        }
    }

    public DataTable getGCSamplesRockAltList()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SampleRockAlt_List", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRockAlt: " + eX.Message);
        }
    }

    public DataTable getGCSamplesRockAltListReport()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Sample";
            arr[0].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SampleRockAlt_ListReport", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRockAltReport: " + eX.Message);
        }
    }
}

