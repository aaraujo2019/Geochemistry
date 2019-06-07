using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsGCSamplesRockMiner
{
    public string sOpcion;
    public string sSample;
    public string sMineral;
    public string sMinStyle;
    public double? dMinPerc;
    public string sObservations;
    public int iSKSampleRMin;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string GCSamplesRockMin_Add()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(7);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;

            arr[2].ParameterName = "@MZMin";
            if (sMineral == null)
                arr[2].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[2].Value = sMineral;

            arr[3].ParameterName = "@MZStyle";
            if (sMinStyle == null)
                arr[3].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[3].Value = sMinStyle;

            arr[4].ParameterName = "@MZPerc";
            if (dMinPerc == null)
                arr[4].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[4].Value = dMinPerc;

            arr[5].ParameterName = "@Obsevations";
            if (sObservations == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = sObservations;

            arr[6].ParameterName = "@SKSampleRMin";
            arr[6].Value = iSKSampleRMin;

            oRes = oData.ExecuteScalar("usp_GC_SampleRockMin_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error GCSamplesRockMin. " + eX.Message); ;
        }
    }

    public string GCSamplesRockMin_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKSampleRMin";
            arr[0].Value = iSKSampleRMin;

            oRes = oData.ExecuteScalar("usp_GC_SampleRockMin_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error GCSamplesRockMin. " + eX.Message); ;
        }
    }

    public DataTable getGCSamplesRockMinList()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SampleRockMin_List", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRockMin: " + eX.Message);
        }
    }

    public DataTable getGCSamplesRockMinListReport()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Sample";
            arr[0].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SampleRockMin_ListReport", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRockMinReport: " + eX.Message);
        }
    }
    
}

