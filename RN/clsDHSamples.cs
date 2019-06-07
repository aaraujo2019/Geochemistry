using System;
using System.Collections.Generic;
using System.Text;
using DataAccess;
using System.Data;
using System.Data.SqlClient;


public class clsDHSamples
{
    public string sOpcion;
    public string sHoleID;
    public string lFrom;
    public string lTo;

    public double iFrom;
    public double iTo;
    public string sSample;
    public string sSampleType;
    public string sDupDe;
    public string sLith;
    public Int64 iDHSampID;
    public string sComments;

    public static string sConsLoggin;
    public static string sStaticFrom;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    //[usp_DH_Samples_ListFromToValid]
    public DataTable getDHSamplesFromToValid()
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@From";
            arr[0].Value = iFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = iTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            arr[3].ParameterName = "@DHSamplesID";
            arr[3].Value = iDHSampID; 
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_ListFromToValid", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamplesFromToValid: " + eX.Message);
        }
    }

    //[usp_DH_Alterations_ListValidFromToNext]
    public DataTable getDHSamplesValidFromToNext()
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@HoleID";
            arr[0].Value = sHoleID;
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_ListValidFromToNext", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamplesFromToValidFromToNext: " + eX.Message);
        }
    }

    public DataTable getDHSamples_Litho_ListValid()
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtDHSamples = oData.ExecuteDataset("usp_DHSamples_Litho_ListValid", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamples_Litho_ListValid: " + eX.Message);
        }
    }

    //[usp_DH_Samples_ListFromToValid]
    public DataTable getDHSamplesValid()
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@From";
            arr[0].Value = iFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = iTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            arr[3].ParameterName = "@DHSamplesID";
            arr[3].Value = iDHSampID;
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_ListValid", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamplesFromToValid: " + eX.Message);
        }
    }

    //[usp_DH_Samples_Delete]
    public string DHSamples_DeleteLoggin()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKDHSamples";
            arr[0].Value = iDHSampID;

            oRes = oData.ExecuteScalar("usp_DH_Samples_Delete", arr, CommandType.StoredProcedure);
            //ds = oDAtos.ExecuteDataset("usp_Datos_ListByID", arr, CommandType.StoredProcedure)
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error DHSample. " + eX.Message); ;
        }
    }


    //usp_DH_Samples_InsertLoggin
    public string DHSamples_AddLoggin()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(10);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            arr[2].ParameterName = "@Sample";
            arr[2].Value = sSample;
            arr[3].ParameterName = "@From";
            arr[3].Value = iFrom;
            arr[4].ParameterName = "@To";
            arr[4].Value = iTo;
            arr[5].ParameterName = "@SampleType";
            arr[5].Value = sSampleType;
            arr[6].ParameterName = "@DupDe";
            arr[6].Value = sDupDe;
            arr[7].ParameterName = "@Lithology";
            if (sLith == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sLith;
            arr[8].ParameterName = "@Comments";
            arr[8].Value = sComments;

            arr[9].ParameterName = "@DHSamplesID";
            arr[9].Value = iDHSampID;
            
            oRes = oData.ExecuteScalar("usp_DH_Samples_InsertLoggin", arr, CommandType.StoredProcedure);
            //ds = oDAtos.ExecuteDataset("usp_Datos_ListByID", arr, CommandType.StoredProcedure)
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error DHSamples. " + eX.Message); ;
        }
    }

    public DataTable getDHSamplesList()
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleId";
            arr[1].Value = sHoleID;
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_List", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getDHSamples: " + eX.Message);
        }
    }

    public DataTable getDHSamplesAll()
    {
        try
        {
            DataSet dtDHSamplesAll = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtDHSamplesAll = oData.ExecuteDataset("usp_DH_Samples_ListAll", arr, CommandType.StoredProcedure);
            return dtDHSamplesAll.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamplesAll: " + eX.Message);
        }
    }

    public DataTable getDHSamplesDistinct()
    {
        try
        {
            DataSet dtDHSamplesAll = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtDHSamplesAll = oData.ExecuteDataset("usp_DH_Samples_ListDistinct", arr, CommandType.StoredProcedure);
            return dtDHSamplesAll.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamplesAll: " + eX.Message);
        }
    }

    public DataTable getDHSamplesId()
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@sampleId";
            arr[0].Value = sSample;
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_ListSampleId", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getDHSamplesId: " + eX.Message);
        }
    }

    public DataTable getDHSamples_getRangeValid(string _sIni, string _sFin, string _Id)
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(3);
            arr[0].ParameterName = "@SampleIni";
            arr[0].Value = _sIni;
            arr[1].ParameterName = "@SampleFin";
            arr[1].Value = _sFin;
            arr[2].ParameterName = "@Id";
            arr[2].Value = int.Parse(_Id);
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_ListRange", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamples_getRangeValid: " + eX.Message);
        }
    }

    public DataTable getDHSamples()
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleId";
            arr[1].Value = sHoleID;
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_List_Login", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamples: " + eX.Message);
        }
    }

    public DataTable getDHSamplesBySamp(string _sample1, string _sample2)
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(3);
            arr[0].ParameterName = "@Sample1";
            arr[0].Value = _sample1;
            arr[1].ParameterName = "@Sample2";
            arr[1].Value = _sample2;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_ListSample", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getDHSamplesBySamp: " + eX.Message);
        }
    }

    public DataTable getDHSamplesById(string _Id1, string _Id2)
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(3);
            arr[0].ParameterName = "@Id1";
            arr[0].Value = _Id1;
            arr[1].ParameterName = "@Id2";
            arr[1].Value = _Id2;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_ListDHID", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getDHSamplesBySamp: " + eX.Message);
        }
    }

    //[usp_DH_Samples_ListFromTo]
    public DataTable getDHSamplesFromTo()
    {
        try
        {
            DataSet dtDHSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(3);
            arr[0].ParameterName = "@From";
            arr[0].Value = lFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = lTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            dtDHSamples = oData.ExecuteDataset("usp_DH_Samples_ListFromTo", arr, CommandType.StoredProcedure);
            return dtDHSamples.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHSamplesFromTo: " + eX.Message);
        }
    }

}

