using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsGCSamplesRockOxides
{
    public string sOpcion;
    public string sSample;
    public string sGoeStyle;
    public string sGoePerc;
    public string sHemStyle;
    public string sHemPerc;
    public string sJarStyle;
    public string sJarPerc;
    public string sLimStyle;
    public string sLimPerc;
    public string sObservations;
    public int iSKSamplesRockOxides;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();


    public string GCSamplesRockOxides_Add()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(12);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;

            arr[2].ParameterName = "@GoeStyle";
            if (sGoeStyle == null)
                arr[2].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[2].Value = sGoeStyle;

            arr[3].ParameterName = "@GoePerc";
            if (sGoePerc == null)
                arr[3].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[3].Value = sGoePerc;

            arr[4].ParameterName = "@HemStyle";
            if (sHemStyle == null)
                arr[4].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[4].Value = sHemStyle;

            arr[5].ParameterName = "@HemPerc";
            if (sHemPerc == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = sHemPerc;

            arr[6].ParameterName = "@JarStyle";
            if (sJarStyle == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sJarStyle;

            arr[7].ParameterName = "@JarPerc";
            if (sJarPerc == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sJarPerc;

            arr[8].ParameterName = "@LimStyle";
            if (sLimStyle == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = sLimStyle;

            arr[9].ParameterName = "@LimPerc";
            if (sLimPerc == null)
                arr[9].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[9].Value = sLimPerc;

            arr[10].ParameterName = "@Observations";
            if (sObservations == null)
                arr[10].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[10].Value = sObservations;

            arr[11].ParameterName = "@SKSamplesRockOxides";
            arr[11].Value = iSKSamplesRockOxides;

            oRes = oData.ExecuteScalar("usp_GC_SampleRockOxides_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error GCSamplesRockOxides. " + eX.Message); ;
        }
    }

    public string GCSamplesRockOxides_Delete()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);

            arr[0].ParameterName = "@SKSamplesRockOxides";
            arr[0].Value = iSKSamplesRockOxides;

            oRes = oData.ExecuteScalar("usp_GC_SampleRockOxides_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error GCSamplesRockOxides. " + eX.Message); ;
        }
    }

    public DataTable getGCSamplesRockOxidesListReport()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Sample";
            arr[0].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SampleRockOxides_List_Report", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRockOxidesReport: " + eX.Message);
        }
    }

    public DataTable getGCSamplesRockOxidesList()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SampleRockOxides_List", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRockOxides: " + eX.Message);
        }
    }

}

