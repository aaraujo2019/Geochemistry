using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsCHMineralizations
{
    public string sOpcion;
    public string sChid;
    public string sSample;
    public string sMineral;
    public string sMinStyle;
    public double? dMinPerc;
    public string sObservations;
    public int iSKMineralizations;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string CHMineralizations_Add()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(8);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Chid";
            arr[1].Value = sChid;
            arr[2].ParameterName = "@Sample";
            arr[2].Value = sSample;

            arr[3].ParameterName = "@MZMin";
            if (sMineral == null)
                arr[3].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[3].Value = sMineral;

            arr[4].ParameterName = "@MZStyle";
            if (sMinStyle == null)
                arr[4].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[4].Value = sMinStyle;

            arr[5].ParameterName = "@MZPerc";
            if (dMinPerc == null)
                arr[5].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[5].Value = dMinPerc;

            arr[6].ParameterName = "@Obsevations";
            if (sObservations == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sObservations;

            arr[7].ParameterName = "@SKMineralizations";
            arr[7].Value = iSKMineralizations;

            oRes = oData.ExecuteScalar("usp_CH_Mineralizations_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error CHMineralizations. " + eX.Message); ;
        }
    }

    public string CHMineralizations_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKMineralizations";
            arr[0].Value = iSKMineralizations;

            oRes = oData.ExecuteScalar("usp_CH_Mineralizations_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error CHMineralizations. " + eX.Message); ;
        }
    }

    public DataTable getCHMineralizationsList()
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
            dtData = oData.ExecuteDataset("usp_CH_Mineralizations_List", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHMineralizations: " + eX.Message);
        }
    }

    public DataTable getCHMineralizationsListReport()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Chid";
            arr[0].Value = sChid;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtData = oData.ExecuteDataset("usp_CH_Mineralizations_ListReport", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHMineralizations: " + eX.Message);
        }
    }


}

