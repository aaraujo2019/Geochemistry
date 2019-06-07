using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

public class clsCHSurveys
{
    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string sOpcion;
    public string sChId;
    public string sSample;
    public double? dTo;
    public double? dAzm;
    public double? dDip;
    public string sMineLocation;
    public string sValidatedby;
    public Int64 iSKCHSurveys;

    public string CH_Surveys_Add()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(9);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Chid";
            arr[1].Value = sChId;
            arr[2].ParameterName = "@Sample";
            arr[2].Value = sSample;
            arr[3].ParameterName = "@To";
            if (dTo == null)
                arr[3].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[3].Value = dTo;

            arr[4].ParameterName = "@Azm";
            if (dAzm == null)
                arr[4].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[4].Value = dAzm;

            arr[5].ParameterName = "@Dip";
            if (dDip == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = dDip;

            arr[6].ParameterName = "@MineLocation";
            if (sMineLocation == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sMineLocation;

            arr[7].ParameterName = "@Validated_by";
            if (sValidatedby == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sValidatedby;

            arr[8].ParameterName = "@SKCHSurveys";
            arr[8].Value = iSKCHSurveys;

            oRes = oData.ExecuteScalar("usp_CH_Surveys_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error CH_Surveys. " + eX.Message); ;
        }
    }

    public string CH_Surveys_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKCHSurveys";
            arr[0].Value = iSKCHSurveys;
            oRes = oData.ExecuteScalar("usp_CH_Surveys_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();
        }
        catch (Exception eX)
        {
            throw new Exception("Delete error CH_Surveys. " + eX.Message); ;
        }
    }

    public DataTable getCH_Surveys()
    {
        try
        {
            DataSet dtCHData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(3);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Chid";
            arr[1].Value = sChId;
            arr[2].ParameterName = "@Sample";
            arr[2].Value = sSample; 
            dtCHData = oData.ExecuteDataset("usp_CH_Surveys_List", arr, CommandType.StoredProcedure);
            return dtCHData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in CHSurveys List: " + eX.Message);
        }
    }


}

