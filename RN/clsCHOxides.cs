using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;



public class clsCHOxides
{
    public string sOpcion;
    public string sChid;
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
    public int iSKOxides;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string CHOxides_Add()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(13);
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

            arr[11].ParameterName = "@SKOxides";
            arr[11].Value = iSKOxides;

            arr[12].ParameterName = "@Chid";
            arr[12].Value = sChid;

            oRes = oData.ExecuteScalar("usp_CH_Oxides_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error CHOxides. " + eX.Message); ;
        }
    }

    public string CHOxides_Delete()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);

            arr[0].ParameterName = "@SKOxides";
            arr[0].Value = iSKOxides;

            oRes = oData.ExecuteScalar("usp_CH_Oxides_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error CHOxides. " + eX.Message); ;
        }
    }

    public DataTable getCHOxidesList()
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
            dtData = oData.ExecuteDataset("usp_CH_Oxides_List", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHOxides: " + eX.Message);
        }
    }

    public DataTable getCHOxidesListReport()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Chid";
            arr[0].Value = sChid;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtData = oData.ExecuteDataset("usp_CH_Oxides_List_Report", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHOxides: " + eX.Message);
        }
    }

}

