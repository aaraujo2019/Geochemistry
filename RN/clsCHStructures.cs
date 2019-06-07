using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

public  class clsCHStructures
{
    public string sOpcion;
    public string sChid;
    public string sSample;
    public string sType;
    public double? dDip;
    public string sDipAz;
    public double? dAThick;
    public double? dRThick;
    public string sFill;
    public string sFill2;
    public string sFill3;
    public double? dNumber;
    public double? dDensity;
    public string sObservations;
    public int iSKStructures;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string CHStructures_Add()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(15);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;

            arr[2].ParameterName = "@StrType";
            if (sType == null)
                arr[2].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[2].Value = sType;

            arr[3].ParameterName = "@StrDip";
            if (dDip == null)
                arr[3].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[3].Value = dDip;

            arr[4].ParameterName = "@StrDipAz";
            if (sDipAz == null)
                arr[4].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[4].Value = sDipAz;

            arr[5].ParameterName = "@StrAThick";
            if (dAThick == null)
                arr[5].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[5].Value = dAThick;

            arr[6].ParameterName = "@StrRThick";
            if (dRThick == null)
                arr[6].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[6].Value = dRThick;

            arr[7].ParameterName = "@StrFill";
            if (sFill == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sFill;

            arr[8].ParameterName = "@StrNumber";
            if (dNumber == null)
                arr[8].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[8].Value = dNumber;

            arr[9].ParameterName = "@StrDensity";
            if (dDensity == null)
                arr[9].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[9].Value = dDensity;

            arr[10].ParameterName = "@Obsevations";
            if (sObservations == null)
                arr[10].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[10].Value = sObservations;

            arr[11].ParameterName = "@SKStructures";
            arr[11].Value = iSKStructures;

            arr[12].ParameterName = "@Chid";
            arr[12].Value = sChid;



            arr[13].ParameterName = "@StrFill2";
            if (sFill2 == null)
                arr[13].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[13].Value = sFill2;

            arr[14].ParameterName = "@StrFill3";
            if (sFill3 == null)
                arr[14].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[14].Value = sFill3;


            oRes = oData.ExecuteScalar("usp_CH_Structures_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error CHStructures. " + eX.Message); ;
        }
    }

    public string CHstructures_Delete()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);

            arr[0].ParameterName = "@SKStructures";
            arr[0].Value = iSKStructures;

            oRes = oData.ExecuteScalar("usp_CH_Structures_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error CHStructures. " + eX.Message); ;
        }
    }

    public DataTable getCHStructuresList()
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
            dtData = oData.ExecuteDataset("usp_CH_Structures_List", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHStructures: " + eX.Message);
        }
    }

    public DataTable getCHStructuresListReport()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Chid";
            arr[0].Value = sChid;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtData = oData.ExecuteDataset("usp_CH_Structures_List_Report", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHStructures: " + eX.Message);
        }
    }
}

