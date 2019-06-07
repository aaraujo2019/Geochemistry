using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

public class clsDHOxides
{
    public string sOpcion;
    public string sHoleID;
    public double dFrom;
    public double dTo;
    public double? dHem;
    public string sGt;
    public string sJar;
    public string sLim;
    public string sCuO;
    public string smmnox_Perc;
    public double? dOther;
    public double? dOtherGr;
    public string sDist;
    public string sIntensity;
    public Int64 iSKDHOxides;

    public static string sStaticFrom;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public DataTable getDHOxides()
    {
        try
        {

            DataSet dtDH_Oxides = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            dtDH_Oxides = oData.ExecuteDataset("usp_DH_Structures_List", arr, CommandType.StoredProcedure);
            return dtDH_Oxides.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DH_Oxides: " + eX.Message);
        }
    }

    public string DH_Oxides_Add()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(15);

            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            arr[2].ParameterName = "@From";
            arr[2].Value = dFrom;
            arr[3].ParameterName = "@To";
            arr[3].Value = dTo;
            arr[4].ParameterName = "@Hem";

            if (dHem == null)
                arr[4].Value = System.Data.SqlTypes.SqlDouble.Null;
            else
                arr[4].Value = dHem;

            arr[5].ParameterName = "@Gt";
            if (sGt == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = sGt;

            arr[6].ParameterName = "@Jar";
            if(sJar == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sJar;

            arr[7].ParameterName = "@Lim";
            if(sLim == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sLim;

            arr[8].ParameterName = "@CuO";
            if(sCuO == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = sCuO;

            arr[9].ParameterName = "@mmnox_Perc";
            if(smmnox_Perc == null)
                arr[9].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[9].Value = smmnox_Perc;

            arr[10].ParameterName = "@Other";
            if(dOther == null)
                arr[10].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[10].Value = dOther;

            arr[11].ParameterName = "@OtherGr";
            if (dOtherGr == null)
                arr[11].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[11].Value = dOtherGr;

            arr[12].ParameterName = "@Dist";
            if (sDist == null)
                arr[12].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[12].Value = sDist;
            

            arr[13].ParameterName = "@Intensity";
            if (sIntensity == null)
                arr[13].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[13].Value = sIntensity;
            
            
            arr[14].ParameterName = "@SKDHOxides";
            arr[14].Value = iSKDHOxides;

            oRes = oData.ExecuteScalar("usp_DH_Oxides_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error Oxidation. " + eX.Message); ;
        }
    }

    public string DH_Oxides_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKDHOxides";
            arr[0].Value = iSKDHOxides;

            oRes = oData.ExecuteScalar("usp_DH_Oxides_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error Oxidation. " + eX.Message); ;
        }
    }

    public DataTable getDHOxidesFromToValid()
    {
        try
        {
            DataSet dtDHOxFromToValid = new DataSet();
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@From";
            arr[0].Value = dFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = dTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            arr[3].ParameterName = "@SKDHOxides";
            arr[3].Value = iSKDHOxides;
            dtDHOxFromToValid = oData.ExecuteDataset("usp_DH_Oxides_ListFromToValid", arr, CommandType.StoredProcedure);
            return dtDHOxFromToValid.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in dtDHOxFromToValid: " + eX.Message);
        }
    }


}

