using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsDHDensity
{
    public string sOpcion;
    public string sHoleID;
    public string sBox;
    public double dFrom;
    public double dTo;
    public double dLenght;
    public double dDiameter;
    public string sSample;
    public string sLith;
    public string sComments;
    public string sVeinName;
    public string sTexture;
    public string sStructure;
    public string sMineral1;
    public string sMineral2;
    public string sSulfphides;
    public string sAltType;
    public string sAltInt;
    public int iSKDHDensity;


    public static string sStaticFrom;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string DH_Dens_Add()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(19);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            arr[2].ParameterName = "@Box";
            arr[2].Value = sBox;
            arr[3].ParameterName = "@From";
            arr[3].Value = dFrom;
            arr[4].ParameterName = "@To";
            arr[4].Value = dTo;

            arr[5].ParameterName = "@Lenght";
            if (dLenght == null)
                arr[5].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[5].Value = dLenght;

            arr[6].ParameterName = "@Diameter";
            if (dDiameter == null)
                arr[6].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[6].Value = dDiameter;

            arr[7].ParameterName = "@Sample";
            arr[7].Value = sSample;

            arr[8].ParameterName = "@Litho";
            if (sLith == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = sLith;

            arr[9].ParameterName = "@Comments";
            if (sComments == null)
                arr[9].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[9].Value = sComments;

            arr[10].ParameterName = "@VeinName";
            if (sVeinName == null)
                arr[10].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[10].Value = sVeinName;

            arr[11].ParameterName = "@Texture";
            if (sTexture == null)
                arr[11].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[11].Value = sTexture;

            arr[12].ParameterName = "@Structure";
            if (sStructure == null)
                arr[12].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[12].Value = sStructure;

            arr[13].ParameterName = "@Mineralization_1";
            if (sMineral1 == null)
                arr[13].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[13].Value = sMineral1;

            arr[14].ParameterName = "@Mineralization_2";
            if (sMineral2 == null)
                arr[14].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[14].Value = sMineral2;

            arr[15].ParameterName = "@Sulfphides_per";
            if (sSulfphides == null)
                arr[15].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[15].Value = sSulfphides;

            arr[16].ParameterName = "@AltType";
            if (sAltType == null)
                arr[16].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[16].Value = sAltType;

            arr[17].ParameterName = "@AltInt";
            if (sAltInt == null)
                arr[17].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[17].Value = sAltInt;

            arr[18].ParameterName = "@SKDHDensity";
            arr[18].Value = iSKDHDensity;

            oRes = oData.ExecuteScalar("usp_DH_Density_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error Density. " + eX.Message); ;
        }
    }

    public string DH_Dens_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKDHDensity";
            arr[0].Value = iSKDHDensity;
            oRes = oData.ExecuteScalar("usp_DH_Density_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error Density. " + eX.Message); ;
        }
    }

    public DataTable getDHDensFromToValid()
    {
        try
        {
            DataSet dtDHDensFromToValid = new DataSet();
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@From";
            arr[0].Value = dFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = dTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            arr[3].ParameterName = "@SKDHDensity";
            arr[3].Value = iSKDHDensity;
            dtDHDensFromToValid = oData.ExecuteDataset("usp_DH_Density_ListFromToValid", arr, CommandType.StoredProcedure);
            return dtDHDensFromToValid.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHDensFromToValid: " + eX.Message);
        }
    }

    public DataTable getDHDensity()
    {
        try
        {
            DataSet dtDHDensFromToValid = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            dtDHDensFromToValid = oData.ExecuteDataset("usp_DH_Density_List", arr, CommandType.StoredProcedure);
            return dtDHDensFromToValid.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHDensity: " + eX.Message);
        }
    }


    public string sOpcionM;
    public int iSKDHDensityMethod;
    public string sLab;
    public double dDrySamp;
    public double dImmersedSamp;
    public double dDensity;
    public string sMethod;
    public int iPriority;


    public string DH_DensMethod_Add()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(9);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcionM;
            arr[1].ParameterName = "@SKDHDensityMethod";
            arr[1].Value = iSKDHDensityMethod;
            arr[2].ParameterName = "@FK_Density";
            arr[2].Value = iSKDHDensity;
            arr[3].ParameterName = "@Lab";
            arr[3].Value = sLab;
            arr[4].ParameterName = "@DrySamp_g";
            arr[4].Value = dDrySamp;
            arr[5].ParameterName = "@ImmersedSamp_g";
            arr[5].Value = dImmersedSamp;
            arr[6].ParameterName = "@Density";
            arr[6].Value = dDensity;
            arr[7].ParameterName = "@Method";
            arr[7].Value = sMethod;
            arr[8].ParameterName = "@Priority";
            arr[8].Value = iPriority;

            oRes = oData.ExecuteScalar("usp_DH_DensityMethod_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error Density. " + eX.Message); ;
        }
    }

    public DataTable getDHDensityMethod()
    {
        try
        {
            DataSet dtDHDens = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcionM;
            arr[1].ParameterName = "@FK_Density";
            arr[1].Value = iSKDHDensity;
            dtDHDens = oData.ExecuteDataset("usp_DH_DensityMethod_List", arr, CommandType.StoredProcedure);
            return dtDHDens.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHDensityMethod: " + eX.Message);
        }
    }

    public string DH_DensMethod_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKDHDensityMethod";
            arr[0].Value = iSKDHDensityMethod;
            oRes = oData.ExecuteScalar("usp_DH_DensityMethod_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error DensityMethod. " + eX.Message); ;
        }
    }
}

