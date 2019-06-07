using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

public class clsDHAlterations
{
    public string sOpcion;
    public string sHoleID;
    public double dFrom;
    public double dTo;
    public string sA1Type;
    public string sA1Int;
    public string sA1Style;
    public string sA1Min;
    public string sA1Min2;
    public string sA1Min3;
    public string sA2Type;
    public string sA2Int;
    public string sA2Style;
    public string sA2Min;
    public string sA2Min2;
    public string sA2Min3;
    public string sComments;
    public Int64 iSHDHAlterarions;
    public string sA1Style2;
    public string sA2Style2;

    public static string sStaticFrom;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    //[usp_DH_Alterations_List]
    public DataTable getDH_Alterations()
    {
        try
        {

            DataSet dtDH_Alterations = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            dtDH_Alterations = oData.ExecuteDataset("usp_DH_Alterations_List", arr, CommandType.StoredProcedure);
            return dtDH_Alterations.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DH_Alterations: " + eX.Message);
        }
    }

    public DataTable getDHAlterationsValidFromToNext()
    {
        try
        {
            DataSet dtDHAlterationsFromToValid = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@HoleID";
            arr[0].Value = sHoleID;
            dtDHAlterationsFromToValid = oData.ExecuteDataset("usp_DH_Alterations_ListValidFromToNext", arr, CommandType.StoredProcedure);
            return dtDHAlterationsFromToValid.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHAlterationsFromToValid: " + eX.Message);
        }
    }


    public DataTable getDHAlterationsValid()
    {
        try
        {
            DataSet dtDHAlterationsFromToValid = new DataSet();
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@From";
            arr[0].Value = dFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = dTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            arr[3].ParameterName = "@SHDHAlterarions";
            arr[3].Value = iSHDHAlterarions;
            dtDHAlterationsFromToValid = oData.ExecuteDataset("usp_DH_Alterations_ListValid", arr, CommandType.StoredProcedure);
            return dtDHAlterationsFromToValid.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHAlterationsFromToValid: " + eX.Message);
        }
    }

    public DataTable getDHAlterationsFromToValid()
    {
        try
        {
            DataSet dtDHAlterationsFromToValid = new DataSet();
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@From";
            arr[0].Value = dFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = dTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            arr[3].ParameterName = "@SHDHAlterarions";
            arr[3].Value = iSHDHAlterarions;
            dtDHAlterationsFromToValid = oData.ExecuteDataset("usp_DH_Alterations_ListFromToValid", arr, CommandType.StoredProcedure);
            return dtDHAlterationsFromToValid.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHAlterationsFromToValid: " + eX.Message);
        }
    }

    public string DH_Alterations_Add()
    {
        try
        {
            
            object oRes;
            SqlParameter[] arr = oData.GetParameters(20);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            arr[2].ParameterName = "@From";
            arr[2].Value = dFrom;
            arr[3].ParameterName = "@To";
            arr[3].Value = dTo;
            arr[4].ParameterName = "@A1Type";
            arr[4].Value = sA1Type;


            arr[5].ParameterName = "@A1Int";
            if (sA1Int == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = sA1Int;

            arr[6].ParameterName = "@A1Style";
            if (sA1Style == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sA1Style;

            arr[7].ParameterName = "@A1Min";
            if (sA1Min == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sA1Min;

            arr[8].ParameterName = "@A2Type";
            if (sA2Type == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = sA2Type;

            arr[9].ParameterName = "@A2Int";
            if (sA2Int == null)
                arr[9].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[9].Value = sA2Int;
            
            arr[10].ParameterName = "@A2Style";
            if (sA2Style == null)
                arr[10].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[10].Value = sA2Style;

            arr[11].ParameterName = "@A2Min";
            if (sA2Min == null)
                arr[11].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[11].Value = sA2Min;
            
            arr[12].ParameterName = "@Comments";
            if (sComments == null)
                arr[12].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[12].Value = sComments;
            
            arr[13].ParameterName = "@SHDHAlterarions";
            arr[13].Value = iSHDHAlterarions;

            arr[14].ParameterName = "@A1Min2";
            if (sA1Min2 == null)
                arr[14].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[14].Value = sA1Min2;

            arr[15].ParameterName = "@A2Min2";
            if (sA2Min2 == null)
                arr[15].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[15].Value = sA2Min2;            

            arr[16].ParameterName = "@A1Style2";
            if (sA1Style2 == null)
                arr[16].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[16].Value = sA1Style2;

            arr[17].ParameterName = "@A2Style2";
            if (sA2Style2 == null)
                arr[17].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[17].Value = sA2Style2;

            arr[18].ParameterName = "@A1Min3";
            if (sA1Min3 == null)
                arr[18].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[18].Value = sA1Min3;

            arr[19].ParameterName = "@A2Min3";
            if (sA2Min3 == null)
                arr[19].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[19].Value = sA2Min3;

            oRes = oData.ExecuteScalar("usp_DH_Alterations_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error Alterations. " + eX.Message); ;
        }
    }

    public string DH_Alterations_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SHDHAlterarions";
            arr[0].Value = iSHDHAlterarions;
            oRes = oData.ExecuteScalar("usp_DH_Alterations_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error Alterations. " + eX.Message); ;
        }
    }

}

