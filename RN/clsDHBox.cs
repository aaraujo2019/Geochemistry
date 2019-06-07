using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsDHBox
{
    public string sOpcion;
    public string sHoleID;
    public double dFrom;
    public double dTo;
    public int iBox;
    public int? iStand;
    public string sColumn;
    public string sRow;
    public Int64 iSKDHBox;
    public int? iPhoto;
    public int? iEditPhoto;

    public static string sStaticFrom;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();


    public DataTable getDH_Box()
    {
        try
        {

            DataSet dtDH_Box = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            dtDH_Box = oData.ExecuteDataset("usp_DH_Box_List", arr, CommandType.StoredProcedure);
            return dtDH_Box.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DH_Box: " + eX.Message);
        }
    }

    public string DH_Box_Add()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(11);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            arr[2].ParameterName = "@From";
            arr[2].Value = dFrom;
            arr[3].ParameterName = "@To";
            arr[3].Value = dTo;
            arr[4].ParameterName = "@Box";
            arr[4].Value = iBox;
            
            arr[5].ParameterName = "@Stand";
            if (iStand == null)
                arr[5].Value = System.Data.SqlTypes.SqlInt32.Null;
            else arr[5].Value = iStand;

            arr[6].ParameterName = "@column";
            if (sColumn == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sColumn;

            arr[7].ParameterName = "@row";
            if (sRow == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sRow;

            arr[8].ParameterName = "@SKDHBox";
            arr[8].Value = iSKDHBox;

            arr[9].ParameterName = "@Photo";
            if (iPhoto == null)
                arr[9].Value = System.Data.SqlTypes.SqlInt16.Null;
            else arr[9].Value = iPhoto;

            arr[10].ParameterName = "@EditPhoto";
            if (iEditPhoto == null)
                arr[10].Value = System.Data.SqlTypes.SqlInt16.Null;
            else arr[10].Value = iEditPhoto;


            oRes = oData.ExecuteScalar("usp_DH_Box_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error Box. " + eX.Message); ;
        }
    }

    public string DH_Box_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKDHBox";
            arr[0].Value = iSKDHBox;

            oRes = oData.ExecuteScalar("usp_DH_Box_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error Box. " + eX.Message); ;
        }
    }

    public DataTable getDHBoxFromToValid()
    {
        try
        {
            DataSet dtDHLitFromToValid = new DataSet();
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@From";
            arr[0].Value = dFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = dTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            arr[3].ParameterName = "@SKDHBox";
            arr[3].Value = iSKDHBox;
            dtDHLitFromToValid = oData.ExecuteDataset("usp_DH_Box_ListFromToValid", arr, CommandType.StoredProcedure);
            return dtDHLitFromToValid.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DHBoxFromToValid: " + eX.Message);
        }
    }

    public DataTable getDHBoxValidExport()
    {
        try
        {
            DataSet dtDHBox = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtDHBox = oData.ExecuteDataset("usp_DH_Box_ListFromToValid_Export", arr, CommandType.StoredProcedure);
            return dtDHBox.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getDHBoxValidExport: " + eX.Message);
        }
    }

}

