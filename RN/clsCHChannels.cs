using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsCHChannels
{
    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string sOpcion;
    public string sChId;
    public double? dLenght;
    public double? dEast;
    public double? dNorth;
    public double? dElevation;
    public string sProjection;
    public string sDatum;
    public string sProject;
    public string sClaim;
    public string sStartDate;
    public string sFinalDate;
    //public string sPurpose;
    public string sStorage;
    public string sSource;
    public string sLocation;
    public string sComments;
    public Int64 iSKCHChannels;
    public string sMineID;
    public string sType;
    public string sInstrument;
    public string sDate_Survey;
    public int? iSamplesTotal;


    public string CH_Collars_Add()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(21);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Chid";
            arr[1].Value = sChId;
            
            arr[2].ParameterName = "@Length";
            if (dLenght == null)
                arr[2].Value = System.Data.SqlTypes.SqlString.Null;
            else 
                arr[2].Value = dLenght;

            arr[3].ParameterName = "@East";
            if (dEast == null)
                arr[3].Value = System.Data.SqlTypes.SqlDouble.Null;
            else
                arr[3].Value = dEast;

            arr[4].ParameterName = "@North";
            if (dNorth == null)
                arr[4].Value = System.Data.SqlTypes.SqlDouble.Null;
            else
                arr[4].Value = dNorth;

            arr[5].ParameterName = "@Elevation";
            if (dElevation == null)
                arr[5].Value = System.Data.SqlTypes.SqlDouble.Null;
            else
                arr[5].Value = dElevation;

            arr[6].ParameterName = "@Projection";
            if (sProjection == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sProjection;

            arr[7].ParameterName = "@Datum";
            if (sDatum == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = arr[7].Value = sDatum;

            arr[8].ParameterName = "@Project";
            if (sProject == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = sProject;

            arr[9].ParameterName = "@Claim";
            if (sClaim == null)
                arr[9].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[9].Value = sClaim;

            arr[10].ParameterName = "@Star_Date";
            if (sStartDate == null)
                arr[10].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[10].Value = Convert.ToDateTime(sStartDate);

            arr[11].ParameterName = "@Final_Date";
            if (sFinalDate == null)
                arr[11].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[11].Value = Convert.ToDateTime( sFinalDate);

            arr[12].ParameterName = "@Storage";
            if (sStorage == null)
                arr[12].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[12].Value = sStorage;

            arr[13].ParameterName = "@Source";
            if (sSource == null)
                arr[13].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[13].Value = sSource;

            //arr[14].ParameterName = "@Location";
            //if (sLocation == null)
            //    arr[14].Value = System.Data.SqlTypes.SqlString.Null;
            //else arr[14].Value = sLocation;

            arr[14].ParameterName = "@Comments";
            if (sComments == null)
                arr[14].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[14].Value = sComments;

            arr[15].ParameterName = "@SKCHChannels";
            arr[15].Value = iSKCHChannels;

            arr[16].ParameterName = "@MineID";
            if (sMineID == null)
                arr[16].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[16].Value = sMineID;

            arr[17].ParameterName = "@Type";
            if (sType == null)
                arr[17].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[17].Value = sType;

            arr[18].ParameterName = "@Instrument";
            if (sInstrument == null)
                arr[18].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[18].Value = sInstrument;

            arr[19].ParameterName = "@Date_Survey";
            if (sDate_Survey == null)
                arr[19].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[19].Value = Convert.ToDateTime(sDate_Survey);

            arr[20].ParameterName = "@SamplesTotal";
            if (iSamplesTotal == null)
                arr[20].Value = System.Data.SqlTypes.SqlInt32.Null;
            else arr[20].Value = iSamplesTotal;

            oRes = oData.ExecuteScalar("usp_CH_Collars_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error CH_Collars. " + eX.Message); ;
        }
    }

    public string CH_Collars_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKCHChannels";
            arr[0].Value = iSKCHChannels;
            oRes = oData.ExecuteScalar("usp_CH_Collars_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();
        }
        catch (Exception eX)
        {
            throw new Exception("Delete error CH_Collars. " + eX.Message); ;
        }
    }

    

    public DataTable getCH_Collars()
    {
        try
        {
            DataSet dtCHData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Chid";
            arr[1].Value = sChId;
            dtCHData = oData.ExecuteDataset("usp_CH_Collars_List", arr, CommandType.StoredProcedure);
            return dtCHData.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in CHCollars List: " + eX.Message);
        }
    }


}

