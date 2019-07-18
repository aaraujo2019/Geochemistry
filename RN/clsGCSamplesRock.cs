using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

public class clsGCSamplesRock
{
    #region properties

    public string sOpcion;
    public string sSample;
    public string sTarget;
    public string sLocation;
    public string sProject;
    public string sGeologist;
    public string sHelper;
    public string sStation;
    public string sDate;
    public double? dCoordE;
    public double? dCoordN;
    public double? dCoordZ;
    public string sCs;
    public double? dGPSepe;
    public string sPhoto;
    public string sPhoto_Azimuth;
    public string sSampleType;
    public string sNotInSitu;
    public string sPorpouse;
    public string sRelativeLoc;
    public double? dLenght;
    public double? dHigh;
    public string sThickness;
    public string sObservations;
    public string sLRock;
    public string sLTexture;
    public string sLGSize;
    public string sLWeathering;
    public string sLRockSorting;
    public string sLRockSphericity;
    public string sLRockRounding;
    public string sLRockObservation;
    public string sLMatrixPerc;
    public string sLMatrixGSize;
    public string sLMatrixObservations;
    public string sLPhenoCPerc;
    public string sLPhenoCGSize;
    public string sLPhenoCObservations;
    public string sVContactType;
    public string sVVeinName;
    public string sVHostRock;
    public string sVObservations;
    public int? iSKSamplesRock;
    public string sSamplingType, sDupOf;
    public string sMine;

    #endregion

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string GCSamplesRock_Add()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(46);
            
            //@Opcion varchar(2) ,
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            //@Sample varchar(20),
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            //@Target varchar(30),
            arr[2].ParameterName = "@Target";
            if (sTarget == null)
                arr[2].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[2].Value = sTarget;
            //@Location varchar(255),
            arr[3].ParameterName = "@Location";
            if (sLocation == null)
                arr[3].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[3].Value = sLocation;
            //@Project varchar(20),
            arr[4].ParameterName = "@Project";
            if (sProject == null)
                arr[4].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[4].Value = sProject;
            //@Geologist varchar(30),
            arr[5].ParameterName = "@Geologist";
            if (sGeologist == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = sGeologist;
            //@Helper varchar(3),
            arr[6].ParameterName = "@Helper";
            if (sHelper == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sHelper;
            //@Station varchar(20),
            arr[7].ParameterName = "@Station";
            if (sStation == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sStation;
            //@Date varchar(10),
            arr[8].ParameterName = "@Date";
            if (sDate == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = Convert.ToDateTime(sDate);
            //@E numeric(18,3),
            arr[9].ParameterName = "@E";
            if (dCoordE == null)
                arr[9].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[9].Value = dCoordE;
            //@N numeric(18,3),
            arr[10].ParameterName = "@N";
            if (dCoordN == null)
                arr[10].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[10].Value = dCoordN;
            //@Z numeric(18,3),
            arr[11].ParameterName = "@Z";
            if (dCoordZ == null)
                arr[11].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[11].Value = dCoordZ;
            //@CS varchar(10),
            arr[12].ParameterName = "@CS";
            if (sCs == null)
                arr[12].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[12].Value = sCs;
            //@GPSepe numeric(18,3),
            arr[13].ParameterName = "@GPSepe";
            if (dGPSepe == null)
                arr[13].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[13].Value = dGPSepe;
            //@Photo varchar(25),
            arr[14].ParameterName = "@Photo";
            if (sPhoto == null)
                arr[14].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[14].Value = sPhoto;
            //@Photo_azimuth varchar(25),
            arr[15].ParameterName = "@Photo_azimuth";
            if (sPhoto_Azimuth == null)
                arr[15].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[15].Value = sPhoto_Azimuth;
            //@SampleType varchar(22),
            arr[16].ParameterName = "@SampleType";
            if (sSampleType == null)
                arr[16].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[16].Value = sSampleType;
            //@NotItSitu varchar(22),
            arr[17].ParameterName = "@NotItSitu";
            if (sNotInSitu == null)
                arr[17].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[17].Value = sNotInSitu;
            //@Porpuose varchar(20),
            arr[18].ParameterName = "@Porpuose";
            if (sPorpouse == null)
                arr[18].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[18].Value = sPorpouse;
            //@Relative_Loc varchar(20),
            arr[19].ParameterName = "@Relative_Loc";
            if (sRelativeLoc == null)
                arr[19].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[19].Value = sRelativeLoc;
            //@length numeric(18,3),
            arr[20].ParameterName = "@length";
            if (dLenght == null)
                arr[20].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[20].Value = dLenght;
            //@High varchar(20),
            arr[21].ParameterName = "@High";
            if (dHigh == null)
                arr[21].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[21].Value = dHigh;
            //@Thickness varchar(20),
            arr[22].ParameterName = "@Thickness";
            if (sThickness == null)
                arr[22].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[22].Value = sThickness;
            //@Obsevations varchar(300),
            arr[23].ParameterName = "@Obsevations";
            if (sObservations == null)
                arr[23].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[23].Value = sObservations;
            //@LRock varchar(4),
            arr[24].ParameterName = "@LRock";
            if (sLRock == null)
                arr[24].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[24].Value = sLRock;
            //@LTexture varchar(3),
            arr[25].ParameterName = "@LTexture";
            if (sLTexture == null)
                arr[25].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[25].Value = sLTexture;
            //@LGSize varchar(3),
            arr[26].ParameterName = "@LGSize";
            if (sLGSize == null)
                arr[26].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[26].Value = sLGSize;
            //@LWeathering varchar(50),
            arr[27].ParameterName = "@LWeathering";
            if (sLWeathering == null)
                arr[27].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[27].Value = sLWeathering;
            //@LRocksSorting varchar(3),
            arr[28].ParameterName = "@LRocksSorting";
            if (sLRockSorting == null)
                arr[28].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[28].Value = sLRockSorting;
            //@LRocksSphercity varchar(3),
            arr[29].ParameterName = "@LRocksSphericity";
            if (sLRockSphericity == null)
                arr[29].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[29].Value = sLRockSphericity;
            //@LRocksRounding varchar(3),
            arr[30].ParameterName = "@LRocksRounding";
            if (sLRockRounding == null)
                arr[30].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[30].Value = sLRockRounding;
            //@LRocksObservation varchar(300),
            arr[31].ParameterName = "@LRocksObservation";
            if (sLRockObservation == null)
                arr[31].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[31].Value = sLRockObservation;
            //@LMatrixPerc varchar(3),
            arr[32].ParameterName = "@LMatrixPerc";
            if (sLMatrixPerc == null)
                arr[32].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[32].Value = sLMatrixPerc;
            //@LMatrixGSize varchar(3),
            arr[33].ParameterName = "@LMatrixGSize";
            if (sLMatrixGSize == null)
                arr[33].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[33].Value = sLMatrixGSize;
            //@LMatrixObsevations varchar(300),
            arr[34].ParameterName = "@LMatrixObsevations";
            if (sLMatrixObservations == null)
                arr[34].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[34].Value = sLMatrixObservations;
            //@LPhenoCPerc numeric(18,2),
            arr[35].ParameterName = "@LPhenoCPerc";
            if (sLPhenoCPerc == null)
                arr[35].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[35].Value = sLPhenoCPerc;
            //@LPhenoCGSize varchar(3),
            arr[36].ParameterName = "@LPhenoCGSize";
            if (sLPhenoCGSize == null)
                arr[36].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[36].Value = sLPhenoCGSize;
            //@LPhenoCObsevations varchar(300),
            arr[37].ParameterName = "@LPhenoCObsevations";
            if (sLPhenoCObservations == null)
                arr[37].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[37].Value = sLPhenoCObservations;
            //@VContactType varchar(70),
            arr[38].ParameterName = "@VContactType";
            if (sVContactType == null)
                arr[38].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[38].Value = sVContactType;
            //@VVeinName varchar(30),
            arr[39].ParameterName = "@VVeinName";
            if (sVVeinName == null)
                arr[39].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[39].Value = sVVeinName;
            //@VHostRock varchar(30),
            arr[40].ParameterName = "@VHostRock";
            if (sVHostRock == null)
                arr[40].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[40].Value = sVHostRock;
            //@VObsevations varchar(300),
            arr[41].ParameterName = "@VObsevations";
            if (sVObservations == null)
                arr[41].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[41].Value = sVObservations;
            //@SKSamplesRock int
            arr[42].ParameterName = "@SKSamplesRock";
            arr[42].Value = iSKSamplesRock;

            arr[43].ParameterName = "@SamplingType";
            if (sSamplingType == null)
                arr[43].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[43].Value = sSamplingType;

            arr[44].ParameterName = "@DupOf";
            if (sDupOf == null)
                arr[44].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[44].Value = sDupOf;

            arr[45].ParameterName = "@Mine";
            if (sMine == null)
                arr[45].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[45].Value = sMine; 

            oRes = oData.ExecuteScalar("usp_GC_SamplesRock_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error GCSamplesRock. " + eX.Message); ;
        }
    }

    public string GCSamplesRock_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKSamplesRock";
            arr[0].Value = iSKSamplesRock;

            oRes = oData.ExecuteScalar("usp_GC_SamplesRock_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Delete error GCSamplesRock. " + eX.Message); ;
        }
    }

    public DataTable getGCSamplesRockListAll()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SamplesRock_List", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRockAll: " + eX.Message);
        }
    }

    public DataTable getGCSamplesRockListBySampleReport()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Sample";
            arr[0].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SamplesRock_ListSampleReport", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRockBySample: " + eX.Message);
        }
    }


    public DataTable getGCSamplesRockList_Sample()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Sample";
            arr[0].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SamplesRock_ListSample", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRock_Sample: " + eX.Message);
        }
    }

   
}

