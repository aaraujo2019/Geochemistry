using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsCHSamples
{
    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public string sOpcion;
    public string sChId;
    public string sSample;
    public double? dFrom;
    public double? dTo;
    public string sTarget;
    //public string sLocation;
    public string sProject;
    public string sGeologist;
    public string sHelper;
    public string sStation;
    public string sDate;
    public double? dE, dN, dZ;
    public double? dE2, dN2, dZ2;
    public string sCS;
    public double? dGPSEpe;
    public string sPhoto;
    public string sPhotoAzimuth;
    public string sSampleType;
    public string sSamplingType;
    public string sDupOf;
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
    public double? sLMatrixPerc;
    public string sLMatrixGSize;
    public string sLMatrixObservations;
    public double? sLPhenoCPerc;
    public string sLPhenoCGSize;
    public string sLPhenoCObservations;
    public string sVContactType;
    public string sVVeinName;
    public string sVHostRock;
    public string sVObservations;
    public Int64 iSKCHSamples;
    public string sMine, sMineEntrance;
    public bool bValited;
    public int? iSampleCont;

    public string CH_Samples_Add()
    {
        try
        {
            
            object oRes;
            SqlParameter[] arr = oData.GetParameters(52);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@Chid";
            arr[1].Value = sChId;
            arr[2].ParameterName = "@Sample";
            if (sSample == null)
                arr[2].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[2].Value = sSample;

            arr[3].ParameterName = "@From";
            if (dFrom == null)
                arr[3].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[3].Value = dFrom;

            arr[4].ParameterName = "@To";
            if (dTo == null)
                arr[4].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[4].Value = dTo;

            arr[5].ParameterName = "@Target";
            if (sTarget == null)
                arr[5].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[5].Value = sTarget;

            //arr[6].ParameterName = "@Location";
            //if (sLocation == null)
            //    arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            //else arr[6].Value = sLocation;

            //arr[6].ParameterName = "@MineEntrance";
            //if (sMineEntrance == null)
            //    arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            //else arr[6].Value = sMineEntrance; 

            arr[6].ParameterName = "@Project";
            if (sProject == null)
                arr[6].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[6].Value = sProject;

            arr[7].ParameterName = "@Geologist";
            if (sGeologist == null)
                arr[7].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[7].Value = sGeologist;

            arr[8].ParameterName = "@Helper";
            if (sHelper == null)
                arr[8].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[8].Value = sHelper;

            arr[9].ParameterName = "@Station";
            if (sStation == null)
                arr[9].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[9].Value = sStation;

            arr[10].ParameterName = "@Date";
            if (sDate == null)
                arr[10].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[10].Value = sDate;

            arr[11].ParameterName = "@E";
            if (dE == null)
                arr[11].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[11].Value = dE;

            arr[12].ParameterName = "@N";
            if (dN == null)
                arr[12].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[12].Value = dN;

            arr[13].ParameterName = "@Z";
            if (dZ == null)
                arr[13].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[13].Value = dZ;

            arr[14].ParameterName = "@CS";
            if (sCS == null)
                arr[14].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[14].Value = sCS;

            arr[15].ParameterName = "@GPSepe";
            if (dGPSEpe == null)
                arr[15].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[15].Value = dGPSEpe;

            arr[16].ParameterName = "@Photo";
            if (sPhoto == null)
                arr[16].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[16].Value = sPhoto;

            arr[17].ParameterName = "@Photo_azimuth";
            if (sPhotoAzimuth == null)
                arr[17].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[17].Value = sPhotoAzimuth;

            arr[18].ParameterName = "@SampleType";
            if (sSampleType == null)
                arr[18].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[18].Value = sSampleType;

            arr[19].ParameterName = "@SamplingType";
            if (sSamplingType == null)
                arr[19].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[19].Value = sSamplingType;

            arr[20].ParameterName = "@DupOf";
            if (sDupOf == null)
                arr[20].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[20].Value = sDupOf;

            arr[21].ParameterName = "@NotItSitu";
            if (sNotInSitu == null)
                arr[21].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[21].Value = sNotInSitu;

            arr[22].ParameterName = "@Porpuose";
            if (sPorpouse == null)
                arr[22].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[22].Value = sPorpouse;

            arr[23].ParameterName = "@Relative_Loc";
            if (sRelativeLoc == null)
                arr[23].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[23].Value = sRelativeLoc;

            arr[24].ParameterName = "@length";
            if (dLenght == null)
                arr[24].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[24].Value = dLenght;

            arr[25].ParameterName = "@High";
            if (dHigh == null)
                arr[25].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[25].Value = dHigh;

            arr[26].ParameterName = "@Thickness";
            if (sThickness == null)
                arr[26].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[26].Value = sThickness;

            arr[27].ParameterName = "@Obsevations";
            if (sObservations == null)
                arr[27].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[27].Value = sObservations;

            arr[28].ParameterName = "@LRock";
            if (sLRock == null)
                arr[28].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[28].Value = sLRock;

            arr[29].ParameterName = "@LTexture";
            if (sLTexture == null)
                arr[29].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[29].Value = sLTexture;

            arr[30].ParameterName = "@LGSize";
            if (sLGSize == null)
                arr[30].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[30].Value = sLGSize;

            arr[31].ParameterName = "@LWeathering";
            if (sLWeathering == null)
                arr[31].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[31].Value = sLWeathering;

            arr[32].ParameterName = "@LRocksSorting";
            if (sLRockSorting == null)
                arr[32].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[32].Value = sLRockSorting;

            arr[33].ParameterName = "@LRocksSphericity";
            if (sLRockSphericity == null)
                arr[33].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[33].Value = sLRockSphericity;

            arr[34].ParameterName = "@LRocksRounding";
            if (sLRockRounding == null)
                arr[34].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[34].Value = sLRockRounding;

            arr[35].ParameterName = "@LRocksObservation";
            if (sLRockObservation == null)
                arr[35].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[35].Value = sLRockObservation;

            arr[36].ParameterName = "@LMatrixPerc";
            if (sLMatrixPerc == null)
                arr[36].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[36].Value = sLMatrixPerc;

            arr[37].ParameterName = "@LMatrixGSize";
            if (sLMatrixGSize == null)
                arr[37].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[37].Value = sLMatrixGSize;

            arr[38].ParameterName = "@LMatrixObsevations";
            if (sLMatrixObservations == null)
                arr[38].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[38].Value = sLMatrixObservations;

            arr[39].ParameterName = "@LPhenoCPerc";
            if (sLPhenoCPerc == null)
                arr[39].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[39].Value = sLPhenoCPerc;

            arr[40].ParameterName = "@LPhenoCGSize";
            if (sLPhenoCGSize == null)
                arr[40].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[40].Value = sLPhenoCGSize;

            arr[41].ParameterName = "@LPhenoCObsevations";
            if (sLPhenoCObservations == null)
                arr[41].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[41].Value = sLPhenoCObservations;

            arr[42].ParameterName = "@VContactType";
            if (sVContactType == null)
                arr[42].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[42].Value = sVContactType;

            arr[43].ParameterName = "@VVeinName";
            if (sVVeinName == null)
                arr[43].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[43].Value = sVVeinName;

            arr[44].ParameterName = "@VHostRock";
            if (sVHostRock == null)
                arr[44].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[44].Value = sVHostRock;

            arr[45].ParameterName = "@VObsevations";
            if (sVObservations == null)
                arr[45].Value = System.Data.SqlTypes.SqlString.Null;
            else arr[45].Value = sVObservations;

            arr[46].ParameterName = "@SKCHSamples";
            arr[46].Value = iSKCHSamples;



            arr[47].ParameterName = "@E2";
            if (dE2 == null)
                arr[47].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[47].Value = dE2;

            arr[48].ParameterName = "@N2";
            if (dN2 == null)
                arr[48].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[48].Value = dN2;

            arr[49].ParameterName = "@Z2";
            if (dZ2 == null)
                arr[49].Value = System.Data.SqlTypes.SqlDouble.Null;
            else arr[49].Value = dZ2;

            arr[50].ParameterName = "@Validated";
            if (bValited == null)
                arr[50].Value = System.Data.SqlTypes.SqlBoolean.Null;
            else arr[50].Value = bValited;

            arr[51].ParameterName = "@SampleCont";
            if (iSampleCont == null)
                arr[51].Value = System.Data.SqlTypes.SqlInt32.Null;
            else arr[51].Value = iSampleCont;
            

            oRes = oData.ExecuteScalar("usp_CH_Samples_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error CH_Samples. " + eX.Message); ;
        }
    }

    public string CH_Samples_Delete()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKCHSamples";
            arr[0].Value = iSKCHSamples;
            oRes = oData.ExecuteScalar("usp_CH_Samples_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();
        }
        catch (Exception eX)
        {
            throw new Exception("Delete error CH_Samples. " + eX.Message); ;
        }
    }

    public DataTable getCHSamplesBySample()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Chid";
            arr[0].Value = sChId;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_CH_Samples_ListBySample", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRock: " + eX.Message);
        }
    }

    public DataTable getCHSamplesBySampleReport()
    {
        try
        {
            DataSet dtGCSamplesRock = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Chid";
            arr[0].Value = sChId;
            arr[1].ParameterName = "@Sample";
            arr[1].Value = sSample;
            dtGCSamplesRock = oData.ExecuteDataset("usp_CH_Samples_ListBySampleReport", arr, CommandType.StoredProcedure);
            return dtGCSamplesRock.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRock: " + eX.Message);
        }
    }
    

    /// <summary>
    /// Procedimiento para poblar el combo de lithology
    /// </summary>
    /// <returns></returns>
    public DataTable getGCSamplesRockList_Sample()
    {
        try
        {
            DataSet dtCH = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Sample";
            arr[0].Value = sSample;
            dtCH = oData.ExecuteDataset("usp_GC_SamplesRock_ListSample", arr, CommandType.StoredProcedure);
            return dtCH.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getGCSamplesRock_Sample: " + eX.Message);
        }
    }

    public DataTable getCHSamplesByChid()
    {
        try
        {
            DataSet dtCHSamp = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@ChId";
            arr[1].Value = sChId;
            dtCHSamp = oData.ExecuteDataset("usp_CH_Samples_List", arr, CommandType.StoredProcedure);
            return dtCHSamp.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in getCHSamplesByChid: " + eX.Message);
        }
    }


}

