using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


    public class clsDH_Weathering
    {

        public string sOpcion;
        public string sHoleID;
        public double dFrom;
        public double dTo;
        public string sWeathering;
        public double? dOxidation;
        public string sColour1;
        public string sSufix1;
        public string sColour2;
        public string sSufix2;
        public string sObservation;

        public string sMineral1;
        public string sMineral2;
        public string sMineral3;
        public string sMineral4;

        public Int64 iDHWeatheringID;

        public static string sStaticFrom;

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

        public DataTable getDH_Weathering()
        {
            try
            {

                DataSet dtDH_Weathering = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                dtDH_Weathering = oData.ExecuteDataset("usp_DH_Weathering_List", arr, CommandType.StoredProcedure);
                return dtDH_Weathering.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DH_Weathering: " + eX.Message);
            }
        }

        public string DH_Weathering_Add()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(16);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                arr[2].ParameterName = "@From";
                arr[2].Value = dFrom;
                arr[3].ParameterName = "@To";
                arr[3].Value = dTo;
                arr[4].ParameterName = "@Weathering";
                arr[4].Value = sWeathering;
                
                arr[5].ParameterName = "@Oxidation";
                if (dOxidation == null)
                    arr[5].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[5].Value = dOxidation;

                arr[6].ParameterName = "@Colour1";
                if (sColour1 == null)
                    arr[6].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[6].Value = sColour1;

                arr[7].ParameterName = "@Sufix1";
                if (sSufix1 == null)
                    arr[7].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[7].Value = sSufix1;

                arr[8].ParameterName = "@Colour2";
                if (sColour2 == null)
                    arr[8].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[8].Value = sColour2;

                arr[9].ParameterName = "@Sufix2";
                if (sSufix2 == null)
                    arr[9].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[9].Value = sSufix2;

                arr[10].ParameterName = "@Observation";
                if (sObservation == null)
                    arr[10].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[10].Value = sObservation;

                arr[11].ParameterName = "@DHWeatheringID";
                arr[11].Value = iDHWeatheringID;

                arr[12].ParameterName = "@Mineral1";
                if (sMineral1 == null)
                    arr[12].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[12].Value = sMineral1;

                arr[13].ParameterName = "@Mineral2";
                if (sMineral2 == null)
                    arr[13].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[13].Value = sMineral2;

                arr[14].ParameterName = "@Mineral3";
                if (sMineral3 == null)
                    arr[14].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[14].Value = sMineral3;

                arr[15].ParameterName = "@Mineral4";
                if (sMineral4 == null)
                    arr[15].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[15].Value = sMineral4;
                
                oRes = oData.ExecuteScalar("usp_DH_Weathering_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Save error Weathering. " + eX.Message); ;
            }
        }

        //[usp_DH_Weathering_Delete]
        public string DH_Weathering_Delete()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                arr[1].ParameterName = "@From";
                arr[1].Value = dFrom;

                oRes = oData.ExecuteScalar("usp_DH_Weathering_Delete", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Delete error Weathering. " + eX.Message); ;
            }
        }

        //[usp_DH_Weathering_ListValidFromToNext]
        public DataTable getDHWeatValidFromToNext()
        {
            try
            {
                DataSet dtDHLitFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHLitFromToValid = oData.ExecuteDataset("usp_DH_Weathering_ListValidFromToNext", arr, CommandType.StoredProcedure);
                return dtDHLitFromToValid.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHWeatFromToValid: " + eX.Message);
            }
        }

        public DataTable getDHWeatValid()
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
                arr[3].ParameterName = "@SKDHWeathering";
                arr[3].Value = iDHWeatheringID;
                dtDHLitFromToValid = oData.ExecuteDataset("usp_DH_Weathering_ListValid", arr, CommandType.StoredProcedure);
                return dtDHLitFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHWeatFromToValid: " + eX.Message);
            }
        }

        public DataTable getDHWeatFromToValid()
        {
            try
            {
                DataSet dtDHLitFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(3);
                arr[0].ParameterName = "@From";
                arr[0].Value = dFrom;
                arr[1].ParameterName = "@To";
                arr[1].Value = dTo;
                arr[2].ParameterName = "@HoleID";
                arr[2].Value = sHoleID;
                //arr[3].ParameterName = "";
                //arr[3].Value = sOpcion;
                dtDHLitFromToValid = oData.ExecuteDataset("usp_DH_Weathering_ListFromToValid", arr, CommandType.StoredProcedure);
                return dtDHLitFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHWeatFromToValid: " + eX.Message);
            }
        }

    }

