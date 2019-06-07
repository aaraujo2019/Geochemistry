using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;



    public class clsDHGeotech
    {

        public string sOpcion;
        public string sHoleID;
        public double iFrom;
        public double iTo;
        public string sLithCod;
        public double? dRecm;
        public double? dRQDcm;
        public double? dNoOfFract;
        public double? dJoinCond;
        public double? dJn;
        public double? dJr;
        public double? dJa;
        public string sDegBreak;
        public string sAltWeath;
        public string sHardness;
        public string sComments;
        public Int64 iDHGeotechID;

        public static string sStaticFrom;

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

        public DataTable getDH_Geotech()
        {
            try
            {

                DataSet dtDH_Geotech = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                dtDH_Geotech = oData.ExecuteDataset("usp_DH_Geotech_List", arr, CommandType.StoredProcedure);
                return dtDH_Geotech.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DH_Geotech: " + eX.Message);
            }
        }

        public DataTable getDHGeotechValidFromToNext()
        {
            try
            {
                DataSet dtDHGeotechFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHGeotechFromToValid = oData.ExecuteDataset("usp_DH_Geotech_FromToNext", arr, CommandType.StoredProcedure);
                return dtDHGeotechFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHGeotechValidFromToNext: " + eX.Message);
            }
        }

        //[usp_DH_GeoTech_ListValid]
        public DataTable getDHGeotechValid()
        {
            try
            {
                DataSet dtDHGeotechFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(4);
                arr[0].ParameterName = "@From";
                arr[0].Value = iFrom;
                arr[1].ParameterName = "@To";
                arr[1].Value = iTo;
                arr[2].ParameterName = "@HoleID";
                arr[2].Value = sHoleID;
                arr[3].ParameterName = "@SKDHGeotech";
                arr[3].Value = iDHGeotechID;
                dtDHGeotechFromToValid = oData.ExecuteDataset("usp_DH_GeoTech_ListValid", arr, CommandType.StoredProcedure);
                return dtDHGeotechFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHGeotechValid: " + eX.Message);
            }
        }

        public DataTable getDHGeotechFromToValid()
        {
            try
            {
                DataSet dtDHGeotechFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(4);
                arr[0].ParameterName = "@From";
                arr[0].Value = iFrom;
                arr[1].ParameterName = "@To";
                arr[1].Value = iTo;
                arr[2].ParameterName = "@HoleID";
                arr[2].Value = sHoleID;
                arr[3].ParameterName = "@SKDHGeotech";
                arr[3].Value = iDHGeotechID; 
                dtDHGeotechFromToValid = oData.ExecuteDataset("usp_DH_GeoTech_ListFromToValid", arr, CommandType.StoredProcedure);
                return dtDHGeotechFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHGeotechFromToValid: " + eX.Message);
            }
        }

        public string DH_Geotech_Delete()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                arr[1].ParameterName = "@From";
                arr[1].Value = iFrom;

                oRes = oData.ExecuteScalar("usp_DH_Geotech_Delete", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Delete error Geotech. " + eX.Message); ;
            }
        }

        public string DH_Geotech_Add()
        {
            try
            {
 
                object oRes;
                SqlParameter[] arr = oData.GetParameters(17);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                arr[2].ParameterName = "@From";
                arr[2].Value = iFrom;
                arr[3].ParameterName = "@To";
                arr[3].Value = iTo;
                
                arr[4].ParameterName = "@LithCod";
                if (sLithCod == null)
                    arr[4].Value = System.Data.SqlTypes.SqlString.Null;
                else
                    arr[4].Value = sLithCod;
                
                arr[5].ParameterName = "@Recm";
                if (dRecm == null)
                    arr[5].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[5].Value = dRecm;
                

                arr[6].ParameterName = "@RQDcm";
                if (dRQDcm == null)
                    arr[6].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[6].Value = dRQDcm;

               
                arr[7].ParameterName = "@NoOfFract";
                if (dNoOfFract == null)
                    arr[7].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[7].Value = dNoOfFract;

                arr[8].ParameterName = "@JointCond";
                if (dJoinCond == null)
                    arr[8].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[8].Value = dJoinCond;

                arr[9].ParameterName = "@Jn";
                if (dJn == null)
                    arr[9].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[9].Value = dJn;

                arr[10].ParameterName = "@Jr";
                if (dJr == null)
                    arr[10].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[10].Value = dJr;

                arr[11].ParameterName = "@Ja";
                if (dJa == null)
                    arr[11].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[11].Value = dJa;

                arr[12].ParameterName = "@DegBreak";
                if (sDegBreak== null)
                    arr[12].Value = "";
                else arr[12].Value = sDegBreak;

                arr[13].ParameterName = "@AltWeath";
                if (sAltWeath == null)
                        arr[13].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[13].Value = sAltWeath;

                arr[14].ParameterName = "@Hardness";
                if (sHardness == null)
                    arr[14].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[14].Value = sHardness;

                arr[15].ParameterName = "@Comments";
                if (sComments == null)
                    arr[15].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[15].Value = sComments;

                arr[16].ParameterName = "@DHGeotechID";
                arr[16].Value = iDHGeotechID; 

                oRes = oData.ExecuteScalar("usp_DH_Geotech_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Save error Geotech. " + eX.Message); ;
            }
        }

    }



