using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


    public class clsDHLithology
    {
        public string sOpcion;
        public string sHoleID;
        public double dFrom;
        public double dTo;
        public string sLithCode;
        public string sObservation;
        public Int64 iDHLithologyID;
        public string sGSize;
        public string sTextures;

        public static string sStaticFrom;

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

        public DataTable getDH_Lithology()
        {
            try
            {

                DataSet dtDH_Lithology = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                dtDH_Lithology = oData.ExecuteDataset("usp_DH_Lithology_List", arr, CommandType.StoredProcedure);
                return dtDH_Lithology.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DH_Lithology: " + eX.Message);
            }
        }

        
        public DataTable getDHLitFromToValidFromToNext()
        {
            try
            {
                DataSet dtDHLitFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHLitFromToValid = oData.ExecuteDataset("usp_DH_Lithology_ListValidFromToNext", arr, CommandType.StoredProcedure);
                return dtDHLitFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHLitFromToValidFromToNext: " + eX.Message);
            }
        }


        public DataTable getDHLitFromToValid()
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
                //arr[3].ParameterName = "@Opcion";
                //arr[3].Value = sOpcion;
                dtDHLitFromToValid = oData.ExecuteDataset("usp_DH_Lithology_ListFromToValid", arr, CommandType.StoredProcedure);
                return dtDHLitFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHLitFromToValid: " + eX.Message);
            }
        }

        public DataTable getDHLitValid()
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
                arr[3].ParameterName = "@SKDHLithology";
                arr[3].Value = iDHLithologyID;
                dtDHLitFromToValid = oData.ExecuteDataset("usp_DH_Lithology_ListValid", arr, CommandType.StoredProcedure);
                return dtDHLitFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHLitFromToValid: " + eX.Message);
            }
        }

        public string DH_Lithology_Add()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(9);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                arr[2].ParameterName = "@From";
                arr[2].Value = dFrom;
                arr[3].ParameterName = "@To";
                arr[3].Value = dTo;
                arr[4].ParameterName = "@CodeLithology";
                arr[4].Value = sLithCode;
                
                arr[5].ParameterName = "@Observation";
                if (sObservation == null)
                    arr[5].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[5].Value = sObservation;

                arr[6].ParameterName = "@DHLithologyID";
                arr[6].Value = iDHLithologyID;

                arr[7].ParameterName = "@GSize";
                if (sGSize == null)
                    arr[7].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[7].Value = sGSize;

                arr[8].ParameterName = "@Textures";
                if (sTextures == null)
                    arr[8].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[8].Value = sTextures;

                oRes = oData.ExecuteScalar("usp_DH_Lithology_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Save error Lithology Insert. " + eX.Message); ;
            }
        }

        public string DH_Lithology_Delete()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                arr[1].ParameterName = "@From";
                arr[1].Value = dFrom;

                oRes = oData.ExecuteScalar("usp_DH_Lithology_Delete", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Delete error Lithology. " + eX.Message); ;
            }
        }

    }

