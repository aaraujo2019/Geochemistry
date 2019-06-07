using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


    public class clsGCSamplesRockLith
    {
        public string sOpcion;
        public string sSample;
        public string sMineral;
        public string sType;
        public int iSkLithRock;

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

        public string GCSamplesRockLith_Add()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(5);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@Sample";
                arr[1].Value = sSample;
                arr[2].ParameterName = "@Mineral";
                arr[2].Value = sMineral;
                arr[3].ParameterName = "@TypeMat_Phe";
                arr[3].Value = sType;
                arr[4].ParameterName = "@SKLithRock";
                arr[4].Value = iSkLithRock;

                oRes = oData.ExecuteScalar("usp_GC_SamplesRockMinLith_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Save error GCSamplesRockLith. " + eX.Message); ;
            }
        }

        public string GCSamplesRockLith_Delete()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@SKLithRock";
                arr[0].Value = iSkLithRock;

                oRes = oData.ExecuteScalar("usp_GC_SamplesRockMinLith_Delete", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Delete error GCSamplesRockLith. " + eX.Message); ;
            }
        }

        public DataTable getGCSamplesRockLithList()
        {
            try
            {
                DataSet dtGCSamplesRock = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@Sample";
                arr[1].Value = sSample; 
                dtGCSamplesRock = oData.ExecuteDataset("usp_GC_SamplesRockMinLith_List", arr, CommandType.StoredProcedure);
                return dtGCSamplesRock.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in getGCSamplesRockLith: " + eX.Message);
            }
        }

    }

