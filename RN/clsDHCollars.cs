using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


    public class clsDHCollars
    {
        public string sHoleID;
        public string sLogged;

        public string sLoggedBy1;
        public string sLoggedBy2;
        public string sLoggedBy3;
        

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

        //[usp_DH_Collars_UpdateAssign]
        public string DHSamples_UpdateAssign()
        {
            try
            {
                
                object oRes;
                SqlParameter[] arr = oData.GetParameters(4);
                arr[0].ParameterName = "@LoggedBy1";
                arr[0].Value = sLoggedBy1;
                arr[1].ParameterName = "@LoggedBy2";
                arr[1].Value = sLoggedBy2;
                arr[2].ParameterName = "@LoggedBy3";
                arr[2].Value = sLoggedBy3;
                arr[3].ParameterName = "@HoleID";
                arr[3].Value = sHoleID;
                /*El procedimiento o la función 'usp_DH_Collars_UpdateAssign' 
                 * esperaba el parámetro '@LoggedBy2', que no se ha especificado.*/

                oRes = oData.ExecuteScalar("usp_DH_Collars_UpdateAssign", arr, CommandType.StoredProcedure);
                //ds = oDAtos.ExecuteDataset("usp_Datos_ListByID", arr, CommandType.StoredProcedure)
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Save error UpdateAssign. " + eX.Message); ;
            }
        }

        //[usp_DH_Collars_List]
        public DataTable getDHCollars()
        {
            try
            {

                DataSet dtDHCollars = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHCollars = oData.ExecuteDataset("usp_DH_Collars_List", arr, CommandType.StoredProcedure);
                return dtDHCollars.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHcollars: " + eX.Message);
            }
        }

        //[usp_DH_Collars_ListAssign]
        public DataTable getDHCollarsListAssign()
        {
            try
            {

                DataSet dtDHCollars = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHCollars = oData.ExecuteDataset("usp_DH_Collars_ListAssign", arr, CommandType.StoredProcedure);
                return dtDHCollars.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHcollarsListAsign: " + eX.Message);
            }
        }

        //[usp_DH_Collars_ListLogged]
        public DataTable getDHCollarsLogged()
        {
            try
            {

                DataSet dtDHCollars = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                arr[1].ParameterName = "@Logged";
                arr[1].Value = sLogged; 
                dtDHCollars = oData.ExecuteDataset("usp_DH_Collars_ListLogged", arr, CommandType.StoredProcedure);
                return dtDHCollars.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DHcollarsLogged: " + eX.Message);
            }
        }

    }

