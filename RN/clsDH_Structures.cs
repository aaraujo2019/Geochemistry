using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

   public class clsDH_Structures
    {

        public string sOpcion;
        public string sHoleID;
        public double iFrom;
        public double iTo;
        public string sType;
        public double? dAngleToCore;
        public double? dUpAngle;
        public double? dBtonAngle;
        public double? dAppThick;
        public string sFill;

        public string sFill2;
        public string sFill3;
        public string sFill4;

        public double? dNumber; 

        public string sComments;
        public double? dLenght;
        public Int64 iDHStructrueID;

        public static string sStaticFrom;

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

        public DataTable getDH_Structures()
        {
            try
            {

                DataSet dtDH_Structures = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                dtDH_Structures = oData.ExecuteDataset("usp_DH_Structures_List", arr, CommandType.StoredProcedure);
                return dtDH_Structures.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in DH_Structures: " + eX.Message);
            }
        }

        public string DH_Structures_Add()
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
                arr[4].ParameterName = "@Type";
                arr[4].Value = sType;
                
                arr[5].ParameterName = "@AngleToCore";
                if (dAngleToCore == null)
                    arr[5].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[5].Value = dAngleToCore;

                arr[6].ParameterName = "@UpAngle";
                if (dUpAngle == null)
                    arr[6].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[6].Value = dUpAngle;

                arr[7].ParameterName = "@BtonAngle";
                if (dBtonAngle == null)
                    arr[7].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[7].Value = dBtonAngle;

                arr[8].ParameterName = "@AppThick";
                if (dAppThick == null)
                    arr[8].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[8].Value = dAppThick;

                arr[9].ParameterName = "@Fill";
                if (sFill == null)
                    arr[9].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[9].Value = sFill;

                arr[10].ParameterName = "@Number";
                if (dNumber == null)
                    arr[10].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[10].Value = dNumber;

                arr[11].ParameterName = "@Comments";
                if (sComments == null)
                    arr[11].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[11].Value = sComments;

                arr[12].ParameterName = "@Lenght";
                if (dLenght == null)
                    arr[12].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[12].Value = dLenght;

                arr[13].ParameterName = "@DHStructrueID";
                arr[13].Value = iDHStructrueID;

                arr[14].ParameterName = "@Fill2";
                if (sFill2 == null)
                    arr[14].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[14].Value = sFill2;

                arr[15].ParameterName = "@Fill3";
                if (sFill3 == null)
                    arr[15].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[15].Value = sFill3;

                arr[16].ParameterName = "@Fill4";
                if (sFill4 == null)
                    arr[16].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[16].Value = sFill4;

                oRes = oData.ExecuteScalar("usp_DH_Structures_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Save error Structure. " + eX.Message); ;
            }
        }

        public string DH_Structures_Delete()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@SKDHStructrue";
                arr[0].Value = iDHStructrueID;
                oRes = oData.ExecuteScalar("usp_DH_Structures_Delete", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Delete error Structure. " + eX.Message); ;
            }
        }


        public DataTable getDH_StructuresValidFromToNext()
        {
            try
            {
                DataSet dtDH_Structures = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDH_Structures = oData.ExecuteDataset("usp_DH_Structures_ListValidFromToNext", arr, CommandType.StoredProcedure);
                return dtDH_Structures.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error in DH_StructuresValidFromToNext: " + eX.Message);
            }
        }


        public DataTable getDH_StructuresValid()
        {
            try
            {
                DataSet dtDH_Structures = new DataSet();
                SqlParameter[] arr = oData.GetParameters(4);
                arr[0].ParameterName = "@From";
                arr[0].Value = iFrom;
                arr[1].ParameterName = "@To";
                arr[1].Value = iTo;
                arr[2].ParameterName = "@HoleID";
                arr[2].Value = sHoleID;
                arr[3].ParameterName = "@SKDHStructrue";
                arr[3].Value = iDHStructrueID;

                dtDH_Structures = oData.ExecuteDataset("usp_DH_Structures_ListValid", arr, CommandType.StoredProcedure);
                return dtDH_Structures.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error in DH_StructuresValid: " + eX.Message);
            }
        }

    }

