using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;



    public class clsDHMineraliz
    {
        public string sOpcion;
        public string sHoleID;
        public double dFrom;
        public double dTo;
        public string sMZ1Mineral;
        public string sMZ1Mineral2;
        public string sMZ1Mineral3;
        public string dMZ1Perc;
        public string sMZ1Style;
        public string sMZ2Mineral;
        public string sMZ2Mineral2;
        public string sMZ2Mineral3;
        public string dMZ2Perc;
        public string sMZ2Style;
        public string sMZ3Mineral;
        public string sMZ3Mineral2;
        public string sMZ3Mineral3;
        public string dMZ3Perc;
        public string sMZ3Style;
        public string sComments;
        public Int64 iDHMinID;

        public string sGSize1;
        public string sGSize2;
        public string sGSize3;

        public static string sStaticFrom;

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

        //[usp_DH_Mineraliz_ListValid]
        public DataTable getDHMinValid()
        {
            try
            {
                DataSet dtDHMinFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(4);
                arr[0].ParameterName = "@From";
                arr[0].Value = dFrom;
                arr[1].ParameterName = "@To";
                arr[1].Value = dTo;
                arr[2].ParameterName = "@HoleID";
                arr[2].Value = sHoleID;
                arr[3].ParameterName = "@SKDHMin";
                arr[3].Value = iDHMinID;
                dtDHMinFromToValid = oData.ExecuteDataset("usp_DH_Mineraliz_ListValid", arr, CommandType.StoredProcedure);
                return dtDHMinFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in dtDHMinFromToValid: " + eX.Message);
            }
        }

        
        public DataTable getDHMinFromToValidFromToNext()
        {
            try
            {
                DataSet dtDHMinFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHMinFromToValid = oData.ExecuteDataset("usp_DH_Mineralizations_ListValidFromToNext", arr, CommandType.StoredProcedure);
                return dtDHMinFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in dtDHMinFromToValidFromToNext: " + eX.Message);
            }
        }


        public DataTable getDHMinFromToValid()
        {
            try
            {
                DataSet dtDHMinFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(3);
                arr[0].ParameterName = "@From";
                arr[0].Value = dFrom;
                arr[1].ParameterName = "@To";
                arr[1].Value = dTo;
                arr[2].ParameterName = "@HoleID";
                arr[2].Value = sHoleID;
                //arr[3].ParameterName = "@Opcion";
                //arr[3].Value = sOpcion;
                dtDHMinFromToValid = oData.ExecuteDataset("usp_DH_Mineraliz_ListFromToValid", arr, CommandType.StoredProcedure);
                return dtDHMinFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in dtDHMinFromToValid: " + eX.Message);
            }
        }

        public DataTable getDHMineraliz()
        {
            try
            {
                DataSet dtDHMin = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                dtDHMin = oData.ExecuteDataset("usp_DH_Mineraliz_List", arr, CommandType.StoredProcedure);
                return dtDHMin.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in dtDHMineraliz: " + eX.Message);
            }
        }

        public string DH_Mineraliz_Delete()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@DHMinID";
                arr[0].Value = iDHMinID;

                oRes = oData.ExecuteScalar("usp_DH_Mineraliz_Delete", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Delete error DH_Mineraliz. " + eX.Message); ;
            }
        }

        public string DH_Mineraliz_Add()
        {
            try
            {

                object oRes;
                SqlParameter[] arr = oData.GetParameters(24);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                arr[2].ParameterName = "@From";
                arr[2].Value = dFrom;
                arr[3].ParameterName = "@To";
                arr[3].Value = dTo;
                arr[4].ParameterName = "@MZ1Mineral";
                arr[4].Value = sMZ1Mineral;

                arr[5].ParameterName = "@MZ1Mineral2";
                if (sMZ1Mineral2 == null)
                    arr[5].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[5].Value = sMZ1Mineral2;

                arr[6].ParameterName = "@MZ1Mineral3";
                if (sMZ1Mineral3 == null)
                    arr[6].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[6].Value = sMZ1Mineral3;

                arr[7].ParameterName = "@MZ1Perc";
                if (dMZ1Perc == null)
                    arr[7].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[7].Value = dMZ1Perc;

                arr[8].ParameterName = "@MZ1Style";
                if (sMZ1Style == null)
                    arr[8].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[8].Value = sMZ1Style;

                arr[9].ParameterName = "@MZ2Mineral";
                if (sMZ2Mineral == null)
                    arr[9].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[9].Value = sMZ2Mineral;

                arr[10].ParameterName = "@MZ2Mineral2";
                if (sMZ2Mineral2 == null)
                    arr[10].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[10].Value = sMZ2Mineral2;

                arr[11].ParameterName = "@MZ2Mineral3";
                if (sMZ2Mineral3 == null)
                    arr[11].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[11].Value = sMZ2Mineral3;

                arr[12].ParameterName = "@MZ2Perc";
                if (dMZ2Perc == null)
                    arr[12].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[12].Value = dMZ2Perc;

                arr[13].ParameterName = "@MZ2Style";
                if (sMZ2Style == null)
                    arr[13].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[13].Value = sMZ2Style;

                arr[14].ParameterName = "@MZ3Mineral";
                if (sMZ3Mineral == null)
                    arr[14].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[14].Value = sMZ3Mineral;

                arr[15].ParameterName = "@MZ3Mineral2";
                if (sMZ3Mineral2 == null)
                    arr[15].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[15].Value = sMZ3Mineral2;

                arr[16].ParameterName = "@MZ3Mineral3";
                if (sMZ3Mineral3 == null)
                    arr[16].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[16].Value = sMZ3Mineral3;

                arr[17].ParameterName = "@MZ3Perc";
                if (dMZ3Perc == null)
                    arr[17].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[17].Value = dMZ3Perc;

                arr[18].ParameterName = "@MZ3Style";
                if (sMZ3Style == null)
                    arr[18].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[18].Value = sMZ3Style;

                arr[19].ParameterName = "@Comments";
                if (sComments == null)
                    arr[19].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[19].Value = sComments;

                arr[20].ParameterName = "@DHMinID";
                arr[20].Value = iDHMinID;


                arr[21].ParameterName = "@Gsize";
                if (sGSize1 == null)
                    arr[21].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[21].Value = sGSize1;

                arr[22].ParameterName = "@GSize2";
                if (sGSize2 == null)
                    arr[22].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[22].Value = sGSize2;

                arr[23].ParameterName = "@GSize3";
                if (sGSize3 == null)
                    arr[23].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[23].Value = sGSize3;
                /*@Gsize varchar(3),
                @GSize2 varchar(3),
                @GSize3 varchar(3)
                 
                public string sGSize1;
                public string sGSize2;
                public string sGSize3;
                 */

                oRes = oData.ExecuteScalar("usp_DH_Mineraliz_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Save error DH_Mineraliz. " + eX.Message); ;
            }
        }

    }

