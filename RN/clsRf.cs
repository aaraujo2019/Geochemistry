using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


public class clsRf
{
    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public static string sUser;
    public static string sIdentification;
    public static string sIdGrupo;
    public static DataSet dsPermisos;
    
    public string sOpcion;
    public int iIdProject;


    public string sCodeLith;



    public DataTable getSphericity()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_RfSphericity_List", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getSorting()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_RfSorting_List", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfPorpuose()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_RfPorpuose_ListCmb", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }
    public DataTable getRfRelativeToVeinLocation()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_RfRelativeToVeinLocation_ListCmb", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfNotInSituCmb()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_RfNotInSitu_ListCmb", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfCoordinateSystemCmb()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_RfCoordinateSystem_ListCmb", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfTargetCmb()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_RfTarget_ListCmb", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRounding()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_RfRounding_List", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getMineEntranceList(int _sOpcion, string _sMineId)
    {
        try
        {
            DataSet dsData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = _sOpcion;
            arr[1].ParameterName = "@MineId";
            arr[1].Value = _sMineId;
            dsData = oData.ExecuteDataset("usp_Mi_MineEntrance_List", arr, CommandType.StoredProcedure);
            return dsData.Tables[0];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    public DataTable getMineEntrance()
    {
        try
        {
            DataSet dtData = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtData = oData.ExecuteDataset("usp_Mi_MineEntrance_Cmb", arr, CommandType.StoredProcedure);
            return dtData.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getUsers(string _sUser)
    {
        try
        {
            DataSet dsUser = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@sUsuario";
            arr[0].Value = _sUser;
            dsUser = oData.ExecuteDataset("usp_getUsuarios_PORTAL", arr, CommandType.StoredProcedure);
            return dsUser.Tables[0];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    public DataTable getLocation(string _sCode)
    {
        try
        {
            DataSet dtLocation = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@CODE";
            arr[0].Value = _sCode;
            dtLocation = oData.ExecuteDataset("usp_RfLocation_List", arr, CommandType.StoredProcedure);
            return dtLocation.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getVersionProject()
    {
        try
        {
            DataSet dtVersion = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Id";
            arr[0].Value = iIdProject;
            dtVersion = oData.ExecuteDataset("usp_getProject", arr, CommandType.StoredProcedure);
            return dtVersion.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }


    public DataTable getRfTextures_ListAll()
    {
        try
        {
            DataSet dtRfTextures = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            dtRfTextures = oData.ExecuteDataset("usp_RfTextures_ListAll", arr, CommandType.StoredProcedure);
            return dtRfTextures.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRFGsize_ListAll()
    {
        try
        {
            DataSet dtRFGsize = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            dtRFGsize = oData.ExecuteDataset("usp_RFGsize_ListAll", arr, CommandType.StoredProcedure);
            return dtRFGsize.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }


    public DataTable getRfTextures_List()
    {
        try
        {
            DataSet dtRfTextures = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Code";
            arr[0].Value = sCodeLith;
            dtRfTextures = oData.ExecuteDataset("usp_RfTextures_List", arr, CommandType.StoredProcedure);
            return dtRfTextures.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRFGsize_List()
    {
        try
        {
            DataSet dtRFGsize = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Code";
            arr[0].Value = sCodeLith;
            dtRFGsize = oData.ExecuteDataset("usp_RFGsize_List", arr, CommandType.StoredProcedure);
            return dtRFGsize.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }
    
    public string InsertTrans(string _sTABLE, string _TRANS, string _LOGINTRANS, string _DATOSTRANS)
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@sTABLE";
            arr[0].Value = _sTABLE;
            arr[1].ParameterName = "@TRANS";
            arr[1].Value = _TRANS;
            arr[2].ParameterName = "@LOGINTRANS";
            arr[2].Value = _LOGINTRANS;
            arr[3].ParameterName = "@DATOSTRANS";
            arr[3].Value = _DATOSTRANS; 
            oRes = oData.ExecuteScalar("[usp_TransactionsAdd]", arr, CommandType.StoredProcedure);
            return oRes.ToString();

        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    //[[usp_saveUserPasswd]]
    public string UpdatePass(string _sPassOld, string _sPass, string _sLogin)
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(3);
            arr[0].ParameterName = "@PasswdOld";
            arr[0].Value = _sPassOld;
            arr[1].ParameterName = "@PasswdNew";
            arr[1].Value = _sPass;
            arr[2].ParameterName = "@LoginUser";
            arr[2].Value = _sLogin;
            oRes = oData.ExecuteScalar("[usp_saveUserPasswd]", arr, CommandType.StoredProcedure);
            return oRes.ToString();

        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    //usp_TransactionsList
    public DataTable getTransList(string _sUser)
    {
        try
        {
            DataSet dtTransList = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@LOGINTRANS";
            arr[0].Value = _sUser;
            dtTransList = oData.ExecuteDataset("usp_TransactionsList", arr, CommandType.StoredProcedure);
            return dtTransList.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    //usp_RfWorker_ListByCred
    public DataTable getRfWorkerCred( string _sCod, string _sPass)
    {
        try
        {
            DataSet dtRfWorker = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Cod";
            arr[0].Value = _sCod;
            arr[1].ParameterName = "@Password";
            arr[1].Value = _sPass;
            dtRfWorker = oData.ExecuteDataset("usp_RfWorker_ListByCred", arr, CommandType.StoredProcedure);
            return dtRfWorker.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    //usp_Inv_Samples_List
    public DataTable getInvSamples()
    {
        try
        {
            DataSet dtInvSamples = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtInvSamples = oData.ExecuteDataset("usp_Inv_Samples_List", arr, CommandType.StoredProcedure);
            return dtInvSamples.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfPrefixW_List()
    {
        try
        {
            DataSet dtRfPrefixW = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfPrefixW = oData.ExecuteDataset("usp_RfPrefixW_List", arr, CommandType.StoredProcedure);
            return dtRfPrefixW.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfTypeStructure_List()
    {
        try
        {
            DataSet dtRfTypeStructure = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfTypeStructure = oData.ExecuteDataset("usp_RfTypeStructure_List", arr, CommandType.StoredProcedure);
            return dtRfTypeStructure.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfGSize_ListMin(string _sOpcion)
    {
        try
        {
            DataSet dtRfGSize = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = _sOpcion;
            dtRfGSize = oData.ExecuteDataset("usp_RFGsize_ListAll", arr, CommandType.StoredProcedure);
            return dtRfGSize.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    //[usp_RfFillStructure_List]
    public DataTable getRfFillStructure_List()
    {
        try
        {
            DataSet dtRfStructure = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfStructure = oData.ExecuteDataset("usp_RfFillStructure_List", arr, CommandType.StoredProcedure);
            return dtRfStructure.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

    }

    //[usp_RfStyleAlt_List]
    public DataTable getRfStyleAlt_List()
    {
        try
        {
            DataSet dtRfStyleAlt = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfStyleAlt = oData.ExecuteDataset("usp_RfStyleAlt_List", arr, CommandType.StoredProcedure);
            return dtRfStyleAlt.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

    }

    //usp_RfMinerAlt_List
    public DataTable getRfMinerAlt_List()
    {
        try
        {
            DataSet dtMinerAlt = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtMinerAlt = oData.ExecuteDataset("usp_RfMinerAlt_List", arr, CommandType.StoredProcedure);
            return dtMinerAlt.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

    }

    //[usp_RfIntensityAlt_List]
    public DataTable getRfIntensityAlt_List(string _sProjectGC)
    {
        try
        {
            DataSet dtRfIntensityAlt = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Project";
            arr[0].Value = _sProjectGC;
            dtRfIntensityAlt = oData.ExecuteDataset("usp_RfIntensityAlt_List", arr, CommandType.StoredProcedure);
            return dtRfIntensityAlt.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    //[usp_RfTypeAlt_List]
    public DataTable getRfTypeAlt_List()
    {
        try
        {
            DataSet dtRfTypeAlt = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfTypeAlt = oData.ExecuteDataset("usp_RfTypeAlt_List", arr, CommandType.StoredProcedure);
            return dtRfTypeAlt.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    //[usp_RfMinPercent_List]
    public DataTable getRfMinerPercent_List(string _sProjectGC)
    {
        try
        {
            DataSet dtRfMinerMinSt = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Project";
            arr[0].Value = _sProjectGC;
            dtRfMinerMinSt = oData.ExecuteDataset("usp_RfMinPercent_List", arr, CommandType.StoredProcedure);
            return dtRfMinerMinSt.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfMinerMinSt_List()
    {
        try
        {
            DataSet dtRfMinerMinSt = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfMinerMinSt = oData.ExecuteDataset("usp_RfMinerStyle_List", arr, CommandType.StoredProcedure);
            return dtRfMinerMinSt.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfMinerMin_ListOxid()
    {
        try
        {
            DataSet dtRfMinerMin = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfMinerMin = oData.ExecuteDataset("usp_RfMinerMin_ListOxid", arr, CommandType.StoredProcedure);
            return dtRfMinerMin.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }


    public DataTable getRfVetas_List(string _sCode)
    {
        try
        {
            DataSet dtRfVetas = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Code";
            arr[0].Value = _sCode;
            dtRfVetas = oData.ExecuteDataset("usp_RfVetas_List", arr, CommandType.StoredProcedure);
            return dtRfVetas.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfContactType_List()
    {
        try
        {
            DataSet dtRfContactType = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfContactType = oData.ExecuteDataset("usp_RfContactType_List", arr, CommandType.StoredProcedure);
            return dtRfContactType.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfMinerMinAlt_ListAll()
    {
        try
        {
            DataSet dtRfMinerMin = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfMinerMin = oData.ExecuteDataset("usp_Miner_MinAlt", arr, CommandType.StoredProcedure);
            return dtRfMinerMin.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfMinerMin_List()
    {
        try
        {
            DataSet dtRfMinerMin = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfMinerMin = oData.ExecuteDataset("usp_RfMinerMin_List", arr, CommandType.StoredProcedure);
            return dtRfMinerMin.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfMinerMinAlt_List()
    {
        try
        {
            DataSet dtRfMinerMin = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfMinerMin = oData.ExecuteDataset("usp_RfMinerAlt_List", arr, CommandType.StoredProcedure);
            return dtRfMinerMin.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }
   
    
    public DataTable getRfColour_List()
    {
        try
        {
            DataSet dtRfColour = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfColour = oData.ExecuteDataset("usp_RfColour_List", arr, CommandType.StoredProcedure);
            return dtRfColour.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    /// <summary>
    /// Oxidation Intensity
    /// </summary>
    /// <returns></returns>
    public DataTable getRfOxidation_List()
    {
        try
        {
            DataSet dtRfOxidation_List = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfOxidation_List = oData.ExecuteDataset("usp_RfOxidation_List", arr, CommandType.StoredProcedure);
            return dtRfOxidation_List.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    /// <summary>
    /// Oxidation Percent
    /// </summary>
    /// <returns></returns>
    public DataTable getRfOxides_List()
    {
        try
        {
            DataSet dtRfOxidation_List = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfOxidation_List = oData.ExecuteDataset("usp_RfOxides_List", arr, CommandType.StoredProcedure);
            return dtRfOxidation_List.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getWeathering()
    {
        try
        {
            DataSet dtRfGeotechAltMet = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfGeotechAltMet = oData.ExecuteDataset("usp_RfWeathering_List", arr, CommandType.StoredProcedure);
            return dtRfGeotechAltMet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfGeotechHardness()
    {
        try
        {
            DataSet dtRfGeotechHardness = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfGeotechHardness = oData.ExecuteDataset("usp_RfGeotechHardness_List", arr, CommandType.StoredProcedure);
            return dtRfGeotechHardness.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfGeotechBreak()
    {
        try
        {
            DataSet dtRfGeotechBreak = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfGeotechBreak = oData.ExecuteDataset("usp_RfGeotechBreak_List", arr, CommandType.StoredProcedure);
            return dtRfGeotechBreak.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfPrefix()
    {
        try
        {
            DataSet dtRfPrefix = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfPrefix = oData.ExecuteDataset("usp_Prefix_List", arr, CommandType.StoredProcedure);
            return dtRfPrefix.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfWorker()
    {
        try
        {
            DataSet dtRfWorker = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfWorker = oData.ExecuteDataset("usp_RfWorker_List", arr, CommandType.StoredProcedure);
            return dtRfWorker.Tables[0];
        }
        catch (Exception ex)
        {
            
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfCodeLab()
    {
        try
        {
            DataSet dtRfCodeLab = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfCodeLab = oData.ExecuteDataset("usp_RfCodeLab_List", arr, CommandType.StoredProcedure);
            return dtRfCodeLab.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfPreparationCode()
    {
        try
        {
            DataSet dtRfPreparationCode = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfPreparationCode = oData.ExecuteDataset("usp_RfPreparationCode_List", arr, CommandType.StoredProcedure);
            return dtRfPreparationCode.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfAnalysisCodeLab()
    {
        try
        {
            DataSet dtRfAnalysisCodeLab = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtRfAnalysisCodeLab = oData.ExecuteDataset("usp_RfAnalysisCodeLab_List", arr, CommandType.StoredProcedure);
            return dtRfAnalysisCodeLab.Tables[0];
        }
        catch (Exception ex)
        {

            throw new Exception(ex.Message);
        }
    }

    public DataTable getUsersPortal(string _sLogin)
    {
        try
        {
            DataSet dtUsersPortal = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@sLogin";
            arr[0].Value = _sLogin;
            dtUsersPortal = oData.ExecuteDataset("usp_getUsersSubpartners_PORTAL", arr, CommandType.StoredProcedure);
            return dtUsersPortal.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }
    
    //[usp_DH_RfTypeSample_List]
    public DataTable getRfTypeSample()
    {
        try
        {
            DataSet dsRfTypeSample = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dsRfTypeSample = oData.ExecuteDataset("usp_DH_RfTypeSample_List", arr, CommandType.StoredProcedure);
            return dsRfTypeSample.Tables[0];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    public DataTable getRfOxidationInt_List()
    {
        try
        {
            DataSet dsRfOxidationInt = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dsRfOxidationInt = oData.ExecuteDataset("usp_RfOxidationIntensity_List", arr, CommandType.StoredProcedure);
            return dsRfOxidationInt.Tables[0];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    public DataSet getRfTypeSampleDataSet()
    {
        try
        {
            DataSet dsRfTypeSample = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dsRfTypeSample = oData.ExecuteDataset("usp_DH_RfTypeSample_List", arr, CommandType.StoredProcedure);
            return dsRfTypeSample;
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    public DataSet getDsRfLithology()
    {
        try
        {
            DataSet dsRfLithology = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dsRfLithology = oData.ExecuteDataset("usp_DH_RfLithology_List", arr, CommandType.StoredProcedure);
            return dsRfLithology;
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    ////Se utilizara para Samples
    public DataTable getRfLithology()
    {
        try
        {
            DataSet dsRfLithology = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dsRfLithology = oData.ExecuteDataset("usp_DH_RfLithology_List", arr, CommandType.StoredProcedure);
            return dsRfLithology.Tables[0];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    //Se utilizara para Lithology
    public DataTable getRfLithologyDH()
    {
        try
        {
            DataSet dsRfLithology = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dsRfLithology = oData.ExecuteDataset("usp_DH_RfLithology_List", arr, CommandType.StoredProcedure);
            return dsRfLithology.Tables[1];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    //Permisos por formulario
    public DataSet getFormsByGrupoAll(string _IdGrupo)
    {
        try
        {
            DataSet dsFormsGroup = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@idGrupo";
            arr[0].Value = _IdGrupo;
            dsFormsGroup = oData.ExecuteDataset("usp_getFormByGrupoAll_PORTAL", arr, CommandType.StoredProcedure);
            return dsFormsGroup;
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    //Permisos por formulario
    public DataTable getFormsByGrupo(string _sIdGrupo, string _sIDGrupo)
    {
        try
        {
            DataSet dtFormsGroup = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@idGrupo";
            arr[0].Value = _sIdGrupo;
            arr[1].ParameterName = "@ID_Project";
            arr[1].Value = _sIDGrupo;
            dtFormsGroup = oData.ExecuteDataset("usp_getFormulariosByGrupo", arr, CommandType.StoredProcedure);
            return dtFormsGroup.Tables[0];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    //Permisos en cada formulario por cada accion (insertar, modificar, eliminar ...)
    public DataTable getPermisosFormsByGrupo(string _IdGrupo)
    {
        try
        {
            DataSet dtPermFormsGroup = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@idGrupo";
            arr[0].Value = _IdGrupo;
            dtPermFormsGroup = oData.ExecuteDataset("usp_getPermisosFormByGrupo_PORTAL", arr, CommandType.StoredProcedure);
            return dtPermFormsGroup.Tables[0];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

    //[usp_getUsuarios_PORTAL]
    public DataTable GetUsuarios(string _IdUser)
    {
        try
        {
            DataSet dtUsers = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@sUsuario";
            arr[0].Value = _IdUser;
            dtUsers = oData.ExecuteDataset("usp_getUsuarios_PORTAL", arr, CommandType.StoredProcedure);
            return dtUsers.Tables[0];
        }
        catch (Exception eX)
        {

            throw new Exception(eX.Message);
        }
    }

}

