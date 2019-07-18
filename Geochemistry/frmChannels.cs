using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

namespace Geochemistry
{
    public partial class frmChannels : Form
    {

        Configuration conf = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

        clsCHChannels oCh = new clsCHChannels();
        clsCHSurveys oSur = new clsCHSurveys();
        clsCHSamples oCHSamp = new clsCHSamples();
        clsCHMinLith oMinLith = new clsCHMinLith();
        clsCHAlterations oAlt = new clsCHAlterations();
        clsCHMineralizations oMin = new clsCHMineralizations();
        clsCHOxides oOxid = new clsCHOxides();
        clsCHStructures oStr = new clsCHStructures();
        clsRf oRf = new clsRf();

        static string sEditCh = "0";
        static string sEditSur = "0";
        static string sEdit = "0";
        static string sEditMinLithMat = "0";
        static string sEditMinLithPhe = "0";
        static string sEditAlt = "0";
        static string sEditMin = "0";
        static string sEditOxid = "0";
        static string sEditStr = "0";
        static string sSampleSelect = "";
        static string sExport = ""; 

        public frmChannels()
        {
            InitializeComponent();
            LoadDgChannels();
            txtProjectCh.Text = ConfigurationSettings.AppSettings["IDProjectGC"].ToString();
            txtProject.Text = ConfigurationSettings.AppSettings["IDProjectGC"].ToString();
            Loadcmb();
            LoadCmbSurvey();
            LoadDgSurveys();
        }

        private void LoadChannelId()
        {
            try
            {
               
                DataTable dtChId = new DataTable();
                oCh.sOpcion = "3";
                oCh.sChId = "0";
                dtChId = oCh.getCH_Collars();
                DataRow dr = dtChId.NewRow();
                dr[0] = "Select an option..";
                dtChId.Rows.Add(dr);
                cmbChannelId.ValueMember = "Chid";
                cmbChannelId.DisplayMember = "Chid";
                cmbChannelId.DataSource = dtChId;
                //cmbChannelId.SelectedValue = "Select an option..";
                //cmbChannelId.AutoCompleteCustomSource = AutoCompleteCmb(dtChId, "Chid");
               
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void Loadcmb()
        {
            try
            {
                txtProject.Text = ConfigurationSettings.AppSettings["IDProjectGC"].ToString(); //Id Proyecto Gran Colombia. Ej GSG, GZG ...

                #region target
                DataTable dtTarget = oRf.getRfTargetCmb();
                DataRow drTarget = dtTarget.NewRow();
                drTarget[0] = "-1";
                drTarget[1] = "Select an option..";
                dtTarget.Rows.Add(drTarget);
                cmbTarget.DisplayMember = "Comb";
                cmbTarget.ValueMember = "Code";
                cmbTarget.DataSource = dtTarget;
                cmbTarget.SelectedValue = -1;
                #endregion

                DataTable dtMineEnt = new DataTable();
                dtMineEnt = oRf.getMineEntrance();
                DataRow drMineEnt = dtMineEnt.NewRow();
                drMineEnt[0] = "-1";
                drMineEnt[1] = "Select an option...";
                dtMineEnt.Rows.Add(drMineEnt);
                cmbMineEntrance.DisplayMember = "cmb";
                cmbMineEntrance.ValueMember = "MineID";
                cmbMineEntrance.DataSource = dtMineEnt;
                cmbMineEntrance.SelectedValue = "-1";


                #region CS (coordinate system)
                DataTable dtCS = oRf.getRfCoordinateSystemCmb();
                DataRow drCS = dtCS.NewRow();
                drCS[0] = "-1";
                drCS[1] = "Select an option..";
                dtCS.Rows.Add(drCS);
                cmbCS.DisplayMember = "Comb";
                cmbCS.ValueMember = "Code";
                cmbCS.DataSource = dtCS;
                cmbCS.SelectedValue = -1;
                #endregion

                #region NotInSitu
                DataTable dtNIS = oRf.getRfNotInSituCmb();
                DataRow drNIS = dtNIS.NewRow();
                drNIS[0] = "-1";
                drNIS[1] = "Select an option..";
                dtNIS.Rows.Add(drNIS);
                cmbNotInSitu.DisplayMember = "Comb";
                cmbNotInSitu.ValueMember = "Code";
                cmbNotInSitu.DataSource = dtNIS;
                cmbNotInSitu.SelectedValue = -1;
                #endregion

                #region Porpuose
                DataTable dtPorpuose = oRf.getRfPorpuose();
                DataRow drPorpuose = dtPorpuose.NewRow();
                drPorpuose[0] = "-1";
                drPorpuose[1] = "Select an option..";
                dtPorpuose.Rows.Add(drPorpuose);
                cmbPorpuose.DisplayMember = "Comb";
                cmbPorpuose.ValueMember = "Code";
                cmbPorpuose.DataSource = dtPorpuose;
                cmbPorpuose.SelectedValue = -1;
                #endregion

                #region Relative Location
                DataTable dtRelLocation = oRf.getRfRelativeToVeinLocation();
                DataRow drRelLoc = dtRelLocation.NewRow();
                drRelLoc[0] = "-1";
                drRelLoc[1] = "Select an option..";
                dtRelLocation.Rows.Add(drRelLoc);
                cmbRelativeLoc.DisplayMember = "Comb";
                cmbRelativeLoc.ValueMember = "Code";
                cmbRelativeLoc.DataSource = dtRelLocation;
                cmbRelativeLoc.SelectedValue = -1;
                #endregion

                #region Geologist

                DataTable dtUsers = new DataTable();
                dtUsers = oRf.getUsers("-99");

                DataRow dr2 = dtUsers.NewRow();
                dr2[0] = "-1";
                dr2[7] = "Select an option..";
                dtUsers.Rows.Add(dr2);

                cmbGeologist.DisplayMember = "cmb";
                cmbGeologist.ValueMember = "id";
                cmbGeologist.DataSource = dtUsers;
                cmbGeologist.SelectedValue = -1;

                #endregion

                //DataTable dtLocationC = new DataTable();
                //dtLocationC.Columns.Add("Key", typeof(String));
                //dtLocationC.Columns.Add("Value", typeof(String));


                //DataRow drConect;
                //for (int i = 0; i < conf.AppSettings.Settings.Count; i++)
                //{
                //    if (conf.AppSettings.Settings.AllKeys[i].ToString().Contains("Loc"))
                //    {

                //        drConect = dtLocationC.NewRow();
                //        //drConect["Con"] = ;
                //        drConect["Key"] = conf.AppSettings.Settings.AllKeys[i].ToString();
                //        drConect["Value"] =
                //            conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value.ToString();
                //        dtLocationC.Rows.Add(drConect);

                //        //MessageBox.Show(conf.AppSettings.Settings.AllKeys[i].ToString());
                //        cmbLocationChannel.Items.Add(conf.AppSettings.Settings.AllKeys[i].ToString());
                //        string s = conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value;
                //    }

                //}

                //drConect = dtLocationC.NewRow();
                //drConect["Key"] = "-1";
                //drConect["Value"] = "Select an option...";
                //dtLocationC.Rows.Add(drConect);

                //cmbLocationChannel.DisplayMember = "Value";
                //cmbLocationChannel.ValueMember = "Key";
                //cmbLocationChannel.DataSource = dtLocationC;
                //cmbLocationChannel.Text = "Select an option...";

                #region ChannelId

                LoadChannelId();
                cmbChannelId.SelectedValue = "Select an option..";

                #endregion

                #region SampleType

                DataSet dtSampleT = new DataSet();
                dtSampleT = oRf.getRfTypeSampleDataSet();

                DataRow drSTy = dtSampleT.Tables[1].NewRow();
                drSTy[0] = "-1";
                drSTy[1] = "Select an option..";
                dtSampleT.Tables[1].Rows.Add(drSTy);
                cmbSampleType.DisplayMember = "Comb";
                cmbSampleType.ValueMember = "Code";
                cmbSampleType.DataSource = dtSampleT.Tables[1];
                cmbSampleType.SelectedValue = -1;



                DataRow drT2 = dtSampleT.Tables[2].NewRow();
                drT2[0] = "-1";
                drT2[1] = "Select an option..";
                dtSampleT.Tables[2].Rows.Add(drT2);
                cmbSamplingType.DisplayMember = "Comb";
                cmbSamplingType.ValueMember = "Code";
                cmbSamplingType.DataSource = dtSampleT.Tables[2];
                cmbSamplingType.SelectedValue = -1;


                DataRow drChn = dtSampleT.Tables[3].NewRow();
                drChn[0] = "-1";
                drChn[1] = "Select an option..";
                dtSampleT.Tables[3].Rows.Add(drChn);
                cmbChannelType.DisplayMember = "Comb";
                cmbChannelType.ValueMember = "Code";
                cmbChannelType.DataSource = dtSampleT.Tables[3];
                cmbChannelType.SelectedValue = -1;

                #endregion

                #region GSize
                oRf.sOpcion = "2";
                DataTable dtGSize = new DataTable();
                dtGSize = oRf.getRFGsize_ListAll();
                DataRow drG = dtGSize.NewRow();
                drG[0] = "-1";
                drG[1] = "Select an option..";
                dtGSize.Rows.Add(drG);
                cmbLGsize.DisplayMember = "Comb";
                cmbLGsize.ValueMember = "Code";
                cmbLGsize.DataSource = dtGSize;
                cmbLGsize.SelectedValue = "-1";

                cmbMatrixGSize.DisplayMember = "Comb";
                cmbMatrixGSize.ValueMember = "Code";
                cmbMatrixGSize.DataSource = dtGSize.Copy();
                cmbMatrixGSize.SelectedValue = "-1";

                cmbPhenoGSize.DisplayMember = "Comb";
                cmbPhenoGSize.ValueMember = "Code";
                cmbPhenoGSize.DataSource = dtGSize.Copy();
                cmbPhenoGSize.SelectedValue = "-1";

                #endregion

                #region Textures
                oRf.sOpcion = "1";
                DataTable dtTextures = new DataTable();
                dtTextures = oRf.getRfTextures_ListAll();
                DataRow drTx = dtTextures.NewRow();
                drTx[0] = "-1";
                drTx[1] = "Select an option..";
                dtTextures.Rows.Add(drTx);
                cmbLTextures.DisplayMember = "Comb";
                cmbLTextures.ValueMember = "Code";
                cmbLTextures.DataSource = dtTextures;
                cmbLTextures.SelectedValue = "-1";

                #endregion

                #region ContactType

                DataTable dtContType = new DataTable();
                dtContType = oRf.getRfContactType_List();
                DataRow drCont = dtContType.NewRow();
                drCont[0] = "-1";
                drCont[1] = "Select an option..";
                dtContType.Rows.Add(drCont);
                cmbContactType.DisplayMember = "Comb";
                cmbContactType.ValueMember = "Code";
                cmbContactType.DataSource = dtContType;
                cmbContactType.SelectedValue = "-1";

                #endregion

                #region Vein
                DataTable dtVein = new DataTable();
                dtVein = oRf.getRfVetas_List("");
                DataRow drVein = dtVein.NewRow();
                drVein[0] = "-1";
                drVein[2] = "Select an option..";
                dtVein.Rows.Add(drVein);
                cmbVeinName.DisplayMember = "Comb";
                cmbVeinName.ValueMember = "Code";
                cmbVeinName.DataSource = dtVein;
                cmbVeinName.SelectedValue = "-1";
                #endregion


                #region Weathering
                DataTable dtWeathering = new DataTable();
                dtWeathering = oRf.getWeathering();
                DataRow drW = dtWeathering.NewRow();
                drW[0] = "-1";
                drW[1] = "Select an option..";
                dtWeathering.Rows.Add(drW);
                cmbLWeathering.DisplayMember = "Comb";
                cmbLWeathering.ValueMember = "Grade";
                cmbLWeathering.DataSource = dtWeathering;
                cmbLWeathering.SelectedValue = -1;

                #endregion

                #region Percent
                DataTable dtMinPerc = new DataTable();
                dtMinPerc = oRf.getRfMinerPercent_List(ConfigurationSettings.AppSettings["IDProjectGC"].ToString()); //Id Proyecto Gran Colombia. Ej GSG, GZG ...
                DataRow drMinPerc = dtMinPerc.NewRow();
                drMinPerc[0] = "-1";
                drMinPerc[1] = "Select an option..";
                dtMinPerc.Rows.Add(drMinPerc);

                //CargarCombosPerc(dtMinPerc, cmbMatrixPorc);
                //CargarCombosPerc(dtMinPerc, cmbPhenoPerc);
                //CargarCombosPerc(dtMinPerc, cmbPorcM);


                DataTable dtMinPercOX = new DataTable();
                dtMinPercOX = oRf.getRfOxides_List();
                DataRow drMinPercOx = dtMinPercOX.NewRow();
                drMinPercOx[0] = "-1";
                drMinPercOx[1] = "Select an option..";
                dtMinPercOX.Rows.Add(drMinPercOx);

                CargarCombosPerc(dtMinPercOX, cmbPercGoe);
                CargarCombosPerc(dtMinPercOX, cmbPercHem);
                CargarCombosPerc(dtMinPercOX, cmbPercJar);
                CargarCombosPerc(dtMinPercOX, cmbPercLim);

                #endregion

                #region Sedimentary
                DataTable dtSorting = oRf.getSorting();
                DataRow drSorting = dtSorting.NewRow();
                drSorting[0] = "-1";
                drSorting[1] = "Select an option..";
                dtSorting.Rows.Add(drSorting);
                cmbRSorting.DisplayMember = "Comb";
                cmbRSorting.ValueMember = "Code";
                cmbRSorting.DataSource = dtSorting;
                cmbRSorting.SelectedValue = "-1";

                DataTable dtSphericity = oRf.getSphericity();
                DataRow drSphericity = dtSphericity.NewRow();
                drSphericity[0] = "-1";
                drSphericity[1] = "Select an option..";
                dtSphericity.Rows.Add(drSphericity);
                cmbRSphericity.DisplayMember = "Comb";
                cmbRSphericity.ValueMember = "Code";
                cmbRSphericity.DataSource = dtSphericity;
                cmbRSphericity.SelectedValue = "-1";

                DataTable dtRounding = oRf.getRounding();
                DataRow drRounding = dtRounding.NewRow();
                drRounding[0] = "-1";
                drRounding[1] = "Select an option..";
                dtRounding.Rows.Add(drRounding);
                cmbRounding.DisplayMember = "Comb";
                cmbRounding.ValueMember = "Code";
                cmbRounding.DataSource = dtRounding;
                cmbRounding.SelectedValue = "-1";
                #endregion

                dgData.DataSource = LoadDataCH(txtSampleHead.Text.ToString());
                dgData.Columns["SKCHSamples"].Visible = false;

                dgLithology.DataSource = LoadDataCH(txtSampleHead.Text.ToString());
                dgLithology.Columns["SKCHSamples"].Visible = false;


                DataTable dtLithology = new DataTable();
                dtLithology = oRf.getDsRfLithology().Tables[1];

                DataRow drL = dtLithology.NewRow();
                drL[0] = "-1";
                drL[1] = "Select an option..";
                dtLithology.Rows.Add(drL);

                cmbLithologyLit.DisplayMember = "Comb";
                cmbLithologyLit.ValueMember = "Code";
                cmbLithologyLit.DataSource = dtLithology;
                cmbLithologyLit.SelectedValue = -1;

                cmbHostRock.DisplayMember = "Comb";
                cmbHostRock.ValueMember = "Code";
                cmbHostRock.DataSource = dtLithology.Copy();
                cmbHostRock.SelectedValue = -1;


                #region Minerals
                DataTable dtMineral = new DataTable();
                dtMineral = oRf.getRfMinerMin_List();
                DataRow drM = dtMineral.NewRow();
                drM[0] = "-1";
                drM[1] = "Select an option..";
                dtMineral.Rows.Add(drM);
                LoadCombos(dtMineral, cmbMineralmin);


                DataTable dtMineralMinAlt = new DataTable();
                dtMineralMinAlt = oRf.getRfMinerMinAlt_ListAll();
                DataRow drMinAlt = dtMineralMinAlt.NewRow();
                drMinAlt[0] = "-1";
                drMinAlt[1] = "Select an option..";
                dtMineralMinAlt.Rows.Add(drMinAlt);
                LoadCombos(dtMineralMinAlt, cmbMineralPh);
                LoadCombos(dtMineralMinAlt, cmbMineralMt);                


                DataTable dtMineralAlt = new DataTable();
                dtMineralAlt = oRf.getRfMinerMinAlt_List();
                DataRow drMA = dtMineralAlt.NewRow();
                drMA[0] = "-1";
                drMA[1] = "Select an option..";
                dtMineralAlt.Rows.Add(drMA);
                LoadCombos(dtMineralAlt, cmbMin1Alt);
                LoadCombos(dtMineralAlt, cmbMin2Alt1);
                LoadCombos(dtMineralAlt, cmbMin3Alt1);



                
                #endregion

                #region Intensity
                DataTable dtMinStyle = new DataTable();
                dtMinStyle = oRf.getRfMinerMinSt_List();
                DataRow drMinStyle = dtMinStyle.NewRow();
                drMinStyle[0] = "-1";
                drMinStyle[1] = "Select an option..";
                dtMinStyle.Rows.Add(drMinStyle);
                LoadCombos(dtMinStyle, cmbStyleM);


                DataTable dtMinInt = new DataTable();
                dtMinInt = oRf.getRfOxidationInt_List();
                DataRow drMin = dtMinInt.NewRow();
                drMin[0] = "-1";
                drMin[1] = "Select an option..";
                dtMinInt.Rows.Add(drMin);

                LoadCombos(dtMinInt, cmbStyleGoe);
                LoadCombos(dtMinInt, cmbStyleHem);
                LoadCombos(dtMinInt, cmbStyleJar);
                LoadCombos(dtMinInt, cmbStyleLim);


                DataTable dtStyleAlt = new DataTable();
                dtStyleAlt = oRf.getRfStyleAlt_List();
                DataRow drStyleA = dtStyleAlt.NewRow();
                drStyleA[0] = "-1";
                drStyleA[1] = "Select an option..";
                dtStyleAlt.Rows.Add(drStyleA);
                LoadCombos(dtStyleAlt, cmbStyleAlt1);

                #endregion

                #region Structure
                DataTable dtFillStr = new DataTable();
                dtFillStr = oRf.getRfFillStructure_List();
                DataRow drFill = dtFillStr.NewRow();
                drFill[0] = "-1";
                drFill[1] = "Select an option..";
                dtFillStr.Rows.Add(drFill);
                cmbFillSt.DisplayMember = "Comb";
                cmbFillSt.ValueMember = "Code";
                cmbFillSt.DataSource = dtFillStr;
                cmbFillSt.SelectedValue = "-1";

                cmbFillSt2.DisplayMember = "Comb";
                cmbFillSt2.ValueMember = "Code";
                cmbFillSt2.DataSource = dtFillStr.Copy();
                cmbFillSt2.SelectedValue = "-1";

                cmbFillSt3.DisplayMember = "Comb";
                cmbFillSt3.ValueMember = "Code";
                cmbFillSt3.DataSource = dtFillStr.Copy();
                cmbFillSt3.SelectedValue = "-1";



                DataTable dtStructType = new DataTable();
                dtStructType = oRf.getRfTypeStructure_List();
                DataRow drS = dtStructType.NewRow();
                drS[0] = "-1";
                drS[1] = "Select an option..";
                dtStructType.Rows.Add(drS);
                cmbStructureTypeSt.DisplayMember = "Comb";
                cmbStructureTypeSt.ValueMember = "Code";
                cmbStructureTypeSt.DataSource = dtStructType;
                cmbStructureTypeSt.SelectedValue = "-1";
                #endregion

                #region Alteration Type,Intensity
                DataTable dtAlt = new DataTable();
                dtAlt = oRf.getRfTypeAlt_List();
                DataRow drAlt = dtAlt.NewRow();
                drAlt[0] = "-1";
                drAlt[1] = "Select an option..";
                dtAlt.Rows.Add(drAlt);
                LoadCombos(dtAlt, cmbTypeAlt);

                DataTable dtIntensity = new DataTable();
                dtIntensity = oRf.getRfIntensityAlt_List(ConfigurationSettings.AppSettings["IDProjectGC"].ToString());
                DataRow drInt = dtIntensity.NewRow();
                drInt[0] = "-1";
                drInt[1] = "Select an option..";
                dtIntensity.Rows.Add(drInt);
                LoadCombos(dtIntensity, cmbIntAlt);
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private DataTable LoadDataCHAll(string _sOpcion)
        {
            try
            {
                oCHSamp.sOpcion = _sOpcion.ToString();
                oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
                DataTable dtCH = oCHSamp.getCHSamplesByChid();
                return dtCH;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private DataTable LoadDataCH(string _sSample)
        {
            try
            {
                oCHSamp.sSample = _sSample.ToString();
                oCHSamp.sChId = txtChId.Text.ToString();
                DataTable dtCH = oCHSamp.getCHSamplesBySample();
                return dtCH;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private DataTable LoadDataCHSurv(string _sSample)
        {
            try
            {
                oCHSamp.sSample = _sSample.ToString();
                oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
                DataTable dtCH = oCHSamp.getCHSamplesBySample();
                return dtCH;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void CargarCombosPerc(DataTable _dt, ComboBox _cbox)
        {
            try
            {
                if (_dt.Rows.Count > 0)
                {
                    _cbox.DataSource = _dt.Copy();
                    _cbox.ValueMember = _dt.Columns[0].ToString();
                    _cbox.DisplayMember = _dt.Columns[1].ToString();
                    _cbox.SelectedValue = "-1";
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void LoadCombos(DataTable _dt, ComboBox _cbox)
        {
            try
            {
                if (_dt.Rows.Count > 0)
                {
                    _cbox.DataSource = _dt.Copy();
                    _cbox.ValueMember = _dt.Columns[0].ToString();
                    _cbox.DisplayMember = _dt.Columns[1].ToString();
                    _cbox.SelectedValue = "-1";
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void LoadDgChannels()
        {
            try
            {
                oCh.sChId = "1";
                oCh.sOpcion = "1";
                dgDataCh.DataSource = oCh.getCH_Collars();
                dgDataCh.Columns["SKCHChannels"].Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private AutoCompleteStringCollection AutoCompleteCmb(DataTable _dtAutoComplete, string _sRow)
        {
            try
            {

                AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();
                foreach (DataRow row in _dtAutoComplete.Rows)
                {
                    stringCol.Add(Convert.ToString(row[_sRow]));
                }

                return stringCol;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void frmChannels_Load(object sender, EventArgs e)
        {

        }

        private void CleanControls()
        {
            try
            {
                txtChId.Text = "";
                txtLenghtCh.Text = "";
                txtEastCh.Text = "";
                txtNorthCh.Text = "";
                txtElevationCh.Text = "";
                txtProjectionCh.Text = "";
                txtDatumCh.Text = "";
                txtProjectCh.Text = ConfigurationSettings.AppSettings["IDProjectGC"].ToString();
                txtClaimCh.Text = "";
                dtStartDateCh.Text = "01/01/1900";
                dtFinalDateCh.Text = "01/01/1900";
                //txtPurposeCh.Text = "";
                txtStorageCh.Text = "";
                txtSourceCh.Text = "";
                //cmbLocationChannel.Text = "Select an option...";
                txtCommentsCh.Text = "";
                cmbMineEntrance.SelectedValue = "-1";
                cmbChannelType.SelectedValue = "-1";
                txtTotalSamples.Text = "";
                sEditCh = "0";

                dTimerDateSur.Text = "01/01/1900";
                cmbInstSur.Text = "TS";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private bool Keypress(KeyPressEventArgs e)
        {

            if (Char.IsNumber(e.KeyChar))
            {
                return false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                return true;
            }

            return false;
        }

        private string ControlsValidateChannel()
        {
            try
            {
                string sresp = "";

                if (cmbMineEntrance.SelectedValue.ToString() == "-1" ||
                    cmbMineEntrance.SelectedValue.ToString() == "")
                {
                    sresp += "Empty Mine Entrance. " + Environment.NewLine;   
                }

                //if (txtChId.Text.ToString() == "")
                //{
                //    sresp += "Empty CHId. " + Environment.NewLine;   
                //}

                //if (txtEastCh.Text.ToString() == "")
                //{
                //    sresp += "Empty East. " + Environment.NewLine;
                //}

                //if (txtNorthCh.Text.ToString() == "")
                //{
                //    sresp += "Empty North. " + Environment.NewLine;
                //}

                //if (txtElevationCh.Text.ToString() == "")
                //{
                //    sresp += "Empty Elevation. " + Environment.NewLine;
                //}

                //if (txtTotalSamples.Text.ToString() == "")
                //{
                //    sresp += "Empty Total Samples. " + Environment.NewLine;
                //}

                //if (txtTotalSamples.Text.ToString() == "0")
                //{
                //    sresp += "You must enter to Total Samples greater than Zero. " + Environment.NewLine;
                //}

                return sresp;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnAddCh_Click(object sender, EventArgs e)
        {
            try
            {
                string sResp = ControlsValidateChannel().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Channels", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                if (sEditCh == "0")
                {
                    oCh.sOpcion = "1";
                    oCh.iSKCHChannels = 0;
                }
                else if (sEditCh == "1")
                {
                    oCh.sOpcion = "2";
                }
                
                oCh.sChId = txtChId.Text.ToString();

                if (txtEastCh.Text == "")
                    oCh.dEast = null;
                else
                    oCh.dEast = double.Parse(txtEastCh.Text.ToString());

                if (txtNorthCh.Text == "")
                    oCh.dNorth = null;
                else
                    oCh.dNorth = double.Parse(txtNorthCh.Text.ToString());

                if (txtElevationCh.Text == "")
                    oCh.dElevation = null;
                else
                    oCh.dElevation = double.Parse(txtElevationCh.Text.ToString());

                if (txtLenghtCh.Text == "")
                    oCh.dLenght = null;
                else
                    oCh.dLenght = double.Parse(txtLenghtCh.Text.ToString());

                oCh.sProjection = txtProjectionCh.Text.ToString();
                oCh.sDatum = txtDatumCh.Text.ToString();
                oCh.sProject = txtProjectCh.Text.ToString();
                oCh.sClaim = txtClaimCh.Text.ToString();

                if (dtStartDateCh.Text == "01/01/1900")
                    oCh.sStartDate = null;
                else
                    oCh.sStartDate = dtStartDateCh.Value.ToShortDateString();

                if (dtFinalDateCh.Text == "01/01/1900")
                    oCh.sFinalDate = null;
                else
                    oCh.sFinalDate = dtFinalDateCh.Value.ToShortDateString();

                //oCh.sPurpose = txtPurposeCh.Text.ToString();
                oCh.sStorage = txtStorageCh.Text.ToString();
                oCh.sSource = txtSourceCh.Text.ToString();
                
                oCh.sComments = txtCommentsCh.Text.ToString();

                if (cmbMineEntrance.SelectedValue.ToString() == "-1" ||
                    cmbMineEntrance.SelectedValue.ToString() == "")
                    oCh.sMineID = null;
                else
                    oCh.sMineID = cmbMineEntrance.SelectedValue.ToString();


                if (cmbChannelType.SelectedValue.ToString() == "-1" ||
                    cmbChannelType.SelectedValue.ToString() == "")
                    oCh.sType = null;
                else
                    oCh.sType = cmbChannelType.SelectedValue.ToString();

                if (cmbInstSur.Text.ToString()  == "")
                    oCh.sInstrument = null;
                else
                    oCh.sInstrument = cmbInstSur.Text.ToString();

                if (dTimerDateSur.Text == "01/01/1900")
                    oCh.sDate_Survey = null;
                else
                    oCh.sDate_Survey = dTimerDateSur.Value.ToShortDateString();
              
                if (txtTotalSamples.Text.ToString() == "")
                    oCh.iSamplesTotal = null;
                else
                    oCh.iSamplesTotal = int.Parse(txtTotalSamples.Text.ToString());

                string sRespAdd = oCh.CH_Collars_Add();
                if (sRespAdd == "OK")
                {
                    LoadDgChannels();
                    LoadChannelId();
                    cmbChannelId.SelectedValue = txtChId.Text.ToString();
                    CleanControls();
                    btnCancelCh_Click(null, null);
                }
                else
                {
                    MessageBox.Show("Error: " + sResp);
                    sEditCh = "0";
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgDataCh_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                oCh.iSKCHChannels = Int64.Parse(dgDataCh.Rows[e.RowIndex].Cells["SKCHChannels"].Value.ToString());
                sEditCh = "1";

                cmbChannelId.SelectedValue = dgDataCh.Rows[e.RowIndex].Cells["Chid"].Value.ToString();
                txtChId.Text = dgDataCh.Rows[e.RowIndex].Cells["Chid"].Value.ToString();
                txtLenghtCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Length"].Value.ToString();
                txtEastCh.Text = dgDataCh.Rows[e.RowIndex].Cells["East"].Value.ToString();
                txtNorthCh.Text = dgDataCh.Rows[e.RowIndex].Cells["North"].Value.ToString();
                txtElevationCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Elevation"].Value.ToString();
                txtProjectionCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Projection"].Value.ToString();
                txtDatumCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Datum"].Value.ToString();
                txtProjectCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Project"].Value.ToString();
                txtClaimCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Claim"].Value.ToString();

                dtStartDateCh.Text =
                    dgDataCh.Rows[e.RowIndex].Cells["Star_Date"].Value.ToString() == ""
                    ? DateTime.Now.ToShortDateString()
                    : dtStartDateCh.Text = Convert.ToDateTime(dgDataCh.Rows[e.RowIndex].Cells["Star_Date"].Value).ToString("dd/MM/yyyy");

                dtFinalDateCh.Text =
                    dgDataCh.Rows[e.RowIndex].Cells["Final_Date"].Value.ToString() == ""
                    ? DateTime.Now.ToShortDateString()
                    : dtFinalDateCh.Text = Convert.ToDateTime(dgDataCh.Rows[e.RowIndex].Cells["Final_Date"].Value).ToString("dd/MM/yyyy");

                //txtPurposeCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Purpose"].Value.ToString();
                txtStorageCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Storage"].Value.ToString();
                txtSourceCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Source"].Value.ToString();

                txtCommentsCh.Text = dgDataCh.Rows[e.RowIndex].Cells["Comments"].Value.ToString();
                
                cmbMineEntrance.SelectedValue = dgDataCh.Rows[e.RowIndex].Cells["MineID"].Value.ToString() == "" ? "-1" :
                    dgDataCh.Rows[e.RowIndex].Cells["MineID"].Value.ToString();

                cmbChannelType.SelectedValue = dgDataCh.Rows[e.RowIndex].Cells["Type"].Value.ToString() == "" ? "-1" :
                    dgDataCh.Rows[e.RowIndex].Cells["Type"].Value.ToString();

                cmbInstSur.Text = dgDataCh.Rows[e.RowIndex].Cells["Instrument"].Value.ToString() == "" ? "TS" :
                    dgDataCh.Rows[e.RowIndex].Cells["Instrument"].Value.ToString();

                dTimerDateSur.Text =
                   dgDataCh.Rows[e.RowIndex].Cells["Date_Survey"].Value.ToString() == ""
                   ? DateTime.Now.ToShortDateString()
                   : dTimerDateSur.Text = Convert.ToDateTime(dgDataCh.Rows[e.RowIndex].Cells["Date_Survey"].Value).ToString("dd/MM/yyyy");


                txtTotalSamples.Text = dgDataCh.Rows[e.RowIndex].Cells["SamplesTotal"].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelCh_Click(object sender, EventArgs e)
        {
            try
            {
                CleanControls();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgDataCh_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "ChId " + dgDataCh.Rows[e.RowIndex].Cells["Chid"].Value.ToString()
                   + " Lenght " + dgDataCh.Rows[e.RowIndex].Cells["Length"].Value.ToString()
                   + " East " + dgDataCh.Rows[e.RowIndex].Cells["East"].Value.ToString()
                   + " North " + dgDataCh.Rows[e.RowIndex].Cells["North"].Value.ToString()
                   , " Channels ", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oCh.iSKCHChannels = Int64.Parse(dgDataCh.Rows[e.RowIndex].Cells["SKCHChannels"].Value.ToString());
                    string sResp = oCh.CH_Collars_Delete();
                    if (sResp == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Channels", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadDgChannels();
                        CleanControls();
                        sEditCh = "0";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CleanControlsSurvey()
        {
            try
            {
                //txtChId.Text = "";
                txtToSur.Text = "";
                txtAzimuthSur.Text = "";
                txtDipSur.Text = "";
                dTimerDateSur.Text = "01/01/1990";
                cmbInstSur.Text = "";
                cmbMineLocation.Text = "";
                txtP1E.Text = "";
                txtP1N.Text = "";
                txtP1Z.Text = "";
                txtP2E.Text = "";
                txtP2N.Text = "";
                txtP2Z.Text = "";
                txtValidatedby.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadDgSurveys()
        {
            try
            {
                oSur.sChId = cmbChannelId.SelectedValue.ToString();
                if (cmbSample.SelectedValue == null)
                {
                    return;
                }
                if (cmbSample.SelectedValue.ToString() == "" ||
                    cmbSample.SelectedValue.ToString() == "-1")
                {
                    return;
                }
                oSur.sSample = cmbSample.SelectedValue.ToString();
                oSur.sOpcion = "2";
                dgSurvey.DataSource = oSur.getCH_Surveys();
                dgSurvey.Columns["SKCHSurveys"].Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddSur_Click(object sender, EventArgs e)
        {
            try
            {
                if (sEditSur == "0")
                {
                    oSur.sOpcion = "1";
                    oSur.iSKCHSurveys = 0;
                }
                else if (sEditSur == "1")
                {
                    oSur.sOpcion = "2";
                }

                if (cmbChannelId.SelectedValue.ToString() == "Select an option.." ||
                    cmbSample.SelectedValue.ToString() == "Select an option.." )
                {
                    MessageBox.Show("Empty ChannelId or Sample");
                    return;
                }

                oSur.sChId = cmbChannelId.SelectedValue.ToString();
                oSur.sSample = cmbSample.SelectedValue.ToString();
                //oSur.dTo = double.Parse(txtToSur.Text.ToString());
                if (txtToSur.Text.ToString() == "")
                    oSur.dTo = null;
                else
                    oSur.dTo = double.Parse(txtToSur.Text.ToString());

                //oSur.dAzm = double.Parse(txtAzimuthSur.Text.ToString());
                if (txtAzimuthSur.Text.ToString() == "")
                    oSur.dAzm = null;
                else
                    oSur.dAzm = double.Parse(txtAzimuthSur.Text.ToString());

                if (txtDipSur.Text.ToString() == "")
                    oSur.dDip = null;
                else
                    oSur.dDip = double.Parse(txtDipSur.Text.ToString());

                if (txtDipSur.Text.ToString() == "")
                    oSur.dDip = null;
                else
                    oSur.dDip = double.Parse(txtDipSur.Text.ToString());

                if (cmbMineLocation.Text.ToString() == "")
                    oSur.sMineLocation = null;
                else
                    oSur.sMineLocation = cmbMineLocation.Text.ToString();

                if (txtValidatedby.Text.ToString() == "")
                    oSur.sValidatedby = null;
                else
                    oSur.sValidatedby = txtValidatedby.Text.ToString();


                string sResp = oSur.CH_Surveys_Add();
                if (sResp == "OK")
                {
                    LoadDgSurveys();
                    CleanControlsSurvey();
                    sEditSur = "0";
                }

                sEditSur = "0";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgSurvey_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditSur = "1";
                oSur.iSKCHSurveys = Int64.Parse(dgSurvey.Rows[e.RowIndex].Cells["SKCHSurveys"].Value.ToString());
                txtChId.Text = dgSurvey.Rows[e.RowIndex].Cells["Chid"].Value.ToString();
                txtToSur.Text = dgSurvey.Rows[e.RowIndex].Cells["To"].Value.ToString();
                txtAzimuthSur.Text = dgSurvey.Rows[e.RowIndex].Cells["Azm"].Value.ToString();
                txtDipSur.Text = dgSurvey.Rows[e.RowIndex].Cells["Dip"].Value.ToString();

                //if (dgSurvey.Rows[e.RowIndex].Cells["Instrument"].Value.ToString() == "")
                //    cmbInstSur.Text = "-1";
                //else
                //    cmbInstSur.Text = dgSurvey.Rows[e.RowIndex].Cells["Instrument"].Value.ToString();

                if (dgSurvey.Rows[e.RowIndex].Cells["MineLocation"].Value.ToString() == "")
                    cmbMineLocation.Text = "";
                else
                    cmbMineLocation.Text = dgSurvey.Rows[e.RowIndex].Cells["MineLocation"].Value.ToString();

                txtP1E.Text = dgSurvey.Rows[e.RowIndex].Cells["P1E"].Value.ToString();
                txtP1N.Text = dgSurvey.Rows[e.RowIndex].Cells["P1N"].Value.ToString();
                txtP1Z.Text = dgSurvey.Rows[e.RowIndex].Cells["P1Z"].Value.ToString();
                txtP2E.Text = dgSurvey.Rows[e.RowIndex].Cells["P2E"].Value.ToString();
                txtP2N.Text = dgSurvey.Rows[e.RowIndex].Cells["P2N"].Value.ToString();
                txtP2Z.Text = dgSurvey.Rows[e.RowIndex].Cells["P2Z"].Value.ToString();
                txtValidatedby.Text = dgSurvey.Rows[e.RowIndex].Cells["Validated_by"].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadCmbSamples()
        {
            try
            {
                oCHSamp.sOpcion = "1";
                oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
                DataTable dtSamp = new DataTable();
                dtSamp = oCHSamp.getCHSamplesByChid();
                DataRow dr = dtSamp.NewRow();
                dr["Sample"] = "Select an option..";
                dtSamp.Rows.Add(dr);
                cmbSample.DisplayMember = "Sample";
                cmbSample.ValueMember = "Sample";
                cmbSample.DataSource = dtSamp;
                cmbSample.SelectedValue = "Select an option..";

                if (sSampleSelect != "" && sSampleSelect != "Select an option..")
                {
                    cmbSample.SelectedValue = sSampleSelect.ToString();
                    sSampleSelect = "";
                }

                //cmbChannelId.AutoCompleteCustomSource = AutoCompleteCmb(dtSamp, "Sample");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void txtChId_TextChanged(object sender, EventArgs e)
        {
            try
            {
                LoadDgSurveys();
                LoadCmbSamples();
                dgData.DataSource = LoadDataCHAll("1");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgSurvey_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "ChId " + dgSurvey.Rows[e.RowIndex].Cells["Chid"].Value.ToString()
                   + " To " + dgSurvey.Rows[e.RowIndex].Cells["To"].Value.ToString()
                   + " Azimuth " + dgSurvey.Rows[e.RowIndex].Cells["Azm"].Value.ToString()
                   + " Dip " + dgSurvey.Rows[e.RowIndex].Cells["Dip"].Value.ToString()
                   , " Surveys ", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oSur.iSKCHSurveys = Int64.Parse(dgSurvey.Rows[e.RowIndex].Cells["SKCHSurveys"].Value.ToString());
                    string sResp = oSur.CH_Surveys_Delete();
                    if (sResp == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Surveys", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadDgSurveys();
                        CleanControlsSurvey();
                        sEditSur = "0";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtFromHeader_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtToHeader_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private string ControlsValidateHeader()
        {
            try
            {
                string sresp = "";

                if (txtSampleHead.Text.ToString() == "")
                {
                    sresp += "Empty Sample. " + Environment.NewLine;
                }

                if (cmbChannelId.SelectedValue == "Select an option.." ||
                    cmbChannelId.SelectedValue == "")
                {
                    sresp += "Empty Chid. " + Environment.NewLine;
                }

                return sresp;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                string sResp = ControlsValidateHeader().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Samples", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                AddHeaderLithology();
                LoadCmbSamples();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CleanControlsHeader()
        {
            try
            {
                oCHSamp.iSKCHSamples = 0;
                sEdit = "0";
                sSampleSelect = "";

                txtFromHeader.Text = txtToHeader.Text;
                txtToHeader.Text = "";
                dtimeDate.Text = DateTime.Now.ToShortDateString();
                txtSampleHead.Text = "";
                cmbTarget.SelectedValue = "-1";
                //txtLocation.Text = "";
                txtProject.Text = ConfigurationSettings.AppSettings["IDProjectGC"].ToString();
                cmbGeologist.SelectedValue = "-1";
                txtHelper.Text = "";
                txtStation.Text = "";
                txtCoordE.Text = "";
                txtCoordN.Text = "";
                txtCoordElevation.Text = "";

                txtCoordE2.Text = "";
                txtCoordN2.Text = "";
                txtCoordZ2.Text = "";


                cmbCS.SelectedValue = "-1";
                txtGPSEPE.Text = "";
                txtPhoto.Text = "";
                txtPhotoAzimuth.Text = "";
                cmbSampleType.SelectedValue = "-1";
                cmbSamplingType.SelectedValue = "-1";
                txtDupOf.Text = "";
                cmbNotInSitu.SelectedValue = "-1";
                cmbPorpuose.SelectedValue = "-1";
                cmbRelativeLoc.SelectedValue = "-1";
                txtLenght.Text = "";
                txtHigh.Text = "";
                txtThickness.Text = "";
                txtObservations.Text = "";
                cmbLithologyLit.SelectedValue = "-1";
                cmbLTextures.SelectedValue = "-1";
                cmbLGsize.SelectedValue = "-1";
                cmbLWeathering.SelectedValue = "-1";
                cmbRSorting.SelectedValue = "-1";
                cmbRSphericity.SelectedValue = "-1";
                cmbRounding.SelectedValue = "-1";
                txtObservSedimentary.Text = "";
                txtMatrixPerc.Text = "";
                cmbMatrixGSize.SelectedValue = "-1";
                txtMatrixObserv.Text = "";
                txtPhenoPerc.Text = "";
                cmbPhenoGSize.SelectedValue = "-1";
                txtPhenoObserv.Text = "";
                cmbContactType.SelectedValue = "-1";
                cmbVeinName.SelectedValue = "-1";
                cmbHostRock.SelectedValue = "-1";
                txtVeinObserv.Text = "";
                //txtMine.Text = "";
                //txtMineEntrance.Text = "";

                chkValidated.Checked = false;

                txtCountSample.Text = "";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void AddHeaderLithology()
        {
            try
            {

                if (txtPhenoPerc.Text != "")
                {
                    if (double.Parse(txtPhenoPerc.Text.ToString()) > 100)
                    {
                        MessageBox.Show("Perc Pheno > 100");
                        txtPhenoPerc.Text = "";
                        txtPhenoPerc.Focus();
                        return;
                    }
                }

                if (txtMatrixPerc.Text != "")
                {
                    if (double.Parse(txtMatrixPerc.Text.ToString()) > 100)
                    {
                        MessageBox.Show("Perc Matrix > 100");
                        txtPhenoPerc.Text = "";
                        txtPhenoPerc.Focus();
                        return;
                    }
                }

                if (txtPhenoPerc.Text != "" && txtMatrixPerc.Text != "")
                {
                    if (double.Parse(txtPhenoPerc.Text.ToString()) +
                        double.Parse(txtMatrixPerc.Text.ToString()) > 100)
                    {
                        MessageBox.Show("Perc Pheno + Perc Matrix > 100");
                        txtPhenoPerc.Text = "";
                        txtPhenoPerc.Focus();
                        return;
                    }

                }

                if (sEdit == "0")
                {
                    oCHSamp.sOpcion = "1";
                    oCHSamp.iSKCHSamples = 0;
                }
                else if (sEdit == "1")
                {
                    oCHSamp.sOpcion = "2";
                }

                oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
                oCHSamp.sSample = txtSampleHead.Text.ToString();

                if (txtFromHeader.Text.ToString() == "")
                    oCHSamp.dFrom = null;
                else
                    oCHSamp.dFrom = double.Parse(txtFromHeader.Text.ToString());

                if (txtToHeader.Text.ToString() == "")
                    oCHSamp.dTo = null;
                else
                    oCHSamp.dTo = double.Parse(txtToHeader.Text.ToString());

                if (cmbTarget.SelectedValue != null)
                {
                    if (cmbTarget.SelectedValue.ToString() == "-1" ||
                    cmbTarget.SelectedValue.ToString() == "")
                        oCHSamp.sTarget = null;
                    else
                        oCHSamp.sTarget = cmbTarget.SelectedValue.ToString();
                }
                else oCHSamp.sTarget = null;

                //oCHSamp.sLocation = txtLocation.Text.ToString();
                oCHSamp.sProject = txtProject.Text.ToString();

                if (cmbGeologist.SelectedValue != null)
                {
                    if (cmbGeologist.SelectedValue.ToString() == "-1" ||
                    cmbGeologist.SelectedValue.ToString() == "")
                        oCHSamp.sGeologist = null;
                    else
                        oCHSamp.sGeologist = cmbGeologist.SelectedValue.ToString();

                }
                else oCHSamp.sGeologist = null;

                
                oCHSamp.sHelper = txtHelper.Text.ToString();
                oCHSamp.sStation = txtStation.Text.ToString();

                DateTime dDate = DateTime.Parse(dtimeDate.Value.ToString());
                string sDate = dDate.Year.ToString().PadLeft(4, '0') + dDate.Month.ToString().PadLeft(2, '0') +
                    dDate.Day.ToString().PadLeft(2, '0');
                oCHSamp.sDate = sDate.ToString();

                if (txtCoordE.Text.ToString() == "")
                    oCHSamp.dE = null;
                else
                    oCHSamp.dE = double.Parse(txtCoordE.Text.ToString());

                if (txtCoordN.Text.ToString() == "")
                    oCHSamp.dN = null;
                else
                    oCHSamp.dN = double.Parse(txtCoordN.Text.ToString());

                if (txtCoordElevation.Text.ToString() == "")
                    oCHSamp.dZ = null;
                else
                    oCHSamp.dZ = double.Parse(txtCoordElevation.Text.ToString());



                if (txtCoordE2.Text.ToString() == "")
                    oCHSamp.dE2 = null;
                else
                    oCHSamp.dE2 = double.Parse(txtCoordE2.Text.ToString());

                if (txtCoordN2.Text.ToString() == "")
                    oCHSamp.dN2 = null;
                else
                    oCHSamp.dN2 = double.Parse(txtCoordN2.Text.ToString());

                if (txtCoordZ2.Text.ToString() == "")
                    oCHSamp.dZ2 = null;
                else
                    oCHSamp.dZ2 = double.Parse(txtCoordZ2.Text.ToString());



                if (cmbCS.SelectedValue.ToString() == "-1" ||
                    cmbCS.SelectedValue.ToString() == "")
                    oCHSamp.sCS = null;
                else
                    oCHSamp.sCS = cmbCS.SelectedValue.ToString();

                if (txtGPSEPE.Text.ToString() == "")
                    oCHSamp.dGPSEpe = null;
                else
                    oCHSamp.dGPSEpe = double.Parse(txtGPSEPE.Text.ToString());

                if (txtPhoto.Text.ToString() == "")
                    oCHSamp.sPhoto = null;
                else
                    oCHSamp.sPhoto = txtPhoto.Text.ToString();

                if (txtPhotoAzimuth.Text.ToString() == "")
                    oCHSamp.sPhotoAzimuth = null;
                else
                    oCHSamp.sPhotoAzimuth = txtPhotoAzimuth.Text.ToString();

                if (cmbSampleType.SelectedValue.ToString() == "-1" ||
                    cmbSampleType.SelectedValue.ToString() == "")
                    oCHSamp.sSampleType = null;
                else
                    oCHSamp.sSampleType = cmbSampleType.SelectedValue.ToString();

                if (cmbSamplingType.SelectedValue.ToString() == "-1" ||
                    cmbSamplingType.SelectedValue.ToString() == "")
                    oCHSamp.sSamplingType = null;
                else
                    oCHSamp.sSamplingType = cmbSamplingType.SelectedValue.ToString();

                if (cmbNotInSitu.SelectedValue.ToString() == "-1" ||
                    cmbNotInSitu.SelectedValue.ToString() == "")
                    oCHSamp.sNotInSitu = null;
                else
                    oCHSamp.sNotInSitu = cmbNotInSitu.SelectedValue.ToString();

                if (cmbPorpuose.SelectedValue.ToString() == "-1" ||
                    cmbPorpuose.SelectedValue.ToString() == "")
                    oCHSamp.sPorpouse = null;
                else
                    oCHSamp.sPorpouse = cmbPorpuose.SelectedValue.ToString();

                if (cmbRelativeLoc.SelectedValue.ToString() == "-1" ||
                    cmbRelativeLoc.SelectedValue.ToString() == "")
                    oCHSamp.sRelativeLoc = null;
                else
                    oCHSamp.sRelativeLoc = cmbRelativeLoc.SelectedValue.ToString();

                if (txtLenght.Text.ToString() == "")
                    oCHSamp.dLenght = null;
                else
                    oCHSamp.dLenght = double.Parse(txtLenght.Text.ToString());

                if (txtHigh.Text.ToString() == "")
                    oCHSamp.dHigh = null;
                else
                    oCHSamp.dHigh = double.Parse(txtHigh.Text.ToString());

                if (txtThickness.Text.ToString() == "")
                    oCHSamp.sThickness = null;
                else
                    oCHSamp.sThickness = txtThickness.Text.ToString();

                if (txtObservations.Text.ToString() == "")
                    oCHSamp.sObservations = null;
                else
                    oCHSamp.sObservations = txtObservations.Text.ToString();

                if (cmbLithologyLit.SelectedValue.ToString() == "" ||
                    cmbLithologyLit.SelectedValue.ToString() == "-1")
                    oCHSamp.sLRock = null;
                else
                    oCHSamp.sLRock = cmbLithologyLit.SelectedValue.ToString();

                if (cmbLTextures.SelectedValue.ToString() == "-1" || cmbLTextures.SelectedValue.ToString() == "")
                    oCHSamp.sLTexture = null;
                else
                    oCHSamp.sLTexture = cmbLTextures.SelectedValue.ToString();

                if (cmbLGsize.SelectedValue.ToString() == "-1" || cmbLGsize.SelectedValue.ToString() == "")
                    oCHSamp.sLGSize = null;
                else
                    oCHSamp.sLGSize = cmbLGsize.SelectedValue.ToString();

                if (cmbLWeathering.SelectedValue.ToString() == "-1" || cmbLWeathering.SelectedValue.ToString() == "")
                    oCHSamp.sLWeathering = null;
                else
                    oCHSamp.sLWeathering = cmbLWeathering.SelectedValue.ToString();

                if (cmbRSorting.SelectedValue.ToString() == "-1" || cmbRSorting.SelectedValue.ToString() == "")
                    oCHSamp.sLRockSorting = null;
                else
                    oCHSamp.sLRockSorting = cmbRSorting.SelectedValue.ToString();

                if (cmbRSphericity.SelectedValue.ToString() == "-1" || cmbRSphericity.SelectedValue.ToString() == "")
                    oCHSamp.sLRockSphericity = null;
                else
                    oCHSamp.sLRockSphericity = cmbRSphericity.SelectedValue.ToString();

                if (cmbRounding.SelectedValue.ToString() == "-1" || cmbRounding.SelectedValue.ToString() == "")
                    oCHSamp.sLRockRounding = null;
                else
                    oCHSamp.sLRockRounding = cmbRounding.SelectedValue.ToString();

                if (txtObservSedimentary.Text.ToString() == "")
                    oCHSamp.sLRockObservation = null;
                else
                    oCHSamp.sLRockObservation = txtObservSedimentary.Text.ToString();

                

                if (cmbMatrixGSize.SelectedValue.ToString() == "-1" || cmbMatrixGSize.SelectedValue.ToString() == "")
                    oCHSamp.sLMatrixGSize = null;
                else
                    oCHSamp.sLMatrixGSize = cmbMatrixGSize.SelectedValue.ToString();

                if (txtMatrixObserv.Text.ToString() == "")
                    oCHSamp.sLMatrixObservations = null;
                else
                    oCHSamp.sLMatrixObservations = txtMatrixObserv.Text.ToString();




                if (txtMatrixPerc.Text.ToString() == "")
                    oCHSamp.sLMatrixPerc = null;
                else
                    oCHSamp.sLMatrixPerc = double.Parse(txtMatrixPerc.Text.ToString());

                if (txtPhenoPerc.Text.ToString() == "")
                    oCHSamp.sLPhenoCPerc = null;
                else
                    oCHSamp.sLPhenoCPerc = double.Parse(txtPhenoPerc.Text.ToString());




                if (cmbPhenoGSize.SelectedValue.ToString() == "-1" || cmbPhenoGSize.SelectedValue.ToString() == "")
                    oCHSamp.sLPhenoCGSize = null;
                else
                    oCHSamp.sLPhenoCGSize = cmbPhenoGSize.SelectedValue.ToString();

                if (txtPhenoObserv.Text.ToString() == "")
                    oCHSamp.sLPhenoCObservations = null;
                else
                    oCHSamp.sLPhenoCObservations = txtPhenoObserv.Text.ToString();

                if (cmbContactType.SelectedValue.ToString() == "" ||
                    cmbContactType.SelectedValue.ToString() == "-1")
                    oCHSamp.sVContactType = null;
                else
                    oCHSamp.sVContactType = cmbContactType.SelectedValue.ToString();

                if (cmbVeinName.SelectedValue.ToString() == "" || 
                    cmbVeinName.SelectedValue.ToString() == "-1")
                    oCHSamp.sVVeinName = null;
                else
                    oCHSamp.sVVeinName = cmbVeinName.SelectedValue.ToString();

                if (cmbHostRock.SelectedValue.ToString() == "" ||
                    cmbHostRock.SelectedValue.ToString() == "-1")
                    oCHSamp.sVHostRock = null;
                else
                    oCHSamp.sVHostRock = cmbHostRock.SelectedValue.ToString();

                if (txtVeinObserv.Text.ToString() == "")
                    oCHSamp.sVObservations = null;
                else
                    oCHSamp.sVObservations = txtVeinObserv.Text.ToString();

                if (cmbSamplingType.SelectedValue.ToString() == "" ||
                    cmbSamplingType.SelectedValue.ToString() == "-1")
                    oCHSamp.sSamplingType = null;
                else
                    oCHSamp.sSamplingType = cmbSamplingType.SelectedValue.ToString();

                if (txtDupOf.Text.ToString() == "")
                    oCHSamp.sDupOf = null;
                else
                    oCHSamp.sDupOf = txtDupOf.Text.ToString();


                oCHSamp.bValited = chkValidated.Checked;


                if (txtCountSample.Text == "")
                    oCHSamp.iSampleCont = null;
                else
                    oCHSamp.iSampleCont = int.Parse(txtCountSample.Text.ToString());


                string sResp = oCHSamp.CH_Samples_Add();
                if (sResp == "OK")
                {
                    
                    dgData.DataSource = LoadDataCHAll("1");
                    dgData.Columns["SKCHSamples"].Visible = false;

                    dgLithology.DataSource = LoadDataCHAll("1");
                    dgLithology.Columns["SKCHSamples"].Visible = false;

                    if (sEdit == "1")
                    {
                        if (dgData.Rows.Count > 1)
                        {
                            DataTable dtSamp = (DataTable)dgData.DataSource;
                            DataRow[] myRow = dtSamp.Select(@"SKCHSamples = '" + oCHSamp.iSKCHSamples.ToString() + "'");
                            int rowindex = dtSamp.Rows.IndexOf(myRow[0]);
                            dgData.Rows[rowindex].Selected = true;
                            dgData.CurrentCell = dgData.Rows[rowindex].Cells[1];

                            DataTable dtSamp2 = (DataTable)dgLithology.DataSource;
                            DataRow[] myRow2 = dtSamp2.Select(@"SKCHSamples = '" + oCHSamp.iSKCHSamples.ToString() + "'");
                            int rowindex2 = dtSamp2.Rows.IndexOf(myRow2[0]);
                            dgLithology.Rows[rowindex2].Selected = true;
                            dgLithology.CurrentCell = dgLithology.Rows[rowindex2].Cells[1];
                        }
                    }

                    CleanControlsHeader();
                    sEdit = "0";

                }
                else
                {
                    MessageBox.Show("Save Error: " + sResp.ToString(), "Channels", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


        private void dgD_CellClick(int _iRowIndex, DataGridView _dg, string _sFuente)
        {
            try
            {
                if (_sFuente == "Grid")
                {
                    sSampleSelect = _dg.Rows[_iRowIndex].Cells["Sample"].Value.ToString(); //Para guardar la seleccion de la muestra y utilizarla en el combo de sample    
                }
                

                oCHSamp.iSKCHSamples = Int64.Parse(_dg.Rows[_iRowIndex].Cells["SKCHSamples"].Value.ToString());
                sEdit = "1";

                cmbTarget.SelectedValue = _dg.Rows[_iRowIndex].Cells["Target"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["Target"].Value.ToString();
                cmbGeologist.SelectedValue = _dg.Rows[_iRowIndex].Cells["Geologist"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["Geologist"].Value.ToString();
                cmbCS.SelectedValue = _dg.Rows[_iRowIndex].Cells["CS"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["CS"].Value.ToString();
                cmbSampleType.SelectedValue = _dg.Rows[_iRowIndex].Cells["SampleType"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["SampleType"].Value.ToString();
                cmbSamplingType.SelectedValue = _dg.Rows[_iRowIndex].Cells["SamplingType"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["SampleType"].Value.ToString();
                cmbNotInSitu.SelectedValue = _dg.Rows[_iRowIndex].Cells["NotItSitu"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["NotItSitu"].Value.ToString();
                cmbPorpuose.SelectedValue = _dg.Rows[_iRowIndex].Cells["Porpuose"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["Porpuose"].Value.ToString();
                cmbRelativeLoc.SelectedValue = _dg.Rows[_iRowIndex].Cells["Relative_Loc"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["Relative_Loc"].Value.ToString();
                cmbLithologyLit.SelectedValue = _dg.Rows[_iRowIndex].Cells["LRock"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LRock"].Value.ToString();
                cmbLTextures.SelectedValue = _dg.Rows[_iRowIndex].Cells["LTexture"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LTexture"].Value.ToString();
                cmbLGsize.SelectedValue = _dg.Rows[_iRowIndex].Cells["LGSize"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LGSize"].Value.ToString();
                cmbLWeathering.SelectedValue = _dg.Rows[_iRowIndex].Cells["LWeathering"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LWeathering"].Value.ToString();
                cmbRSorting.SelectedValue = _dg.Rows[_iRowIndex].Cells["LRocksSorting"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LRocksSorting"].Value.ToString();
                cmbRSphericity.SelectedValue = _dg.Rows[_iRowIndex].Cells["LRocksSphericity"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LRocksSphericity"].Value.ToString();
                cmbRounding.SelectedValue = _dg.Rows[_iRowIndex].Cells["LRocksRounding"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LRocksRounding"].Value.ToString();
                txtMatrixPerc.Text = _dg.Rows[_iRowIndex].Cells["LMatrixPerc"].Value.ToString() == "" ? "" : _dg.Rows[_iRowIndex].Cells["LMatrixPerc"].Value.ToString();
                cmbMatrixGSize.SelectedValue = _dg.Rows[_iRowIndex].Cells["LMatrixGSize"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LMatrixGSize"].Value.ToString();
                txtPhenoPerc.Text= _dg.Rows[_iRowIndex].Cells["LPhenoCPerc"].Value.ToString() == "" ? "" : _dg.Rows[_iRowIndex].Cells["LPhenoCPerc"].Value.ToString();
                cmbPhenoGSize.SelectedValue = _dg.Rows[_iRowIndex].Cells["LPhenoCGSize"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["LPhenoCGSize"].Value.ToString();
                cmbSamplingType.SelectedValue = _dg.Rows[_iRowIndex].Cells["SamplingType"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["SamplingType"].Value.ToString();


                DateTime dDate =
                    _dg.Rows[_iRowIndex].Cells["Date"].Value.ToString() == ""
                    ? DateTime.Parse("1900/01/01")
                    : DateTime.Parse(_dg.Rows[_iRowIndex].Cells["Date"].Value.ToString());
                dtimeDate.Value = dDate;
                dtimeDate.Text = dtimeDate.Value.ToString();


                txtToHeader.Text = _dg.Rows[_iRowIndex].Cells["To"].Value.ToString();
                txtFromHeader.Text = _dg.Rows[_iRowIndex].Cells["From"].Value.ToString();
                dtimeDate.Text = _dg.Rows[_iRowIndex].Cells["Date"].Value.ToString();

                txtSampleHead.Text = _dg.Rows[_iRowIndex].Cells["Sample"].Value.ToString();
                cmbSample.SelectedValue = _dg.Rows[_iRowIndex].Cells["Sample"].Value.ToString();

                //txtLocation.Text = _dg.Rows[_iRowIndex].Cells["Location"].Value.ToString();
                txtProject.Text = _dg.Rows[_iRowIndex].Cells["Project"].Value.ToString();
                txtHelper.Text = _dg.Rows[_iRowIndex].Cells["Helper"].Value.ToString();
                txtStation.Text = _dg.Rows[_iRowIndex].Cells["Station"].Value.ToString();
                
                txtCoordE.Text = _dg.Rows[_iRowIndex].Cells["E"].Value.ToString();
                txtCoordN.Text = _dg.Rows[_iRowIndex].Cells["N"].Value.ToString();
                txtCoordElevation.Text = _dg.Rows[_iRowIndex].Cells["Z"].Value.ToString();

                txtCoordE2.Text = _dg.Rows[_iRowIndex].Cells["E2"].Value.ToString();
                txtCoordN2.Text = _dg.Rows[_iRowIndex].Cells["N2"].Value.ToString();
                txtCoordZ2.Text = _dg.Rows[_iRowIndex].Cells["Z2"].Value.ToString();

                txtGPSEPE.Text = _dg.Rows[_iRowIndex].Cells["GPSepe"].Value.ToString();
                txtPhoto.Text = _dg.Rows[_iRowIndex].Cells["Photo"].Value.ToString();
                txtPhotoAzimuth.Text = _dg.Rows[_iRowIndex].Cells["Photo_azimuth"].Value.ToString();
                txtDupOf.Text = _dg.Rows[_iRowIndex].Cells["DupOf"].Value.ToString();
                txtLenght.Text = _dg.Rows[_iRowIndex].Cells["length"].Value.ToString();
                txtHigh.Text = _dg.Rows[_iRowIndex].Cells["High"].Value.ToString();
                txtThickness.Text = _dg.Rows[_iRowIndex].Cells["Thickness"].Value.ToString();
                txtObservations.Text = _dg.Rows[_iRowIndex].Cells["Obsevations"].Value.ToString();
                txtObservSedimentary.Text = _dg.Rows[_iRowIndex].Cells["LRocksObservation"].Value.ToString();
                txtMatrixObserv.Text = _dg.Rows[_iRowIndex].Cells["LMatrixObsevations"].Value.ToString();
                txtPhenoObserv.Text = _dg.Rows[_iRowIndex].Cells["LPhenoCObsevations"].Value.ToString();
                cmbContactType.SelectedValue = _dg.Rows[_iRowIndex].Cells["VContactType"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["VContactType"].Value.ToString();
                cmbVeinName.SelectedValue = _dg.Rows[_iRowIndex].Cells["VVeinName"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["VVeinName"].Value.ToString();
                cmbHostRock.SelectedValue = _dg.Rows[_iRowIndex].Cells["VHostRock"].Value.ToString() == "" ? "-1" : _dg.Rows[_iRowIndex].Cells["VHostRock"].Value.ToString();
                txtVeinObserv.Text = _dg.Rows[_iRowIndex].Cells["VObsevations"].Value.ToString();
                //txtMine.Text = _dg.Rows[_iRowIndex].Cells["Mine"].Value.ToString();
                //txtMineEntrance.Text = _dg.Rows[_iRowIndex].Cells["MineEntrance"].Value.ToString();
                chkValidated.Checked = bool.Parse(_dg.Rows[_iRowIndex].Cells["Validated"].Value.ToString());

                txtCountSample.Text = _dg.Rows[_iRowIndex].Cells["SampleCont"].Value.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //oCHSamp.iSKCHSamples = Int64.Parse(dgData.Rows[e.RowIndex].Cells["SKCHSamples"].Value.ToString());
                //sEdit = "1";

                dgD_CellClick(e.RowIndex, dgData, "Grid");

                //cmbTarget.SelectedValue = dgData.Rows[e.RowIndex].Cells["Target"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["Target"].Value.ToString();
                //cmbGeologist.SelectedValue = dgData.Rows[e.RowIndex].Cells["Geologist"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["Geologist"].Value.ToString();
                //cmbCS.SelectedValue = dgData.Rows[e.RowIndex].Cells["CS"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["CS"].Value.ToString();
                //cmbSampleType.SelectedValue = dgData.Rows[e.RowIndex].Cells["SampleType"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["SampleType"].Value.ToString();
                //cmbSamplingType.SelectedValue = dgData.Rows[e.RowIndex].Cells["SampleType"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["SampleType"].Value.ToString();
                //cmbNotInSitu.SelectedValue = dgData.Rows[e.RowIndex].Cells["NotItSitu"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["NotItSitu"].Value.ToString();
                //cmbPorpuose.SelectedValue = dgData.Rows[e.RowIndex].Cells["Porpuose"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["Porpuose"].Value.ToString();
                //cmbRelativeLoc.SelectedValue = dgData.Rows[e.RowIndex].Cells["Relative_Loc"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["Relative_Loc"].Value.ToString();
                //cmbLithologyLit.SelectedValue = dgData.Rows[e.RowIndex].Cells["LRock"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LRock"].Value.ToString();
                //cmbLTextures.SelectedValue = dgData.Rows[e.RowIndex].Cells["LTexture"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LTexture"].Value.ToString();
                //cmbLGsize.SelectedValue = dgData.Rows[e.RowIndex].Cells["LGSize"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LGSize"].Value.ToString();
                //cmbLWeathering.SelectedValue = dgData.Rows[e.RowIndex].Cells["LWeathering"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LWeathering"].Value.ToString();
                //cmbRSorting.SelectedValue = dgData.Rows[e.RowIndex].Cells["LRocksSorting"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LRocksSorting"].Value.ToString();
                //cmbRSphericity.SelectedValue = dgData.Rows[e.RowIndex].Cells["LRocksSphericity"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LRocksSphericity"].Value.ToString();
                //cmbRounding.SelectedValue = dgData.Rows[e.RowIndex].Cells["LRocksRounding"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LRocksRounding"].Value.ToString();
                //cmbMatrixPorc.SelectedValue = dgData.Rows[e.RowIndex].Cells["LMatrixPerc"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LMatrixPerc"].Value.ToString();
                //cmbMatrixGSize.SelectedValue = dgData.Rows[e.RowIndex].Cells["LMatrixGSize"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LMatrixGSize"].Value.ToString();
                //cmbPhenoPerc.SelectedValue = dgData.Rows[e.RowIndex].Cells["LPhenoCPerc"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LPhenoCPerc"].Value.ToString();
                //cmbPhenoGSize.SelectedValue = dgData.Rows[e.RowIndex].Cells["LPhenoCGSize"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["LPhenoCGSize"].Value.ToString();
                //cmbSamplingType.SelectedValue = dgData.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString();

                ////dtimeDate.Text =
                ////        dgData.Rows[e.RowIndex].Cells["Date"].Value.ToString() == ""
                ////    ? "1900/01/01"
                ////    : dtimeDate.Text = dgData.Rows[e.RowIndex].Cells["Date"].Value.ToString();


                //DateTime dDate =
                //    dgData.Rows[e.RowIndex].Cells["Date"].Value.ToString() == ""
                //    ? DateTime.Parse("1900/01/01")
                //    : DateTime.Parse(dgData.Rows[e.RowIndex].Cells["Date"].Value.ToString());
                //dtimeDate.Value = dDate;
                //dtimeDate.Text = dtimeDate.Value.ToString();
                ////string sDate = dDate.Year.ToString().PadLeft(4, '0') + dDate.Month.ToString().PadLeft(2, '0') +
                ////    dDate.Day.ToString().PadLeft(2, '0');
                ////dtimeDate.Text = sDate.ToString();


                //txtToHeader.Text = dgData.Rows[e.RowIndex].Cells["To"].Value.ToString();
                //txtFromHeader.Text = dgData.Rows[e.RowIndex].Cells["From"].Value.ToString();
                //dtimeDate.Text = dgData.Rows[e.RowIndex].Cells["Date"].Value.ToString();
                
                //txtSampleHead.Text = dgData.Rows[e.RowIndex].Cells["Sample"].Value.ToString();
                //cmbSample.SelectedValue = dgData.Rows[e.RowIndex].Cells["Sample"].Value.ToString();
                //sSampleSelect = dgData.Rows[e.RowIndex].Cells["Sample"].Value.ToString(); //Para guardar la seleccion de la muestra y utilizarla en el combo de sample

                ////txtLocation.Text = dgData.Rows[e.RowIndex].Cells["Location"].Value.ToString();
                //txtProject.Text = dgData.Rows[e.RowIndex].Cells["Project"].Value.ToString();
                //txtHelper.Text = dgData.Rows[e.RowIndex].Cells["Helper"].Value.ToString();
                //txtStation.Text = dgData.Rows[e.RowIndex].Cells["Station"].Value.ToString();
                //txtCoordE.Text = dgData.Rows[e.RowIndex].Cells["E"].Value.ToString();
                //txtCoordN.Text = dgData.Rows[e.RowIndex].Cells["N"].Value.ToString();
                //txtCoordElevation.Text = dgData.Rows[e.RowIndex].Cells["Z"].Value.ToString();
                //txtGPSEPE.Text = dgData.Rows[e.RowIndex].Cells["GPSepe"].Value.ToString();
                //txtPhoto.Text = dgData.Rows[e.RowIndex].Cells["Photo"].Value.ToString();
                //txtPhotoAzimuth.Text = dgData.Rows[e.RowIndex].Cells["Photo_azimuth"].Value.ToString();
                //txtDupOf.Text = dgData.Rows[e.RowIndex].Cells["DupOf"].Value.ToString();
                //txtLenght.Text = dgData.Rows[e.RowIndex].Cells["length"].Value.ToString();
                //txtHigh.Text = dgData.Rows[e.RowIndex].Cells["High"].Value.ToString();
                //txtThickness.Text = dgData.Rows[e.RowIndex].Cells["Thickness"].Value.ToString();
                //txtObservations.Text = dgData.Rows[e.RowIndex].Cells["Obsevations"].Value.ToString();
                //txtObservSedimentary.Text = dgData.Rows[e.RowIndex].Cells["LRocksObservation"].Value.ToString();
                //txtMatrixObserv.Text = dgData.Rows[e.RowIndex].Cells["LMatrixObsevations"].Value.ToString();
                //txtPhenoObserv.Text = dgData.Rows[e.RowIndex].Cells["LPhenoCObsevations"].Value.ToString();
                //cmbContactType.SelectedValue = dgData.Rows[e.RowIndex].Cells["VContactType"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["VContactType"].Value.ToString();
                //cmbVeinName.SelectedValue = dgData.Rows[e.RowIndex].Cells["VVeinName"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["VVeinName"].Value.ToString();
                //cmbHostRock.SelectedValue = dgData.Rows[e.RowIndex].Cells["VHostRock"].Value.ToString() == "" ? "-1" : dgData.Rows[e.RowIndex].Cells["VHostRock"].Value.ToString(); 
                //txtVeinObserv.Text = dgData.Rows[e.RowIndex].Cells["VObsevations"].Value.ToString();
                ////txtMine.Text = dgData.Rows[e.RowIndex].Cells["Mine"].Value.ToString();
                ////txtMineEntrance.Text = dgData.Rows[e.RowIndex].Cells["MineEntrance"].Value.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadData_CHLith()
        {
            try
            {
                dgData.DataSource = LoadDataCHAll("1");
                dgData.Columns["SKCHSamples"].Visible = false;
                dgLithology.DataSource = LoadDataCHAll("1");
                dgLithology.Columns["SKCHSamples"].Visible = false;
                LoadDataMinLith("1");
                LoadDataMinLith("2");

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void txtSampleHead_Leave(object sender, EventArgs e)
        {
            try
            {
                //LoadData_CHLith();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgData.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Project " + dgData.Rows[e.RowIndex].Cells["Project"].Value.ToString()
                   + " Geologist " + dgData.Rows[e.RowIndex].Cells["Geologist"].Value.ToString()
                   , "Geochemistry Rocks", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oCHSamp.iSKCHSamples = int.Parse(dgData.Rows[e.RowIndex].Cells["SKCHSamples"].Value.ToString());
                    string sRespDel = oCHSamp.CH_Samples_Delete();
                    if (sRespDel == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Geochemistry", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dgData.DataSource = LoadDataCHAll("1");
                        dgData.Columns["SKCHSamples"].Visible = false;

                        dgLithology.DataSource = LoadDataCHAll("1");
                        dgLithology.Columns["SKCHSamples"].Visible = false;

                        //CleanControls();
                    }
                    else
                    {
                        MessageBox.Show("Error: " + sRespDel, "Geochemistry", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                CleanControlsHeader();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void cmbSample_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{
                
            //    LoadData_CHLith();
            //    LoadDgSurveys();
            //    LoadDataAlterations("2");
            //    LoadDataMineralizations("2");
            //    LoadDataOxides("2");
            //    LoadDataStructures("2");

            //    oCHSamp.sSample = cmbSample.SelectedValue.ToString();
            //    oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
            //    DataTable dtSamp = LoadDataCHSurv(cmbSample.SelectedValue.ToString());
            //    //oCHSamp.getCHSamplesBySample();

            //    //
            //    if (dtSamp != null)
            //    {
            //        if (dtSamp.Rows.Count > 0)
            //        {
            //            txtToSur.Text = dtSamp.Rows[0]["To"].ToString();
            //        }
            //    }


            //    if (cmbSample.SelectedValue.ToString() != "Select an option..")
            //    {
            //        DataTable dgSamp = (DataTable)dgData.DataSource;
            //        DataRow[] myRow = dgSamp.Select(@"Sample = '" + cmbSample.SelectedValue.ToString() + "'");
            //        int rowindex = dgSamp.Rows.IndexOf(myRow[0]);
            //        dgData.Rows[rowindex].Selected = true;
            //        dgData.CurrentCell = dgData.Rows[rowindex].Cells[1];

            //        dgD_CellClick(rowindex, dgData, "Cmb");
            //    }
               


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            
        }

        private void LoadCmbSurvey()
        {
            try
            {
                DataTable dtConections = new DataTable();
                dtConections.Columns.Add("Key", typeof(String));
                dtConections.Columns.Add("Value", typeof(String));


                for (int i = 0; i < conf.AppSettings.Settings.Count; i++)
                {
                    if (conf.AppSettings.Settings.AllKeys[i].ToString().Contains("Inst"))
                    {

                        DataRow drConect = dtConections.NewRow();
                        //drConect["Con"] = ;
                        drConect["Key"] = conf.AppSettings.Settings.AllKeys[i].ToString();
                        drConect["Value"] =
                            conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value.ToString();
                        dtConections.Rows.Add(drConect);

                        //MessageBox.Show(conf.AppSettings.Settings.AllKeys[i].ToString());
                        cmbInstSur.Items.Add(conf.AppSettings.Settings.AllKeys[i].ToString());
                        string s = conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value;
                    }

                }

                cmbInstSur.DisplayMember = "Value";
                cmbInstSur.ValueMember = "Key";
                cmbInstSur.DataSource = dtConections;

                dtConections = new DataTable();

                dtConections.Columns.Add("Key", typeof(String));
                dtConections.Columns.Add("Value", typeof(String));


                for (int i = 0; i < conf.AppSettings.Settings.Count; i++)
                {
                    if (conf.AppSettings.Settings.AllKeys[i].ToString().Contains("MineL"))
                    {

                        DataRow drConect = dtConections.NewRow();
                        //drConect["Con"] = ;
                        drConect["Key"] = conf.AppSettings.Settings.AllKeys[i].ToString();
                        drConect["Value"] =
                            conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value.ToString();
                        dtConections.Rows.Add(drConect);

                        //MessageBox.Show(conf.AppSettings.Settings.AllKeys[i].ToString());
                        cmbMineLocation.Items.Add(conf.AppSettings.Settings.AllKeys[i].ToString());
                        string s = conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value;
                    }

                }

                cmbMineLocation.DisplayMember = "Value";
                cmbMineLocation.ValueMember = "Key";
                cmbMineLocation.DataSource = dtConections;

                dtConections = new DataTable();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnCancelSur_Click(object sender, EventArgs e)
        {
            try
            {
                sEditSur = "0";
                CleanControlsSurvey();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void txtAzimuthSur_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtDipSur_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtP1E_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtP1N_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtP1Z_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtP2E_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtP2N_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtP2Z_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtCoordE_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtCoordN_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtCoordElevation_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtGPSEPE_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtLenght_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtHigh_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void btnAddMinMat_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbSample.Text != "")
                {
                    if (sEditMinLithMat == "0")
                    {
                        oMinLith.iSKMinLith = 0;
                        oMinLith.sOpcion = "1";
                    }
                    else if (sEditMinLithMat == "1")
                    {
                        oMinLith.sOpcion = "2";
                    }

                    oMinLith.sChid = cmbChannelId.SelectedValue.ToString();
                    oMinLith.sMineral = cmbMineralMt.SelectedValue.ToString();
                    oMinLith.sSample = cmbSample.Text.ToString();
                    oMinLith.sType = "Mat";
                    string sResp = oMinLith.CHMinLith_Add();
                    if (sResp == "OK")
                    {
                        cmbMineralMt.SelectedValue = "-1";
                        LoadDataMinLith("1");
                        sEditMinLithMat = "0";
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadDataMinLith(string _sOpcion)
        {
            try
            {
                if (_sOpcion == "1")
                {
                    oMinLith.sOpcion = _sOpcion;
                    oMinLith.sSample = cmbSample.SelectedValue.ToString();
                    dgLithMatrix.DataSource = oMinLith.getGCSamplesRockLithList();
                    dgLithMatrix.Columns["SKMinLith"].Visible = false;

                }
                else if (_sOpcion == "2")
                {
                    oMinLith.sOpcion = _sOpcion;
                    oMinLith.sSample = cmbSample.SelectedValue.ToString();
                    dgLithPheno.DataSource = oMinLith.getGCSamplesRockLithList();
                    dgLithPheno.Columns["SKMinLith"].Visible = false;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddMinPhe_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbSample.Text != "")
                {
                    if (sEditMinLithPhe == "0")
                    {
                        oMinLith.iSKMinLith = 0;
                        oMinLith.sOpcion = "1";
                    }
                    else if (sEditMinLithPhe == "1")
                    {
                        oMinLith.sOpcion = "2";
                    }

                    oMinLith.sChid = cmbChannelId.SelectedValue.ToString();
                    oMinLith.sMineral = cmbMineralPh.SelectedValue.ToString();
                    oMinLith.sSample = cmbSample.Text.ToString();
                    oMinLith.sType = "Phe";
                    string sResp = oMinLith.CHMinLith_Add();
                    if (sResp == "OK")
                    {
                        cmbMineralPh.SelectedValue = "-1";
                        LoadDataMinLith("2");
                        sEditMinLithPhe = "0";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void dgLithMatrix_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                oMinLith.iSKMinLith = int.Parse(dgLithMatrix.Rows[e.RowIndex].Cells["SKMinLith"].Value.ToString());
                sEditMinLithMat = "1";

                if (dgLithMatrix.Rows[e.RowIndex].Cells["Mineral"].Value.ToString() == "")
                    cmbMineralMt.SelectedValue = "-1";
                else cmbMineralMt.SelectedValue = dgLithMatrix.Rows[e.RowIndex].Cells["Mineral"].Value.ToString();

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgLithPheno_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                oMinLith.iSKMinLith = int.Parse(dgLithPheno.Rows[e.RowIndex].Cells["SKMinLith"].Value.ToString());
                sEditMinLithPhe = "1";

                if (dgLithPheno.Rows[e.RowIndex].Cells["Mineral"].Value.ToString() == "")
                    cmbMineralPh.SelectedValue = "-1";
                else cmbMineralPh.SelectedValue = dgLithPheno.Rows[e.RowIndex].Cells["Mineral"].Value.ToString();

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelMinMat_Click(object sender, EventArgs e)
        {
            try
            {
                sEditMinLithMat = "0";
                cmbMineralMt.SelectedValue = "-1";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 
        }

        private void btnAddLithology_Click(object sender, EventArgs e)
        {
            try
            {
                 AddHeaderLithology();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelMinPhe_Click(object sender, EventArgs e)
        {
            try
            {
                sEditMinLithPhe = "0";
                cmbMineralPh.SelectedValue = "-1";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 
        }

        private void dgLithMatrix_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgLithMatrix.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Mineral " + dgLithMatrix.Rows[e.RowIndex].Cells["Mineral"].Value.ToString()
                   , "Lithology Matrix", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oMinLith.iSKMinLith = int.Parse(dgLithMatrix.Rows[e.RowIndex].Cells["SKMinLith"].Value.ToString());
                    oMinLith.CHMinLithLith_Delete();
                    LoadDataMinLith("1");
                    sEditMinLithMat = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgLithPheno_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgLithPheno.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Mineral " + dgLithPheno.Rows[e.RowIndex].Cells["Mineral"].Value.ToString()
                   , "Lithology Matrix", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oMinLith.iSKMinLith = int.Parse(dgLithPheno.Rows[e.RowIndex].Cells["SKMinLith"].Value.ToString());
                    oMinLith.CHMinLithLith_Delete();
                    LoadDataMinLith("2");
                    sEditMinLithPhe = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddAlt_Click(object sender, EventArgs e)
        {
            try
            {
                if (sEditAlt == "0")
                {
                    oAlt.sOpcion = "1";
                    oAlt.iSKAlteration = 0;
                }
                else if (sEditAlt == "1")
                {
                    oAlt.sOpcion = "2";
                }
                
                oAlt.sChid = cmbChannelId.SelectedValue.ToString();
                if (cmbSample.SelectedValue.ToString() == "-1" ||
                    cmbSample.SelectedValue.ToString() == "")
                {
                    MessageBox.Show("Empty Sample");
                    return;
                }
                oAlt.sSample = cmbSample.SelectedValue.ToString();

                if (cmbTypeAlt.SelectedValue.ToString() == "-1" ||
                    cmbTypeAlt.SelectedValue.ToString() == "")
                    oAlt.sAltType = null;
                else
                    oAlt.sAltType = cmbTypeAlt.SelectedValue.ToString();

                if (cmbIntAlt.SelectedValue.ToString() == "-1" ||
                    cmbIntAlt.SelectedValue.ToString() == "")
                    oAlt.sAltInt = null;
                else
                    oAlt.sAltInt = cmbIntAlt.SelectedValue.ToString();

                if (cmbStyleAlt1.SelectedValue.ToString() == "-1" ||
                    cmbStyleAlt1.SelectedValue.ToString() == "")
                    oAlt.sAltStyle = null;
                else
                    oAlt.sAltStyle = cmbStyleAlt1.SelectedValue.ToString();

                if (cmbMin1Alt.SelectedValue.ToString() == "-1" ||
                    cmbMin1Alt.SelectedValue.ToString() == "")
                    oAlt.sAltMin = null;
                else
                    oAlt.sAltMin = cmbMin1Alt.SelectedValue.ToString();

                if (cmbMin2Alt1.SelectedValue.ToString() == "-1" ||
                    cmbMin2Alt1.SelectedValue.ToString() == "")
                    oAlt.sAltMin2 = null;
                else
                    oAlt.sAltMin2 = cmbMin2Alt1.SelectedValue.ToString();

                if (cmbMin3Alt1.SelectedValue.ToString() == "-1" ||
                    cmbMin3Alt1.SelectedValue.ToString() == "")
                    oAlt.sAltMin3 = null;
                else
                    oAlt.sAltMin3 = cmbMin3Alt1.SelectedValue.ToString();

                if (txtObservAlt.Text.ToString() == "")
                    oAlt.sObservations = null;
                else
                    oAlt.sObservations = txtObservAlt.Text.ToString();

                string sResp = oAlt.CHAlterations_Add();
                if (sResp == "OK")
                {
                    LoadDataAlterations("2");
                    sEditAlt = "0";
                    CleanControlsAlt();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CleanControlsAlt()
        {
            try
            {
                cmbTypeAlt.SelectedValue = "-1";
                cmbIntAlt.SelectedValue = "-1";
                cmbStyleAlt1.SelectedValue = "-1";
                cmbMin1Alt.SelectedValue = "-1";
                cmbMin2Alt1.SelectedValue = "-1";
                cmbMin3Alt1.SelectedValue = "-1";
                txtObservAlt.Text = "";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void LoadDataAlterations(string _sOpcion)
        {
            try
            {

                oAlt.sOpcion = _sOpcion;
                if (cmbSample.SelectedValue == null)
                {
                    return;
                }
                if (cmbSample.SelectedValue.ToString() == "" ||
                    cmbSample.SelectedValue.ToString() == "-1")
                {
                    return;
                }
                oAlt.sSample = cmbSample.SelectedValue.ToString();
                oAlt.sChid = cmbChannelId.SelectedValue.ToString();
                dgAlterations.DataSource = oAlt.getCHAlterationsList();
                dgAlterations.Columns["SKAlteration"].Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgAlterations_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditAlt = "1";
                oAlt.iSKAlteration = int.Parse(dgAlterations.Rows[e.RowIndex].Cells["SKAlteration"].Value.ToString());

                if (dgAlterations.Rows[e.RowIndex].Cells["ALTType"].Value.ToString() == "")
                    cmbTypeAlt.SelectedValue = "-1";
                else cmbTypeAlt.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["ALTType"].Value.ToString();

                if (dgAlterations.Rows[e.RowIndex].Cells["ALTInt"].Value.ToString() == "")
                    cmbIntAlt.SelectedValue = "-1";
                else cmbIntAlt.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["ALTInt"].Value.ToString();

                if (dgAlterations.Rows[e.RowIndex].Cells["ALTStyle"].Value.ToString() == "")
                    cmbStyleAlt1.SelectedValue = "-1";
                else cmbStyleAlt1.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["ALTStyle"].Value.ToString();

                if (dgAlterations.Rows[e.RowIndex].Cells["ALTMin"].Value.ToString() == "")
                    cmbMin1Alt.SelectedValue = "-1";
                else cmbMin1Alt.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["ALTMin"].Value.ToString();

                if (dgAlterations.Rows[e.RowIndex].Cells["ALTMin2"].Value.ToString() == "")
                    cmbMin2Alt1.SelectedValue = "-1";
                else cmbMin2Alt1.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["ALTMin2"].Value.ToString();

                if (dgAlterations.Rows[e.RowIndex].Cells["ALTMin3"].Value.ToString() == "")
                    cmbMin3Alt1.SelectedValue = "-1";
                else cmbMin3Alt1.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["ALTMin3"].Value.ToString();

                txtObservAlt.Text = dgAlterations.Rows[e.RowIndex].Cells["Obsevations"].Value.ToString();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgAlterations_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgAlterations.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Type " + dgAlterations.Rows[e.RowIndex].Cells["ALTType"].Value.ToString()
                   + " Intensity " + dgAlterations.Rows[e.RowIndex].Cells["ALTInt"].Value.ToString()
                   , "Alterations ", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oAlt.iSKAlteration = int.Parse(dgAlterations.Rows[e.RowIndex].Cells["SKAlteration"].Value.ToString());
                    oAlt.CHAlterations_Delete();
                    LoadDataAlterations("2");
                    sEditAlt = "0";
                    CleanControlsAlt();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddMin_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMinPerc.Text != "")
                {
                    if (double.Parse(txtMinPerc.Text) > 100)
                    {
                        MessageBox.Show("Percentage isn´t more than 100");
                        txtMinPerc.Focus();
                        return;
                    }
                }


                if (sEditMin == "0")
                {
                    oMin.sOpcion = "1";
                    oMin.iSKMineralizations = 0;
                }
                else
                {
                    oMin.sOpcion = "2";
                }

                oMin.sChid = cmbChannelId.SelectedValue.ToString();
                oMin.sSample = cmbSample.SelectedValue.ToString();

                if (cmbMineralmin.SelectedValue.ToString() == "-1" ||
                    cmbMineralmin.SelectedValue.ToString() == "")
                    oMin.sMineral = null;
                else
                    oMin.sMineral = cmbMineralmin.SelectedValue.ToString();

                if (cmbStyleM.SelectedValue.ToString() == "-1" ||
                    cmbStyleM.SelectedValue.ToString() == "")
                    oMin.sMinStyle = null;
                else
                    oMin.sMinStyle = cmbStyleM.SelectedValue.ToString();

                if (txtMinPerc.Text.ToString() == "")
                    oMin.dMinPerc = null;
                else
                    oMin.dMinPerc = double.Parse(txtMinPerc.Text.ToString());

                if (txtObservMin.Text.ToString() == "")
                    oMin.sObservations = null;
                else
                    oMin.sObservations = txtObservMin.Text.ToString();

                string sResp = oMin.CHMineralizations_Add();
                if (sResp == "OK")
                {
                    CleanControlsMin();
                    sEditMin = "0";
                    LoadDataMineralizations("2");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadDataMineralizations(string _sOpcion)
        {
            try
            {

                oMin.sOpcion = _sOpcion;
                if (cmbSample.SelectedValue == null)
                {
                    return;
                }
                if (cmbSample.SelectedValue.ToString() == "" ||
                    cmbSample.SelectedValue.ToString() == "-1")
                {
                    return;
                }
                oMin.sSample = cmbSample.SelectedValue.ToString();
                oMin.sChid = cmbChannelId.SelectedValue.ToString();
                dgMineralizations.DataSource = oMin.getCHMineralizationsList();
                dgMineralizations.Columns["SKMineralizations"].Visible = false;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void CleanControlsMin()
        {
            try
            {
                cmbMineralmin.SelectedValue = "-1";
                cmbStyleM.SelectedValue = "-1";
                txtMinPerc.Text = "";
                txtObservMin.Text = "";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgSamplesRockMin_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditMin = "1";
                oMin.iSKMineralizations = int.Parse(dgMineralizations.Rows[e.RowIndex].Cells["SKMineralizations"].Value.ToString());

                if (dgMineralizations.Rows[e.RowIndex].Cells["MZMin"].Value.ToString() == "")
                    cmbMineralmin.SelectedValue = "-1";
                else cmbMineralmin.SelectedValue = dgMineralizations.Rows[e.RowIndex].Cells["MZMin"].Value.ToString();

                if (dgMineralizations.Rows[e.RowIndex].Cells["MZStyle"].Value.ToString() == "")
                    cmbStyleM.SelectedValue = "-1";
                else cmbStyleM.SelectedValue = dgMineralizations.Rows[e.RowIndex].Cells["MZStyle"].Value.ToString();

                if (dgMineralizations.Rows[e.RowIndex].Cells["MZPerc"].Value.ToString() == "")
                    txtMinPerc.Text = "";
                else txtMinPerc.Text = dgMineralizations.Rows[e.RowIndex].Cells["MZPerc"].Value.ToString();

                txtObservMin.Text = dgMineralizations.Rows[e.RowIndex].Cells["Obsevations"].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgSamplesRockMin_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgMineralizations.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Mineral " + dgMineralizations.Rows[e.RowIndex].Cells["MZMin"].Value.ToString()
                   + " Style " + dgMineralizations.Rows[e.RowIndex].Cells["MZStyle"].Value.ToString()
                   , "Alterations ", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oMin.iSKMineralizations = int.Parse(dgMineralizations.Rows[e.RowIndex].Cells["SKMineralizations"].Value.ToString());
                    oMin.CHMineralizations_Delete();
                    LoadDataMineralizations("1");
                    sEditMin = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddOxides_Click(object sender, EventArgs e)
        {
            try
            {
                if (sEditOxid == "0")
                {
                    oOxid.sOpcion = "1";
                    oOxid.iSKOxides = 0;
                }
                else
                {
                    oOxid.sOpcion = "2";
                }

                oOxid.sChid = cmbChannelId.SelectedValue.ToString();
                oOxid.sSample = cmbSample.SelectedValue.ToString();

                if (cmbStyleGoe.SelectedValue.ToString() == "-1" ||
                    cmbStyleGoe.SelectedValue.ToString() == "")
                    oOxid.sGoeStyle = null;
                else
                    oOxid.sGoeStyle = cmbStyleGoe.SelectedValue.ToString();

                if (cmbPercGoe.SelectedValue.ToString() == "-1" ||
                   cmbPercGoe.SelectedValue.ToString() == "")
                    oOxid.sGoePerc = null;
                else
                    oOxid.sGoePerc = cmbPercGoe.SelectedValue.ToString();

                if (cmbStyleHem.SelectedValue.ToString() == "-1" ||
                    cmbStyleHem.SelectedValue.ToString() == "")
                    oOxid.sHemStyle = null;
                else
                    oOxid.sHemStyle = cmbStyleHem.SelectedValue.ToString();

                if (cmbPercHem.SelectedValue.ToString() == "-1" ||
                   cmbPercHem.SelectedValue.ToString() == "")
                    oOxid.sHemPerc = null;
                else
                    oOxid.sHemPerc = cmbPercHem.SelectedValue.ToString();


                if (cmbStyleJar.SelectedValue.ToString() == "-1" ||
                   cmbStyleJar.SelectedValue.ToString() == "")
                    oOxid.sJarStyle = null;
                else
                    oOxid.sJarStyle = cmbStyleJar.SelectedValue.ToString();

                if (cmbPercJar.SelectedValue.ToString() == "-1" ||
                   cmbPercJar.SelectedValue.ToString() == "")
                    oOxid.sJarPerc = null;
                else
                    oOxid.sJarPerc = cmbPercJar.SelectedValue.ToString();

                if (cmbStyleLim.SelectedValue.ToString() == "-1" ||
                   cmbStyleLim.SelectedValue.ToString() == "")
                    oOxid.sLimStyle = null;
                else
                    oOxid.sLimStyle = cmbStyleLim.SelectedValue.ToString();

                if (cmbPercLim.SelectedValue.ToString() == "-1" ||
                   cmbPercLim.SelectedValue.ToString() == "")
                    oOxid.sLimPerc = null;
                else
                    oOxid.sLimPerc = cmbPercLim.SelectedValue.ToString();

                if (txtObservOxides.Text.ToString() == "")
                    oOxid.sObservations = null;
                else
                    oOxid.sObservations = txtObservOxides.Text.ToString();

                string sResp = oOxid.CHOxides_Add();
                if (sResp == "OK")
                {
                    sEditOxid = "0";
                    CleanControlsOxides();
                    LoadDataOxides("2");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadDataOxides(string _sOpcion)
        {
            try
            {

                oOxid.sOpcion = _sOpcion;
                if (cmbSample.SelectedValue == null)
                {
                    return;
                }
                if (cmbSample.SelectedValue.ToString() == "" ||
                    cmbSample.SelectedValue.ToString() == "-1")
                {
                    return;
                }
                oOxid.sSample = cmbSample.SelectedValue.ToString();
                oOxid.sChid = cmbChannelId.SelectedValue.ToString();
                dgOxides.DataSource = oOxid.getCHOxidesList();
                dgOxides.Columns["SKOxides"].Visible = false;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void CleanControlsOxides()
        {
            try
            {

                cmbStyleGoe.SelectedValue = "-1";
                cmbPercGoe.SelectedValue = "-1";
                cmbStyleHem.SelectedValue = "-1";
                cmbPercHem.SelectedValue = "-1";
                cmbStyleJar.SelectedValue = "-1";
                cmbPercJar.SelectedValue = "-1";
                cmbStyleLim.SelectedValue = "-1";
                cmbPercLim.SelectedValue = "-1";
                txtObservOxides.Text = "";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgOxides_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditOxid = "1";
                oOxid.iSKOxides = int.Parse(dgOxides.Rows[e.RowIndex].Cells["SKOxides"].Value.ToString());

                if (dgOxides.Rows[e.RowIndex].Cells["GoeStyle"].Value.ToString() == "")
                    cmbStyleGoe.SelectedValue = "-1";
                else cmbStyleGoe.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["GoeStyle"].Value.ToString();

                if (dgOxides.Rows[e.RowIndex].Cells["GoePerc"].Value.ToString() == "")
                    cmbPercGoe.SelectedValue = "-1";
                else cmbPercGoe.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["GoePerc"].Value.ToString();

                if (dgOxides.Rows[e.RowIndex].Cells["HemStyle"].Value.ToString() == "")
                    cmbStyleHem.SelectedValue = "-1";
                else cmbStyleHem.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["HemStyle"].Value.ToString();

                if (dgOxides.Rows[e.RowIndex].Cells["HemPerc"].Value.ToString() == "")
                    cmbPercHem.SelectedValue = "-1";
                else cmbPercHem.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["HemPerc"].Value.ToString();

                if (dgOxides.Rows[e.RowIndex].Cells["JarStyle"].Value.ToString() == "")
                    cmbStyleJar.SelectedValue = "-1";
                else cmbStyleJar.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["JarStyle"].Value.ToString();

                if (dgOxides.Rows[e.RowIndex].Cells["JarPerc"].Value.ToString() == "")
                    cmbPercJar.SelectedValue = "-1";
                else cmbPercJar.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["JarPerc"].Value.ToString();

                if (dgOxides.Rows[e.RowIndex].Cells["LimStyle"].Value.ToString() == "")
                    cmbStyleLim.SelectedValue = "-1";
                else cmbStyleLim.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["LimStyle"].Value.ToString();

                if (dgOxides.Rows[e.RowIndex].Cells["LimPerc"].Value.ToString() == "")
                    cmbPercLim.SelectedValue = "-1";
                else cmbPercLim.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["LimPerc"].Value.ToString();

                txtObservOxides.Text = dgOxides.Rows[e.RowIndex].Cells["Observations"].Value.ToString();
                cmbSample.SelectedValue = dgOxides.Rows[e.RowIndex].Cells["Sample"].Value.ToString();
                txtChId.Text = dgOxides.Rows[e.RowIndex].Cells["Chid"].Value.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgOxides_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgOxides.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Goethite Style " + dgOxides.Rows[e.RowIndex].Cells["GoeStyle"].Value.ToString()
                   + " Hematite Style " + dgOxides.Rows[e.RowIndex].Cells["HemStyle"].Value.ToString()
                   , " Oxides ", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oOxid.iSKOxides = int.Parse(dgOxides.Rows[e.RowIndex].Cells["SKOxides"].Value.ToString());
                    oOxid.CHOxides_Delete();
                    LoadDataOxides("2");
                    sEditOxid = "0";
                    CleanControlsOxides();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddStr_Click(object sender, EventArgs e)
        {
            try
            {
                if (sEditStr == "0")
                {
                    oStr.sOpcion = "1";
                    oStr.iSKStructures = 0;
                }
                else
                {
                    oStr.sOpcion = "2";
                }

                oStr.sChid = cmbChannelId.SelectedValue.ToString();
                oStr.sSample = cmbSample.SelectedValue.ToString();

                if (cmbStructureTypeSt.SelectedValue.ToString() == "-1" ||
                    cmbStructureTypeSt.SelectedValue.ToString() == "")
                    oStr.sType = null;
                else
                    oStr.sType = cmbStructureTypeSt.SelectedValue.ToString();

                if (txtDipStr.Text.ToString() == "")
                    oStr.dDip = null;
                else
                    oStr.dDip = double.Parse(txtDipStr.Text.ToString());

                if (txtDipAzStr.Text.ToString() == "")
                    oStr.sDipAz = null;
                else
                    oStr.sDipAz = txtDipAzStr.Text.ToString();

                if (txtAppThickSt.Text.ToString() == "")
                    oStr.dAThick = null;
                else
                    oStr.dAThick = double.Parse(txtAppThickSt.Text.ToString());

                if (txtRThickStr.Text.ToString() == "")
                    oStr.dRThick = null;
                else
                    oStr.dRThick = double.Parse(txtRThickStr.Text.ToString());

                if (cmbFillSt.SelectedValue.ToString() == "-1" ||
                    cmbFillSt.SelectedValue.ToString() == "")
                    oStr.sFill = null;
                else
                    oStr.sFill = cmbFillSt.SelectedValue.ToString();


                if (cmbFillSt2.SelectedValue.ToString() == "-1" ||
                    cmbFillSt2.SelectedValue.ToString() == "")
                    oStr.sFill2 = null;
                else
                    oStr.sFill2 = cmbFillSt2.SelectedValue.ToString();

                if (cmbFillSt3.SelectedValue.ToString() == "-1" ||
                    cmbFillSt3.SelectedValue.ToString() == "")
                    oStr.sFill3 = null;
                else
                    oStr.sFill3 = cmbFillSt3.SelectedValue.ToString();

                if (txtNumberSt.Text.ToString() == "")
                    oStr.dNumber = null;
                else
                    oStr.dNumber = double.Parse(txtNumberSt.Text.ToString());

                if (txtDensityStr.Text.ToString() == "")
                    oStr.dDensity = null;
                else
                    oStr.dDensity = double.Parse(txtDensityStr.Text.ToString());

                if (txtObservStr.Text.ToString() == "")
                    oStr.sObservations = null;
                else
                    oStr.sObservations = txtObservStr.Text.ToString();

                string sResp = oStr.CHStructures_Add();
                if (sResp == "OK")
                {
                    sEditStr = "0";
                    CleanControlsStr();
                    LoadDataStructures("2");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CleanControlsStr()
        {
            try
            {
                cmbStructureTypeSt.SelectedValue = "-1";
                txtDipStr.Text = "";
                txtDipAzStr.Text = "";
                txtAppThickSt.Text = "";
                txtRThickStr.Text = "";
                cmbFillSt.SelectedValue = "-1";
                cmbFillSt2.SelectedValue = "-1";
                cmbFillSt3.SelectedValue = "-1";
                txtNumberSt.Text = "";
                txtDensityStr.Text = "";
                txtObservStr.Text = "";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void LoadDataStructures(string _sOpcion)
        {
            try
            {
                oStr.sOpcion = _sOpcion;
                if (cmbSample.SelectedValue == null)
                {
                    return;
                }
                if (cmbSample.SelectedValue.ToString() == "" ||
                    cmbSample.SelectedValue.ToString() == "-1")
                {
                    return;
                }
                oStr.sSample = cmbSample.SelectedValue.ToString();
                oStr.sChid = cmbChannelId.SelectedValue.ToString();
                dgStructures.DataSource = oStr.getCHStructuresList();
                dgStructures.Columns["SKStructures"].Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgStructures_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditStr = "1";
                oStr.iSKStructures = int.Parse(dgStructures.Rows[e.RowIndex].Cells["SKStructures"].Value.ToString());

                if (dgStructures.Rows[e.RowIndex].Cells["StrType"].Value.ToString() == "")
                    cmbStructureTypeSt.SelectedValue = "-1";
                else cmbStructureTypeSt.SelectedValue = dgStructures.Rows[e.RowIndex].Cells["StrType"].Value.ToString();

                txtDipStr.Text = dgStructures.Rows[e.RowIndex].Cells["StrDip"].Value.ToString();
                txtDipAzStr.Text = dgStructures.Rows[e.RowIndex].Cells["StrDipAz"].Value.ToString();
                txtAppThickSt.Text = dgStructures.Rows[e.RowIndex].Cells["StrAThick"].Value.ToString();
                txtRThickStr.Text = dgStructures.Rows[e.RowIndex].Cells["StrRThick"].Value.ToString();

                if (dgStructures.Rows[e.RowIndex].Cells["StrFill"].Value.ToString() == "")
                    cmbFillSt.SelectedValue = "-1";
                else cmbFillSt.SelectedValue = dgStructures.Rows[e.RowIndex].Cells["StrFill"].Value.ToString();

                if (dgStructures.Rows[e.RowIndex].Cells["StrFill2"].Value.ToString() == "")
                    cmbFillSt2.SelectedValue = "-1";
                else cmbFillSt2.SelectedValue = dgStructures.Rows[e.RowIndex].Cells["StrFill2"].Value.ToString();

                if (dgStructures.Rows[e.RowIndex].Cells["StrFill3"].Value.ToString() == "")
                    cmbFillSt3.SelectedValue = "-1";
                else cmbFillSt3.SelectedValue = dgStructures.Rows[e.RowIndex].Cells["StrFill3"].Value.ToString();

                txtNumberSt.Text = dgStructures.Rows[e.RowIndex].Cells["StrNumber"].Value.ToString();
                txtDensityStr.Text = dgStructures.Rows[e.RowIndex].Cells["StrDensity"].Value.ToString();
                txtObservStr.Text = dgStructures.Rows[e.RowIndex].Cells["Obsevations"].Value.ToString();
                txtChId.Text = dgStructures.Rows[e.RowIndex].Cells["Chid"].Value.ToString();
                cmbSample.SelectedValue = dgStructures.Rows[e.RowIndex].Cells["Sample"].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtDipStr_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtAppThickSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtRThickStr_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtNumberSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtDensityStr_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtEastCh_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtNorthCh_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtElevationCh_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtLenghtCh_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtToSur_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void dgStructures_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgStructures.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                      + " Dip " + dgStructures.Rows[e.RowIndex].Cells["StrDip"].Value.ToString()
                      + " Fill " + dgStructures.Rows[e.RowIndex].Cells["StrFill"].Value.ToString()
                      , "Structures", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                  MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oStr.iSKStructures = int.Parse(dgStructures.Rows[e.RowIndex].Cells["SKStructures"].Value.ToString());
                    oStr.CHstructures_Delete();
                    LoadDataStructures("2");
                    sEditStr = "0";
                    CleanControlsStr();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TbRocks_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == (char)Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void cmbChannelId_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    LoadDgSurveys();
            //    LoadCmbSamples();
            //    dgData.DataSource = LoadDataCHAll("1");
            //}
            //catch (Exception ex)
            //{
            //    throw new Exception(ex.Message);
            //}
        }

        private void txtTotalSamples_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void LoadCmbSamplingFilter()
        {
            try
            {
                DataTable dCh = (DataTable)dgDataCh.DataSource;

                DataSet dtSampleT = new DataSet();
                dtSampleT = oRf.getRfTypeSampleDataSet();

                IEnumerable<DataRow> query =
                    from dC in dCh.AsEnumerable()
                    where dC.Field<int?>("MineID") == 133
                        && dC.Field<String>("Chid") == cmbChannelId.SelectedValue.ToString()
                    select dC;

                DataTable dtSam = new DataTable();
                if (query.Count() > 0)
                {
                    //MessageBox.Show("Surface");

                    IEnumerable<DataRow> queryS =
                       from dSampleT in dtSampleT.Tables[2].AsEnumerable()
                       where dSampleT.Field<String>("Code").Substring(0,1) == "O"
                       select dSampleT;
                    dtSam = queryS.CopyToDataTable<DataRow>();

                   
                }
                else
                {
                    //MessageBox.Show("<> Surface");

                    IEnumerable<DataRow> queryS =
                       from dSampleT in dtSampleT.Tables[2].AsEnumerable()
                       where dSampleT.Field<String>("Code").Substring(0, 1) != "O"
                       select dSampleT;
                    dtSam = queryS.CopyToDataTable<DataRow>();

                   
                }

                DataRow drT2 = dtSam.NewRow();
                drT2[0] = "-1";
                drT2[1] = "Select an option..";
                dtSam.Rows.Add(drT2);
                cmbSamplingType.DisplayMember = "Comb";
                cmbSamplingType.ValueMember = "Code";
                cmbSamplingType.DataSource = dtSam;
                cmbSamplingType.SelectedValue = -1;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cmbChannelId_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                oCh.sChId = cmbChannelId.SelectedValue.ToString();
                oCh.sOpcion = "2";
                DataTable dtCh = oCh.getCH_Collars();
                lblSamplesTotalHeader.Text = " Of: " + dtCh.Rows[0]["SamplesTotal"].ToString();

                LoadDgSurveys();
                sSampleSelect = "";
                LoadCmbSamples();
                dgData.DataSource = LoadDataCHAll("1");
                LoadCmbSamplingFilter();

                CleanControlsAll();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnExporExcelAll_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbChannelId.SelectedValue == null)
                {
                     MessageBox.Show("Select Channel");
                    return;
                }

                if (cmbSample.SelectedValue == null)
                {
                    MessageBox.Show("Select Samples");
                    return;
                }

                if (cmbChannelId.SelectedValue.ToString() == "" ||
                    cmbChannelId.SelectedValue.ToString() == "-1")
                {
                    MessageBox.Show("Select Channel");
                    return;
                }

                if (cmbSample.SelectedValue.ToString() == "" ||
                    cmbSample.SelectedValue.ToString() == "-1")
                {
                    MessageBox.Show("Select Samples");
                    return;
                }

                sExport = "Geochemistry"; //Ejecuta los eventos bgw_DoWork, bgw_ProgressChanged y bgw_RunWorkerCompleted
                bgw.RunWorkerAsync();

            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                Thread.Sleep(100);

                DateTime start = DateTime.Now;
                e.Result = "";
                for (int i = 0; i < 100; i++)
                {
                    System.Threading.Thread.Sleep(50);

                    bgw.ReportProgress(i, DateTime.Now);


                    if (bgw.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }
                }

                TimeSpan duration = DateTime.Now - start;

                e.Result = "Duration: " + duration.TotalMilliseconds.ToString() + " ms.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                //SamplesValid();pbLogging

                pbGeochemistry.Visible = true;
                pbGeochemistry.Value = e.ProgressPercentage; //actualizamos la barra de progreso
                DateTime time = Convert.ToDateTime(e.UserState); //obtenemos información adicional si procede

                if (pbGeochemistry.Value > 98)
                {
                    pbGeochemistry.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                ExportExcelGeoch();
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
            
        }

        private void ExpGeochemistry()
        {
            try
            {

                DataTable dtSample = new DataTable();
                oCHSamp.sSample = cmbSample.SelectedValue.ToString();
                oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
                dtSample = oCHSamp.getCHSamplesBySampleReport();

                DataTable dtChannel = new DataTable();
                oCh.sChId = cmbChannelId.SelectedValue.ToString();
                oCh.sOpcion = "2";
                dtChannel = oCh.getCH_Collars();


                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                oXL = new Excel.Application();
                oXL.Visible = true;

                oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ExcelGeoch"].ToString(),
                    0, false, 5,
                Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                Type.Missing, false, false, false);

                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                oSheet.Cells[1, 24] = " Sample No: " + cmbSample.SelectedValue.ToString();  
                oSheet.Cells[2, 24] = "Channel ID: " + cmbChannelId.SelectedValue.ToString();

                if (dtSample.Rows.Count>0)
                {
                    oSheet.Cells[5, 9] = dtSample.Rows[0]["Nom_Target"].ToString();
                    //oSheet.Cells[5, 25] = dtChannel.Rows[0]["Location"].ToString();
                    oSheet.Cells[5, 43] = dtSample.Rows[0]["Project"].ToString();
                    oSheet.Cells[7, 9] = dtSample.Rows[0]["Nom_Geologist"].ToString();
                    oSheet.Cells[7, 18] = dtSample.Rows[0]["Helper"].ToString();
                    oSheet.Cells[7, 27] = dtSample.Rows[0]["Station"].ToString();

                    DateTime dDateS = DateTime.Parse(dtSample.Rows[0]["Date"].ToString());
                    string sDateS = dDateS.Day.ToString().PadLeft(2, '0') + "/" + dDateS.Month.ToString().PadLeft(2, '0')
                    + "/" + dDateS.Year.ToString().PadLeft(4, '0');
                    oSheet.Cells[7, 41] = sDateS.ToString();

                    //Encabezado
                    oSheet.Cells[9, 12] = dtSample.Rows[0]["E1"].ToString();
                    oSheet.Cells[9, 26] = dtSample.Rows[0]["N1"].ToString();
                    oSheet.Cells[9, 44] = dtSample.Rows[0]["Z1"].ToString();

                    //Survey
                    oSheet.Cells[11, 54] = dtSample.Rows[0]["E1"].ToString();
                    oSheet.Cells[13, 54] = dtSample.Rows[0]["N1"].ToString();
                    oSheet.Cells[15, 54] = dtSample.Rows[0]["Z1"].ToString();

                    oSheet.Cells[11, 67] = dtSample.Rows[0]["E2"].ToString();
                    oSheet.Cells[13, 67] = dtSample.Rows[0]["N2"].ToString();
                    oSheet.Cells[15, 67] = dtSample.Rows[0]["Z2"].ToString();


                    oSheet.Cells[11, 7] = dtSample.Rows[0]["CS"].ToString();
                    oSheet.Cells[11, 22] = dtSample.Rows[0]["GPSepe"].ToString();
                    oSheet.Cells[11, 29] = dtSample.Rows[0]["Photo"].ToString();
                    oSheet.Cells[11, 44] = dtSample.Rows[0]["Photo_azimuth"].ToString();

                    oSheet.Cells[15, 11] = dtSample.Rows[0]["SamplingType"].ToString();
                    oSheet.Cells[15, 22] = dtSample.Rows[0]["NotItSitu"].ToString();
                    oSheet.Cells[15, 30] = dtSample.Rows[0]["Porpuose"].ToString();
                    oSheet.Cells[15, 45] = dtSample.Rows[0]["Relative_Loc"].ToString();

                    oSheet.Cells[17, 7] = dtSample.Rows[0]["length"].ToString();
                    oSheet.Cells[17, 15] = dtSample.Rows[0]["High"].ToString();
                    oSheet.Cells[17, 26] = dtSample.Rows[0]["Thickness"].ToString();
                    oSheet.Cells[17, 32] = dtSample.Rows[0]["Obsevations"].ToString();

                    oSheet.Cells[23, 4] = dtSample.Rows[0]["LRock"].ToString();
                    oSheet.Cells[23, 9] = dtSample.Rows[0]["LTexture"].ToString();
                    oSheet.Cells[23, 13] = dtSample.Rows[0]["LGSize"].ToString();
                    oSheet.Cells[23, 18] = dtSample.Rows[0]["LWeathering"].ToString();

                    oSheet.Cells[27, 12] = dtSample.Rows[0]["LRocksSorting"].ToString();
                    oSheet.Cells[27, 17] = dtSample.Rows[0]["LRocksSphericity"].ToString();
                    oSheet.Cells[27, 20] = dtSample.Rows[0]["LRocksRounding"].ToString();
                    oSheet.Cells[30, 4] = dtSample.Rows[0]["LRocksObservation"].ToString();

                    oSheet.Cells[21, 27] = dtSample.Rows[0]["LMatrixPerc"].ToString();
                    oSheet.Cells[21, 32] = dtSample.Rows[0]["LMatrixGSize"].ToString();
                    oSheet.Cells[27, 26] = dtSample.Rows[0]["LMatrixObsevations"].ToString();

                    oSheet.Cells[21, 41] = dtSample.Rows[0]["LPhenoCPerc"].ToString();
                    oSheet.Cells[21, 46] = dtSample.Rows[0]["LPhenoCGSize"].ToString();
                    oSheet.Cells[27, 42] = dtSample.Rows[0]["LPhenoCObsevations"].ToString();

                    oSheet.Cells[7, 71] = dtSample.Rows[0]["From"].ToString();
                    oSheet.Cells[7, 75] = dtSample.Rows[0]["To"].ToString();

                    oSheet.Cells[51, 40] = dtSample.Rows[0]["VContactType"].ToString();
                    oSheet.Cells[53, 40] = dtSample.Rows[0]["VVeinName"].ToString();
                    oSheet.Cells[55, 40] = dtSample.Rows[0]["VHostRock"].ToString();
                    oSheet.Cells[57, 42] = dtSample.Rows[0]["VObsevations"].ToString();

                    oSheet.Cells[7, 55] = dtSample.Rows[0]["chId"].ToString();

                    oSheet.Cells[38, 62] = dtSample.Rows[0]["SampleType"].ToString();    

                }


                switch (dtChannel.Rows[0]["Instrument"].ToString())
                {
                    case "TS":
                        oSheet.Cells[9, 79] = "TS";
                        break;
                    case "CB":
                        oSheet.Cells[9, 81] = "GB";
                        break;
                    case "GPS":
                        oSheet.Cells[9, 83] = "GPS";
                        break;
                    default:
                        break;
                }


                DataTable dtMatrix = new DataTable();
                DataTable dtPheno = new DataTable();
                dtMatrix = getMinerals_Ph_Mx("1");
                dtPheno = getMinerals_Ph_Mx("2");

                #region Matrix Phenocryst
                if (dtMatrix.Rows.Count > 0)
                {
                    for (int i = 0; i < dtMatrix.Rows.Count; i++)
                    {
                        if (i < 4)
                        {
                            switch (i)
                            {
                                case 0:
                                    oSheet.Cells[23, 27] = dtMatrix.Rows[i]["Mineral"].ToString();
                                    break;
                                case 1:
                                    oSheet.Cells[23, 31] = dtMatrix.Rows[i]["Mineral"].ToString();
                                    break;
                                case 2:
                                    oSheet.Cells[25, 27] = dtMatrix.Rows[i]["Mineral"].ToString();
                                    break;
                                case 3:
                                    oSheet.Cells[25, 31] = dtMatrix.Rows[i]["Mineral"].ToString();
                                    break;

                                default:
                                    break;
                            }
                        }
                    }
                }

                if (dtPheno.Rows.Count > 0)
                {
                    for (int i = 0; i < dtPheno.Rows.Count; i++)
                    {
                        if (i < 4)
                        {
                            switch (i)
                            {
                                case 0:
                                    oSheet.Cells[23, 42] = dtPheno.Rows[i]["Mineral"].ToString();
                                    break;
                                case 1:
                                    oSheet.Cells[23, 46] = dtPheno.Rows[i]["Mineral"].ToString();
                                    break;
                                case 2:
                                    oSheet.Cells[25, 42] = dtPheno.Rows[i]["Mineral"].ToString();
                                    break;
                                case 3:
                                    oSheet.Cells[25, 46] = dtPheno.Rows[i]["Mineral"].ToString();
                                    break;

                                default:
                                    break;
                            }
                        }
                    }
                }
                
                #endregion

                DataTable dtAlterations = new DataTable();
                oAlt.sChid = cmbChannelId.SelectedValue.ToString();
                oAlt.sSample = cmbSample.SelectedValue.ToString();
                dtAlterations = oAlt.getCHAlteration_ListReport();

                #region Alteration
                if (dtAlterations.Rows.Count > 0)
                {
                    for (int i = 0; i < dtAlterations.Rows.Count; i++)
                    {
                        if (i < 2)
                        {
                            switch (i)
                            {
                                case 0:
                                    oSheet.Cells[36, 9] = dtAlterations.Rows[i]["ALTType"].ToString();
                                    oSheet.Cells[36, 13] = dtAlterations.Rows[i]["ALTInt"].ToString();
                                    oSheet.Cells[36, 18] = dtAlterations.Rows[i]["ALTStyle"].ToString();
                                    oSheet.Cells[38, 8] = dtAlterations.Rows[i]["ALTMin"].ToString();
                                    oSheet.Cells[38, 12] = dtAlterations.Rows[i]["ALTMin2"].ToString();
                                    oSheet.Cells[38, 17] = dtAlterations.Rows[i]["ALTMin3"].ToString();
                                    break;
                                case 1:
                                    oSheet.Cells[42, 9] = dtAlterations.Rows[i]["ALTType"].ToString();
                                    oSheet.Cells[42, 13] = dtAlterations.Rows[i]["ALTInt"].ToString();
                                    oSheet.Cells[42, 18] = dtAlterations.Rows[i]["ALTStyle"].ToString();
                                    oSheet.Cells[44, 8] = dtAlterations.Rows[i]["ALTMin"].ToString();
                                    oSheet.Cells[44, 12] = dtAlterations.Rows[i]["ALTMin2"].ToString();
                                    oSheet.Cells[44, 17] = dtAlterations.Rows[i]["ALTMin3"].ToString();
                                    break;

                                default:
                                    break;
                            }
                        }
                    }

                    oSheet.Cells[46, 8] = dtAlterations.Rows[0]["Obsevations"].ToString();
                }
                

                #endregion

                DataTable dtMineralizations = new DataTable();
                oMin.sChid = cmbChannelId.SelectedValue.ToString();
                oMin.sSample = cmbSample.SelectedValue.ToString();
                dtMineralizations = oMin.getCHMineralizationsListReport();

                #region Mineralization
                if (dtMineralizations.Rows.Count > 0)
                {
                    for (int i = 0; i < dtMineralizations.Rows.Count; i++)
                    {
                        if (i < 4)
                        {
                            switch (i)
                            {
                                case 0:
                                    oSheet.Cells[36, 25] = dtMineralizations.Rows[i]["MZMin"].ToString();
                                    oSheet.Cells[36, 29] = dtMineralizations.Rows[i]["MZStyle"].ToString();
                                    oSheet.Cells[36, 32] = dtMineralizations.Rows[i]["MZPerc"].ToString();
                                    break;
                                case 1:
                                    oSheet.Cells[38, 25] = dtMineralizations.Rows[i]["MZMin"].ToString();
                                    oSheet.Cells[38, 29] = dtMineralizations.Rows[i]["MZStyle"].ToString();
                                    oSheet.Cells[38, 32] = dtMineralizations.Rows[i]["MZPerc"].ToString();
                                    break;
                                case 2:
                                    oSheet.Cells[40, 25] = dtMineralizations.Rows[i]["MZMin"].ToString();
                                    oSheet.Cells[40, 29] = dtMineralizations.Rows[i]["MZStyle"].ToString();
                                    oSheet.Cells[40, 32] = dtMineralizations.Rows[i]["MZPerc"].ToString();
                                    break;
                                case 3:
                                    oSheet.Cells[42, 25] = dtMineralizations.Rows[i]["MZMin"].ToString();
                                    oSheet.Cells[42, 29] = dtMineralizations.Rows[i]["MZStyle"].ToString();
                                    oSheet.Cells[42, 32] = dtMineralizations.Rows[i]["MZPerc"].ToString();
                                    break;

                                default:
                                    break;
                            }
                        }
                    }
                    oSheet.Cells[44, 24] = dtMineralizations.Rows[0]["Obsevations"].ToString();
                }
                
                #endregion


                DataTable dtOxides = new DataTable();
                oOxid.sChid = cmbChannelId.SelectedValue.ToString();
                oOxid.sSample = cmbSample.SelectedValue.ToString();
                dtOxides = oOxid.getCHOxidesListReport();
                if (dtOxides.Rows.Count > 0)
                {
                    oSheet.Cells[36, 44] = dtOxides.Rows[0]["GoeStyle"].ToString();
                    oSheet.Cells[36, 47] = dtOxides.Rows[0]["GoePerc"].ToString();
                    oSheet.Cells[38, 44] = dtOxides.Rows[0]["HemStyle"].ToString();
                    oSheet.Cells[38, 47] = dtOxides.Rows[0]["HemPerc"].ToString();
                    oSheet.Cells[40, 44] = dtOxides.Rows[0]["JarStyle"].ToString();
                    oSheet.Cells[40, 47] = dtOxides.Rows[0]["JarPerc"].ToString();
                    oSheet.Cells[42, 44] = dtOxides.Rows[0]["LimStyle"].ToString();
                    oSheet.Cells[42, 47] = dtOxides.Rows[0]["LimPerc"].ToString();
                    oSheet.Cells[44, 42] = dtOxides.Rows[0]["Observations"].ToString();
                }

                

                DataTable dtStructures = new DataTable();
                oStr.sChid = cmbChannelId.SelectedValue.ToString();
                oStr.sSample = cmbSample.SelectedValue.ToString();
                dtStructures = oStr.getCHStructuresListReport();

                #region Structure
                if (dtStructures.Rows.Count > 0)
                {
                    for (int i = 0; i < dtStructures.Rows.Count; i++)
                    {
                        if (i < 3)
                        {
                            switch (i)
                            {
                                case 0:
                                    oSheet.Cells[51, 6] = dtStructures.Rows[i]["StrType"].ToString();
                                    oSheet.Cells[51, 10] = dtStructures.Rows[i]["StrDip"].ToString();
                                    oSheet.Cells[51, 15] = dtStructures.Rows[i]["StrDipAz"].ToString();
                                    oSheet.Cells[51, 19] = dtStructures.Rows[i]["StrAThick"].ToString();
                                    oSheet.Cells[51, 22] = dtStructures.Rows[i]["StrRThick"].ToString();
                                    oSheet.Cells[51, 25] = dtStructures.Rows[i]["StrFill"].ToString();
                                    oSheet.Cells[51, 28] = dtStructures.Rows[i]["StrFill2"].ToString();
                                    oSheet.Cells[51, 31] = dtStructures.Rows[i]["StrFill3"].ToString();
                                    oSheet.Cells[51, 34] = dtStructures.Rows[i]["StrNumber"].ToString();
                                    oSheet.Cells[51, 36] = dtStructures.Rows[i]["StrDensity"].ToString();
                                    break;
                                case 1:
                                    oSheet.Cells[53, 6] = dtStructures.Rows[i]["StrType"].ToString();
                                    oSheet.Cells[53, 10] = dtStructures.Rows[i]["StrDip"].ToString();
                                    oSheet.Cells[53, 15] = dtStructures.Rows[i]["StrDipAz"].ToString();
                                    oSheet.Cells[53, 19] = dtStructures.Rows[i]["StrAThick"].ToString();
                                    oSheet.Cells[53, 22] = dtStructures.Rows[i]["StrRThick"].ToString();
                                    oSheet.Cells[53, 25] = dtStructures.Rows[i]["StrFill"].ToString();
                                    oSheet.Cells[53, 28] = dtStructures.Rows[i]["StrFill2"].ToString();
                                    oSheet.Cells[53, 31] = dtStructures.Rows[i]["StrFill3"].ToString();
                                    oSheet.Cells[53, 34] = dtStructures.Rows[i]["StrNumber"].ToString();
                                    oSheet.Cells[53, 36] = dtStructures.Rows[i]["StrDensity"].ToString();
                                    break;
                                case 2:
                                    oSheet.Cells[55, 6] = dtStructures.Rows[i]["StrType"].ToString();
                                    oSheet.Cells[55, 10] = dtStructures.Rows[i]["StrDip"].ToString();
                                    oSheet.Cells[55, 15] = dtStructures.Rows[i]["StrDipAz"].ToString();
                                    oSheet.Cells[55, 19] = dtStructures.Rows[i]["StrAThick"].ToString();
                                    oSheet.Cells[55, 22] = dtStructures.Rows[i]["StrRThick"].ToString();
                                    oSheet.Cells[55, 25] = dtStructures.Rows[i]["StrFill"].ToString();
                                    oSheet.Cells[55, 28] = dtStructures.Rows[i]["StrFill2"].ToString();
                                    oSheet.Cells[55, 31] = dtStructures.Rows[i]["StrFill3"].ToString();
                                    oSheet.Cells[55, 34] = dtStructures.Rows[i]["StrNumber"].ToString();
                                    oSheet.Cells[55, 36] = dtStructures.Rows[i]["StrDensity"].ToString();
                                    break;

                                default:
                                    break;
                            }
                        }
                    }
                    oSheet.Cells[57, 8] = dtStructures.Rows[0]["Obsevations"].ToString();
                }
               
                #endregion


                DataTable dtSurvey = new DataTable();
                oSur.sOpcion = "2";
                oSur.sChId = cmbChannelId.SelectedValue.ToString();
                oSur.sSample = cmbSample.SelectedValue.ToString();
                dtSurvey = oSur.getCH_Surveys();

                
                if (dtSurvey.Rows.Count > 0)
                {

                    oSheet.Cells[7, 61] = dtSurvey.Rows[0]["Azm"].ToString();
                    oSheet.Cells[7, 66] = dtSurvey.Rows[0]["Dip"].ToString();

                    //oSheet.Cells[11, 54] = dtSurvey.Rows[0]["P1E"].ToString();
                    //oSheet.Cells[11, 67] = dtSurvey.Rows[0]["P2E"].ToString();
                    //oSheet.Cells[13, 54] = dtSurvey.Rows[0]["P1N"].ToString();
                    //oSheet.Cells[13, 67] = dtSurvey.Rows[0]["P2N"].ToString();
                    //oSheet.Cells[15, 54] = dtSurvey.Rows[0]["P1Z"].ToString();
                    //oSheet.Cells[15, 67] = dtSurvey.Rows[0]["P2Z"].ToString();

                    switch (dtSurvey.Rows[0]["MineLocation"].ToString())
                    {
                        case "Wall":
                            oSheet.Cells[13, 80] = "X";
                            break;
                        case "FaceCut":
                            oSheet.Cells[15, 80] = "X";
                            break;
                        case "Flor":
                            oSheet.Cells[17, 80] = "X";
                            break;
                        default:
                            break;
                    }

                    
                }

               
                DataTable dtMineEnt = oRf.getMineEntranceList(1, dtChannel.Rows[0]["MineID"].ToString());

                if (dtMineEnt.Rows.Count > 0)
                {
                    oSheet.Cells[36, 56] = dtMineEnt.Rows[0]["Name"].ToString();
                    //oSheet.Cells[38, 62] = dtMineEnt.Rows[0]["Coordenadas"].ToString();    
                }

                //oSheet.Cells[36, 56] = dtSample.Rows[0]["Mine"].ToString();
                //oSheet.Cells[38, 62] = dtSample.Rows[0]["MineEntrance"].ToString();    


                oXL.Visible = true;
                oXL.UserControl = true;

                MessageBox.Show("Successful Export");
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Opcion 1= Matrix, Opcion 2= PhenoCryst
        /// </summary>
        /// <param name="_sOpcion"></param>
        /// <returns></returns>
        private DataTable getMinerals_Ph_Mx(string _sOpcion)
        {
            try
            {
                DataTable dtResp = new DataTable();
                if (_sOpcion == "1")
                {
                    //Matrix
                    oMinLith.sOpcion = _sOpcion;
                    oMinLith.sSample = cmbSample.SelectedValue.ToString();
                    dtResp = oMinLith.getGCSamplesRockLithList();

                }
                else if (_sOpcion == "2")
                {
                    //Phenocryst
                    oMinLith.sOpcion = _sOpcion;
                    oMinLith.sSample = cmbSample.SelectedValue.ToString();
                    dtResp = oMinLith.getGCSamplesRockLithList();
                }

                return dtResp;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void ExportExcelGeoch()
        {
            try
            {
                switch (sExport)
                {
                    case "Geochemistry":
                        ExpGeochemistry();
                        //MessageBox.Show("Export " + sExport.ToString());
                        break;


                    default:
                        Console.WriteLine("Default case");
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnCancelStr_Click(object sender, EventArgs e)
        {
            try
            {
                sEditStr = "0";
                CleanControlsStr();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(cmbSample.Text.ToString());
            if (cmbSample.SelectedValue != null)
            {
                MessageBox.Show(cmbSample.SelectedValue.ToString());
            }
            else
            { MessageBox.Show("No hay seleccion en el combo de sample"); }
            
        }

        private void txtToHeader_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtToHeader.Text == "")
                {
                    txtLenght.Text = "";
                    return;
                }

                if (txtFromHeader.Text == "")
                {
                    txtLenght.Text = "";
                    return;
                }

                txtLenght.Text = (double.Parse(txtToHeader.Text.ToString()) -
                    double.Parse(txtFromHeader.Text.ToString())).ToString();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnActualizar_Click(object sender, EventArgs e)
        {
            try
            {
                frmChannels ofrm = new frmChannels();
                ofrm.MdiParent = this.MdiParent;
                ofrm.Show();              
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelAlt_Click(object sender, EventArgs e)
        {
            try
            {
                sEditAlt = "0";
                CleanControlsAlt();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelMin_Click(object sender, EventArgs e)
        {
            try
            {
                sEditMin = "0";
                 CleanControlsMin();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelOx_Click(object sender, EventArgs e)
        {
            try
            {
                sEditOxid = "0";
                CleanControlsOxides();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void CleanControlsAll()
        {
            try
            {

                CleanControlsHeader();

                sEditSur = "0";
                CleanControlsSurvey();

                sEditAlt = "0";
                CleanControlsAlt();

                sEditMin = "0";
                CleanControlsMin();

                sEditOxid = "0";
                CleanControlsOxides();

                sEditStr = "0";
                CleanControlsStr();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cmbSample_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {

                LoadData_CHLith();
                LoadDgSurveys();
                LoadDataAlterations("2");
                LoadDataMineralizations("2");
                LoadDataOxides("2");
                LoadDataStructures("2");

                oCHSamp.sSample = cmbSample.SelectedValue.ToString();
                oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
                DataTable dtSamp = LoadDataCHSurv(cmbSample.SelectedValue.ToString());
                //oCHSamp.getCHSamplesBySample();


                CleanControlsAll();

                if (dtSamp != null)
                {
                    if (dtSamp.Rows.Count > 0)
                    {
                        txtToSur.Text = dtSamp.Rows[0]["To"].ToString();
                    }
                }


                if (cmbSample.SelectedValue.ToString() != "Select an option..")
                {
                    DataTable dgSamp = (DataTable)dgData.DataSource;
                    DataRow[] myRow = dgSamp.Select(@"Sample = '" + cmbSample.SelectedValue.ToString() + "'");
                    int rowindex = dgSamp.Rows.IndexOf(myRow[0]);
                    dgData.Rows[rowindex].Selected = true;
                    dgData.CurrentCell = dgData.Rows[rowindex].Cells[1];

                    dgD_CellClick(rowindex, dgData, "Cmb");
                }

                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void txtMatrixPerc_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtPhenoPerc_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtPhenoPerc_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtPhenoPerc.Text != "")
                {
                    if (double.Parse(txtPhenoPerc.Text.ToString()) > 100)
                    {
                        MessageBox.Show("Perc Pheno > 100");
                        txtPhenoPerc.Text = "";
                        txtPhenoPerc.Focus();
                    }
                }

                if (txtPhenoPerc.Text != "" && txtMatrixPerc.Text != "")
                {
                    if (double.Parse(txtPhenoPerc.Text.ToString()) +
                        double.Parse(txtMatrixPerc.Text.ToString()) > 100)
                    {
                        MessageBox.Show("Perc Pheno + Perc Matrix > 100");
                        txtPhenoPerc.Text = "";
                        txtPhenoPerc.Focus();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtMatrixPerc_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtMatrixPerc.Text != "")
                {
                    if (double.Parse(txtMatrixPerc.Text.ToString()) > 100)
                    {
                        MessageBox.Show("Perc Matrix > 100");
                        txtMatrixPerc.Text = "";
                        txtMatrixPerc.Focus();
                    }
                }

                if (txtPhenoPerc.Text != "" && txtMatrixPerc.Text != "")
                {
                    if (double.Parse(txtPhenoPerc.Text.ToString()) +
                        double.Parse(txtMatrixPerc.Text.ToString()) > 100)
                    {
                        MessageBox.Show("Perc Pheno + Perc Matrix > 100");
                        txtPhenoPerc.Text = "";
                        txtPhenoPerc.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtMinPerc_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtMinPerc_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtMinPerc.Text != "")
                {
                    if (double.Parse(txtMinPerc.Text) > 100)
                    {
                        MessageBox.Show("Percentage isn´t more than 100");
                        txtMinPerc.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtCountSample_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                e.Handled = Keypress(e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtChId_Leave(object sender, EventArgs e)
        {
            foreach (DataGridViewRow Row in dgDataCh.Rows)
            {
                var strFila = Row.Index;
                string Valor = Convert.ToString(Row.Cells[1].Value);

                if (Valor == txtChId.Text.ToUpper())
                {
                    BuscarChannerPorId(strFila);
                }
            }
        }

        private void BuscarChannerPorId(int index)
        {
            try
            {
                oCh.iSKCHChannels = Int64.Parse(dgDataCh.Rows[index].Cells["SKCHChannels"].Value.ToString());
                sEditCh = "1";

                cmbChannelId.SelectedValue = dgDataCh.Rows[index].Cells["Chid"].Value.ToString();
                txtChId.Text = dgDataCh.Rows[index].Cells["Chid"].Value.ToString();
                txtLenghtCh.Text = dgDataCh.Rows[index].Cells["Length"].Value.ToString();
                txtEastCh.Text = dgDataCh.Rows[index].Cells["East"].Value.ToString();
                txtNorthCh.Text = dgDataCh.Rows[index].Cells["North"].Value.ToString();
                txtElevationCh.Text = dgDataCh.Rows[index].Cells["Elevation"].Value.ToString();
                txtProjectionCh.Text = dgDataCh.Rows[index].Cells["Projection"].Value.ToString();
                txtDatumCh.Text = dgDataCh.Rows[index].Cells["Datum"].Value.ToString();
                txtProjectCh.Text = dgDataCh.Rows[index].Cells["Project"].Value.ToString();
                txtClaimCh.Text = dgDataCh.Rows[index].Cells["Claim"].Value.ToString();

                dtStartDateCh.Text =
                    dgDataCh.Rows[index].Cells["Star_Date"].Value.ToString() == string.Empty
                    ? DateTime.Now.ToShortDateString()
                    : dtStartDateCh.Text = Convert.ToDateTime(dgDataCh.Rows[index].Cells["Star_Date"].Value).ToString("dd/MM/yyyy");

                dtFinalDateCh.Text =
                    dgDataCh.Rows[index].Cells["Final_Date"].Value.ToString() == string.Empty
                    ? DateTime.Now.ToShortDateString()
                    : dtFinalDateCh.Text = Convert.ToDateTime(dgDataCh.Rows[index].Cells["Final_Date"].Value).ToString("dd/MM/yyyy");

                //txtPurposeCh.Text = dgDataCh.Rows[index].Cells["Purpose"].Value.ToString();
                txtStorageCh.Text = dgDataCh.Rows[index].Cells["Storage"].Value.ToString();
                txtSourceCh.Text = dgDataCh.Rows[index].Cells["Source"].Value.ToString();

                txtCommentsCh.Text = dgDataCh.Rows[index].Cells["Comments"].Value.ToString();

                cmbMineEntrance.SelectedValue = dgDataCh.Rows[index].Cells["MineID"].Value.ToString() == "" ? "-1" :
                    dgDataCh.Rows[index].Cells["MineID"].Value.ToString();

                cmbChannelType.SelectedValue = dgDataCh.Rows[index].Cells["Type"].Value.ToString() == "" ? "-1" :
                    dgDataCh.Rows[index].Cells["Type"].Value.ToString();

                cmbInstSur.Text = dgDataCh.Rows[index].Cells["Instrument"].Value.ToString() == "" ? "TS" :
                    dgDataCh.Rows[index].Cells["Instrument"].Value.ToString();

                dTimerDateSur.Text =
                   dgDataCh.Rows[index].Cells["Date_Survey"].Value.ToString() == string.Empty
                   ? DateTime.Now.ToShortDateString()
                   : dTimerDateSur.Text = Convert.ToDateTime(dgDataCh.Rows[index].Cells["Date_Survey"].Value).ToString("dd/MM/yyyy");
                
                txtTotalSamples.Text = dgDataCh.Rows[index].Cells["SamplesTotal"].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



    }
}
