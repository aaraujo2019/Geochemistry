using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Geochemistry
{
    public partial class frmRock : Form
    {
        clsGCSamplesRock oGCRock = new clsGCSamplesRock();
        clsGCSamplesRockLith oGCRLith = new clsGCSamplesRockLith();
        clsGCSamplesRockAlterations oAlt = new clsGCSamplesRockAlterations();
        clsGCSamplesRockMiner oMin = new clsGCSamplesRockMiner();
        clsGCSamplesRockOxides oOxid = new clsGCSamplesRockOxides();
        clsGCSamplesRockStructures oStr = new clsGCSamplesRockStructures();
        clsRf oRf = new clsRf();
        static string sEdit = "0";
        static string sEditLithMat = "0";
        static string sEditLthPhe = "0";
        static string sEditAlt = "0";
        static string sEditMin = "0";
        static string sEditOxid = "0";
        static string sEditStr = "0";
        static string sMineLoc = "0";
        static string sExport = "";
        private bool swCargado = false;

        public frmRock()
        {
            InitializeComponent();
            Loadcmb();
            LoadDataMinLith("1");
            LoadDataMinLith("2");
            LoadDataAlterations("1");
        }

        private void Loadcmb()
        {
            try
            {
                txtProject.Text = ConfigurationSettings.AppSettings["IDProjectGC"].ToString(); //Id Proyecto Gran Colombia. Ej GSG, GZG ...

                #region sample

                DataTable dtSamples = new DataTable();
                oGCRock.sOpcion = "1";
                dtSamples = oGCRock.getGCSamplesRockListAll();
                DataRow drSamp = dtSamples.NewRow();
                drSamp[0] = "Select an option..";
                dtSamples.Rows.Add(drSamp);
                cmbSample.DisplayMember = "sample";
                cmbSample.ValueMember = "sample";
                cmbSample.DataSource = dtSamples;
                cmbSample.SelectedValue = "Select an option..";
                swCargado = true;

                #endregion

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

                //DataRow[] datoU = dtUsers.Select("grupo in ('Administradores','Perforacion')");
                //DataTable dtToR = dtUsers.Clone();

                //foreach (DataRow r in datoU)
                //{
                //    dtToR.ImportRow(r);
                //}

                DataRow dr2 = dtUsers.NewRow();
                dr2[0] = "-1";
                dr2[7] = "Select an option..";
                dtUsers.Rows.Add(dr2);

                cmbGeologist.DisplayMember = "cmb";
                cmbGeologist.ValueMember = "id";
                cmbGeologist.DataSource = dtUsers;
                cmbGeologist.SelectedValue = -1;

                #endregion

                #region SampleType

                DataSet dtSampleT = new DataSet();
                dtSampleT = oRf.getRfTypeSampleDataSet();

                DataRow dr = dtSampleT.Tables[1].NewRow();
                dr[0] = "-1";
                dr[1] = "Select an option..";
                dtSampleT.Tables[1].Rows.Add(dr);
                cmbSampleType.DisplayMember = "Comb";
                cmbSampleType.ValueMember = "Code";
                cmbSampleType.DataSource = dtSampleT.Tables[1];
                cmbSampleType.SelectedValue = -1;


                DataRow drT2 = dtSampleT.Tables[4].NewRow();
                drT2[0] = "-1";
                drT2[1] = "Select an option..";
                dtSampleT.Tables[4].Rows.Add(drT2);
                cmbSamplingType.DisplayMember = "Comb";
                cmbSamplingType.ValueMember = "Code";
                cmbSamplingType.DataSource = dtSampleT.Tables[4];
                cmbSamplingType.SelectedValue = -1;

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

                CargarCombosPerc(dtMinPerc, cmbMatrixPorc);
                CargarCombosPerc(dtMinPerc, cmbPhenoPerc);
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

                DataTable dtData = LoadDataRocksAll("2");
                dgData.DataSource = dtData;
                LoadDataRocks(txtSample.Text.ToString());
                //dgData.Columns["SKSamplesRock"].Visible = false;



                //DataTable dtData = LoadDataRocksAll("2");

                //// Query the SalesOrderHeader table for orders placed 
                //// after August 8, 2001.
                //IEnumerable<DataRow> query =
                //    from dtDat in dtData.AsEnumerable()
                //    where dtDat.Field<String>("Sample") == cmbSample.SelectedValue.ToString()
                //    select dtDat;

                //// Create a table from the query.
                //DataTable boundTable = query.CopyToDataTable<DataRow>();

                //// Bind the table to a System.Windows.Forms.BindingSource object, 
                //// which acts as a proxy for a System.Windows.Forms.DataGridView object.
                ////bindingSource.DataSource = boundTable;

                //dgLithology.DataSource = boundTable;//LoadDataRocks(txtSample.Text.ToString());
                //dgLithology.Columns["SKSamplesRock"].Visible = false;


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


                #region Minerals
                DataTable dtMineral = new DataTable();
                dtMineral = oRf.getRfMinerMin_List();
                DataRow drM = dtMineral.NewRow();
                drM[0] = "-1";
                drM[1] = "Select an option..";
                dtMineral.Rows.Add(drM);
                LoadCombos(dtMineral, cmbMineralPh);
                LoadCombos(dtMineral, cmbMineralMt);
                LoadCombos(dtMineral, cmbMin1Alt);
                LoadCombos(dtMineral, cmbMin2Alt1);
                LoadCombos(dtMineral, cmbMin3Alt1);
                LoadCombos(dtMineral, cmbMineralmin);
                #endregion

                #region Style
                DataTable dtMinStyle = new DataTable();
                dtMinStyle = oRf.getRfMinerMinSt_List();
                DataRow drMin = dtMinStyle.NewRow();
                drMin[0] = "-1";
                drMin[1] = "Select an option..";
                dtMinStyle.Rows.Add(drMin);

                LoadCombos(dtMinStyle, cmbStyleM);


                DataTable dtMinInt = new DataTable();
                dtMinInt = oRf.getRfOxidationInt_List();
                DataRow drMinI = dtMinInt.NewRow();
                drMinI[0] = "-1";
                drMinI[1] = "Select an option..";
                dtMinInt.Rows.Add(drMinI);
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
                //LoadCombos(dtMinStyle, cmbStyleAlt12);
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

        private DataTable LoadDataRocksAll(string _sOpcion)
        {
            try
            {
                oGCRock.sOpcion = _sOpcion.ToString();
                DataTable dtRocks = oGCRock.getGCSamplesRockListAll();
                return dtRocks;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private DataTable LoadDataRocks(string _sSample)
        {
            try
            {
                oGCRock.sSample = _sSample.ToString();
                DataTable dtRocks = oGCRock.getGCSamplesRockList_Sample();
                return dtRocks;
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

        private void frmRock_Load(object sender, EventArgs e)
        {
        
        }

        private void AddHeaderLithology()
        {
            try
            {

                if (sEdit == "0")
                {
                    oGCRock.sOpcion = "1";
                    oGCRock.iSKSamplesRock = 0;
                }
                else
                {
                    oGCRock.sOpcion = "2";
                }


                oGCRock.sSample = txtSample.Text.ToString();
                oGCRock.sTarget = cmbTarget.SelectedValue.ToString();
                oGCRock.sLocation = txtLocation.Text.ToString();
                oGCRock.sProject = txtProject.Text.ToString();


                if (cmbGeologist.SelectedValue != null)
                {
                    if (cmbGeologist.SelectedValue.ToString() == "-1" ||
                    cmbGeologist.SelectedValue.ToString() == "")
                        oGCRock.sGeologist = null;
                    else
                        oGCRock.sGeologist = cmbGeologist.SelectedValue.ToString();

                }
                else oGCRock.sGeologist = null; 
                
                
                oGCRock.sHelper = txtHelper.Text.ToString();
                oGCRock.sStation = txtStation.Text.ToString();

                //DateTime dDate = dtimeDate.Value;
                //string sDate = dDate.Year.ToString().PadLeft(4, '0') + dDate.Month.ToString().PadLeft(2, '0') +
                //    dDate.Day.ToString().PadLeft(2, '0');

                oGCRock.sDate = dtimeDate.Value.ToString();
                //sDate.ToString();

                if (txtCoordE.Text.ToString() == "")
                    oGCRock.dCoordE = null;
                else
                    oGCRock.dCoordE = double.Parse(txtCoordE.Text.ToString());

                if (txtCoordN.Text.ToString() == "")
                    oGCRock.dCoordN = null;
                else
                    oGCRock.dCoordN = double.Parse(txtCoordN.Text.ToString());

                if (txtCoordElevation.Text.ToString() == "")
                    oGCRock.dCoordZ = null;
                else
                    oGCRock.dCoordZ = double.Parse(txtCoordElevation.Text.ToString());

                if (cmbCS.SelectedValue != null)
                {

                    if (cmbCS.SelectedValue.ToString() == "-1" ||
                        cmbCS.SelectedValue.ToString() == "")
                        oGCRock.sCs = null;
                    else
                        oGCRock.sCs = cmbCS.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sCs = null;
                }

                if (txtGPSEPE.Text.ToString() == "")
                    oGCRock.dGPSepe = null;
                else
                    oGCRock.dGPSepe = double.Parse(txtGPSEPE.Text.ToString());

                if (txtPhoto.Text.ToString() == "")
                    oGCRock.sPhoto = null;
                else
                    oGCRock.sPhoto = txtPhoto.Text.ToString();

                if (txtPhotoAzimuth.Text.ToString() == "")
                    oGCRock.sPhoto_Azimuth = null;
                else
                    oGCRock.sPhoto_Azimuth = txtPhotoAzimuth.Text.ToString();

                if (cmbSampleType.SelectedValue != null)
                {

                    if (cmbSampleType.SelectedValue.ToString() == "-1" ||
                        cmbSampleType.SelectedValue.ToString() == "")
                        oGCRock.sSampleType = null;
                    else
                        oGCRock.sSampleType = cmbSampleType.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sSampleType = null;
                }

                if (cmbNotInSitu.SelectedValue != null)
                {

                    if (cmbNotInSitu.SelectedValue.ToString() == "-1" ||
                        cmbNotInSitu.SelectedValue.ToString() == "")
                        oGCRock.sNotInSitu = null;
                    else
                        oGCRock.sNotInSitu = cmbNotInSitu.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sNotInSitu = null;
                }

                if (cmbPorpuose.SelectedValue != null)
                {

                    if (cmbPorpuose.SelectedValue.ToString() == "-1" ||
                        cmbPorpuose.SelectedValue.ToString() == "")
                        oGCRock.sPorpouse = null;
                    else
                        oGCRock.sPorpouse = cmbPorpuose.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sPorpouse = null;
                }

                if (cmbRelativeLoc.SelectedValue != null)
                {

                    if (cmbRelativeLoc.SelectedValue.ToString() == "-1" ||
                        cmbRelativeLoc.SelectedValue.ToString() == "")
                        oGCRock.sRelativeLoc = null;
                    else
                        oGCRock.sRelativeLoc = cmbRelativeLoc.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sRelativeLoc = null;
                }

                if (txtLenght.Text.ToString() == "")
                    oGCRock.dLenght = null;
                else
                    oGCRock.dLenght = double.Parse(txtLenght.Text.ToString());

                if (txtHigh.Text.ToString() == "")
                    oGCRock.dHigh = null;
                else
                    oGCRock.dHigh = double.Parse(txtHigh.Text.ToString());

                if (txtThickness.Text.ToString() == "")
                    oGCRock.sThickness = null;
                else
                    oGCRock.sThickness = txtThickness.Text.ToString();

                if (txtObservations.Text.ToString() == "")
                    oGCRock.sObservations = null;
                else
                    oGCRock.sObservations = txtObservations.Text.ToString();

                if (cmbLithologyLit.SelectedValue != null)
                {

                    if (cmbLithologyLit.SelectedValue.ToString() == "" ||
                        cmbLithologyLit.SelectedValue.ToString() == "-1")
                        oGCRock.sLRock = null;
                    else
                        oGCRock.sLRock = cmbLithologyLit.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLRock = null;
                }

                if (cmbLTextures.SelectedValue != null)
                {

                    if (cmbLTextures.SelectedValue.ToString() == "-1" || cmbLTextures.SelectedValue.ToString() == "")
                        oGCRock.sLTexture = null;
                    else
                        oGCRock.sLTexture = cmbLTextures.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLTexture = null;
                }


                if (cmbLGsize.SelectedValue != null)
                {

                    if (cmbLGsize.SelectedValue.ToString() == "-1" || cmbLGsize.SelectedValue.ToString() == "")
                        oGCRock.sLGSize = null;
                    else
                        oGCRock.sLGSize = cmbLGsize.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLGSize = null;
                }

                if (cmbLWeathering.SelectedValue != null)
                {

                    if (cmbLWeathering.SelectedValue.ToString() == "-1" || cmbLWeathering.SelectedValue.ToString() == "")
                        oGCRock.sLWeathering = null;
                    else
                        oGCRock.sLWeathering = cmbLWeathering.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLWeathering = null;
                }


                if (cmbRSorting.SelectedValue != null)
                {

                    if (cmbRSorting.SelectedValue.ToString() == "-1" || cmbRSorting.SelectedValue.ToString() == "")
                        oGCRock.sLRockSorting = null;
                    else
                        oGCRock.sLRockSorting = cmbRSorting.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLRockSorting = null;
                }

                if (cmbRSphericity.SelectedValue != null)
                {

                    if (cmbRSphericity.SelectedValue.ToString() == "-1" || cmbRSphericity.SelectedValue.ToString() == "")
                        oGCRock.sLRockSphericity = null;
                    else
                        oGCRock.sLRockSphericity = cmbRSphericity.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLRockSphericity = null;
                }


                if (cmbRounding.SelectedValue != null)
                {

                    if (cmbRounding.SelectedValue.ToString() == "-1" || cmbRounding.SelectedValue.ToString() == "")
                        oGCRock.sLRockRounding = null;
                    else
                        oGCRock.sLRockRounding = cmbRounding.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLRockRounding = null;
                }

                if (txtObservSedimentary.Text.ToString() == "")
                    oGCRock.sLRockObservation = null;
                else
                    oGCRock.sLRockObservation = txtObservSedimentary.Text.ToString();

                if (cmbMatrixPorc.SelectedValue != null)
                {

                    if (cmbMatrixPorc.SelectedValue.ToString() == "-1" || cmbMatrixPorc.SelectedValue.ToString() == "")
                        oGCRock.sLMatrixPerc = null;
                    else
                        oGCRock.sLMatrixPerc = cmbMatrixPorc.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLMatrixPerc = null;
                }

                if (cmbMatrixGSize.SelectedValue != null)
                {

                    if (cmbMatrixGSize.SelectedValue.ToString() == "-1" || cmbMatrixGSize.SelectedValue.ToString() == "")
                        oGCRock.sLMatrixGSize = null;
                    else
                        oGCRock.sLMatrixGSize = cmbMatrixGSize.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLMatrixGSize = null;
                }

                if (txtMatrixObserv.Text.ToString() == "")
                    oGCRock.sLMatrixObservations = null;
                else
                    oGCRock.sLMatrixObservations = txtMatrixObserv.Text.ToString();

                if (cmbPhenoPerc.SelectedValue != null)
                {

                    if (cmbPhenoPerc.SelectedValue.ToString() == "-1" || cmbPhenoPerc.SelectedValue.ToString() == "")
                        oGCRock.sLPhenoCPerc = null;
                    else
                        oGCRock.sLPhenoCPerc = cmbPhenoPerc.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLPhenoCPerc = null;
                }

                if (cmbPhenoGSize.SelectedValue != null)
                {

                    if (cmbPhenoGSize.SelectedValue.ToString() == "-1" || cmbPhenoGSize.SelectedValue.ToString() == "")
                        oGCRock.sLPhenoCGSize = null;
                    else
                        oGCRock.sLPhenoCGSize = cmbPhenoGSize.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sLPhenoCGSize = null;
                }

                if (txtPhenoObserv.Text.ToString() == "")
                    oGCRock.sLPhenoCObservations = null;
                else
                    oGCRock.sLPhenoCObservations = txtPhenoObserv.Text.ToString();

                if (cmbContactType.SelectedValue != null)
                {
                    if (cmbContactType.SelectedValue.ToString() == "" ||
                        cmbContactType.SelectedValue.ToString() == "-1")
                        oGCRock.sVContactType = null;
                    else
                        oGCRock.sVContactType = cmbContactType.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sVContactType = null;
                }

                if (cmbVeinName.SelectedValue != null)
                {
                    if (cmbVeinName.SelectedValue.ToString() == "" ||
                    cmbVeinName.SelectedValue.ToString() == "-1")
                        oGCRock.sVVeinName = null;
                    else
                        oGCRock.sVVeinName = cmbVeinName.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sVVeinName = null;
                }

                if (cmbHostRock.SelectedValue != null)
                {
                    if (cmbHostRock.SelectedValue.ToString() == "" ||
                    cmbHostRock.SelectedValue.ToString() == "-1")
                        oGCRock.sVHostRock = null;
                    else
                        oGCRock.sVHostRock = cmbHostRock.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sVHostRock = null;
                }

                
                if (txtVeinObserv.Text.ToString() == "")
                    oGCRock.sVObservations = null;
                else
                    oGCRock.sVObservations = txtVeinObserv.Text.ToString();

                if (cmbSamplingType.SelectedValue != null)
                {
                    if (cmbSamplingType.SelectedValue.ToString() == "" ||
                        cmbSamplingType.SelectedValue.ToString() == "-1")
                        oGCRock.sSamplingType = null;
                    else
                        oGCRock.sSamplingType = cmbSamplingType.SelectedValue.ToString();
                }
                else
                {
                    oGCRock.sSamplingType = null;
                }
                

                if (txtDupOf.Text.ToString() == "")
                    oGCRock.sDupOf = null;
                else
                    oGCRock.sDupOf = txtDupOf.Text.ToString();


                if (cmbSamplingType.SelectedValue.ToString().Substring(0, 1) == "O")
                {
                    //bMineLocation(false);
                    oGCRock.sMine = txtMineLocation.Text.ToString() ;
                }
                else
                {
                    //bMineLocation(true);
                    oGCRock.sMine = cmbMineLocation.SelectedValue.ToString() ;
                }

                string sResp = oGCRock.GCSamplesRock_Add();
                if (sResp == "OK")
                {
                    
                    dgData.DataSource = LoadDataRocksAll("2");
                    dgData.Columns["SKSamplesRock"].Visible = false;
                    //dgLithology.DataSource = LoadDataRocksAll("2");
                    //dgLithology.Columns["SKSamplesRock"].Visible = false;

                    if (sEdit == "1")
                    {
                        if (dgData.Rows.Count > 1)
                        {
                            DataTable dtSamp = (DataTable)dgData.DataSource;
                            DataRow[] myRow = dtSamp.Select(@"SKSamplesRock = '" + oGCRock.iSKSamplesRock.ToString() + "'");
                            int rowindex = dtSamp.Rows.IndexOf(myRow[0]);
                            dgData.Rows[rowindex].Selected = true;
                            dgData.CurrentCell = dgData.Rows[rowindex].Cells[1];


                            DataTable dtData = LoadDataRocksAll("2");
                            // Query the SalesOrderHeader table for orders placed 
                            // after August 8, 2001.
                            IEnumerable<DataRow> query =
                                from dtDat in dtData.AsEnumerable()
                                where dtDat.Field<String>("Sample") == cmbSample.SelectedValue.ToString()
                                select dtDat;

                            DataTable boundTable = new DataTable();
                            if (query.Count() > 0)
                            {
                                // Create a table from the query.
                                boundTable = query.CopyToDataTable<DataRow>();
                                dgLithology.DataSource = boundTable;//LoadDataRocks(txtSample.Text.ToString());
                                dgLithology.Columns["SKSamplesRock"].Visible = false;
                            }
                            else
                            {
                                boundTable = null;
                                dgLithology.DataSource = boundTable;//LoadDataRocks(txtSample.Text.ToString());
                            }

                            //DataTable dtSamp2 = (DataTable)dgLithology.DataSource;
                            //DataRow[] myRow2 = dtSamp2.Select(@"SKSamplesRock = '" + oGCRock.iSKSamplesRock.ToString() + "'");
                            //int rowindex2 = dtSamp2.Rows.IndexOf(myRow2[0]);
                            //dgLithology.Rows[rowindex2].Selected = true;
                            //dgLithology.CurrentCell = dgLithology.Rows[rowindex2].Cells[1];

                        }
                    }

                    ControlsClean();
                    sEdit = "0";

                 
                }
                else
                {
                    MessageBox.Show("Save Error" + sResp.ToString(), "Geochemistry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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
                AddHeaderLithology();
                btnCancel_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private void txtSample_Leave(object sender, EventArgs e)
        {
            try
            {
                dgData.DataSource = LoadDataRocksAll("2"); //LoadDataRocks(txtSample.Text.ToString());
                dgData.Columns["SKSamplesRock"].Visible = false;
                dgLithology.DataSource = LoadDataRocksAll("2"); //LoadDataRocks(txtSample.Text.ToString());
                dgLithology.Columns["SKSamplesRock"].Visible = false;
                LoadDataMinLith("1"); LoadDataMinLith("2"); LoadDataAlterations("1");
                LoadDataMineralizations("1"); LoadDataOxides("1"); LoadDataStructures("1");
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

                oGCRock.iSKSamplesRock = int.Parse(dgData.Rows[e.RowIndex].Cells["SKSamplesRock"].Value.ToString());
                sEdit = "1";

                DateTime dDate =
                    dgData.Rows[e.RowIndex].Cells["Date"].Value.ToString() == ""
                    ? DateTime.Parse("1900/01/01")
                    : DateTime.Parse(dgData.Rows[e.RowIndex].Cells["Date"].Value.ToString());
                dtimeDate.Value = dDate;
                dtimeDate.Text = dgData.Rows[e.RowIndex].Cells["Date"].Value.ToString();
                
                txtSample.Text = dgData.Rows[e.RowIndex].Cells["Sample"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["Target"].Value.ToString() == "")
                    cmbTarget.SelectedValue = "-1";
                else cmbTarget.SelectedValue = dgData.Rows[e.RowIndex].Cells["Target"].Value.ToString();
                
                txtLocation.Text = dgData.Rows[e.RowIndex].Cells["Location"].Value.ToString();
                txtProject.Text = dgData.Rows[e.RowIndex].Cells["Project"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["Geologist"].Value.ToString() == "")
                    cmbGeologist.SelectedValue = "-1";
                else cmbGeologist.SelectedValue = dgData.Rows[e.RowIndex].Cells["Geologist"].Value.ToString();
                
                txtHelper.Text = dgData.Rows[e.RowIndex].Cells["Helper"].Value.ToString();
                txtStation.Text = dgData.Rows[e.RowIndex].Cells["Station"].Value.ToString();
                txtCoordE.Text = dgData.Rows[e.RowIndex].Cells["E"].Value.ToString();
                txtCoordN.Text = dgData.Rows[e.RowIndex].Cells["N"].Value.ToString();
                txtCoordElevation.Text = dgData.Rows[e.RowIndex].Cells["Z"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["CS"].Value.ToString() == "")
                    cmbCS.SelectedValue = "-1";
                else cmbCS.SelectedValue = dgData.Rows[e.RowIndex].Cells["CS"].Value.ToString();
                
                txtGPSEPE.Text = dgData.Rows[e.RowIndex].Cells["GPSepe"].Value.ToString();
                txtPhoto.Text = dgData.Rows[e.RowIndex].Cells["Photo"].Value.ToString();
                txtPhotoAzimuth.Text = dgData.Rows[e.RowIndex].Cells["Photo_azimuth"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["SampleType"].Value.ToString() == "")
                    cmbSampleType.SelectedValue = "-1";
                else cmbSampleType.SelectedValue = dgData.Rows[e.RowIndex].Cells["SampleType"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["NotItSitu"].Value.ToString() == "")
                    cmbNotInSitu.SelectedValue = "-1";
                else cmbNotInSitu.SelectedValue = dgData.Rows[e.RowIndex].Cells["NotItSitu"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["Porpuose"].Value.ToString() == "")
                    cmbPorpuose.SelectedValue = "-1";
                else cmbPorpuose.SelectedValue = dgData.Rows[e.RowIndex].Cells["Porpuose"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["Relative_Loc"].Value.ToString() == "")
                    cmbRelativeLoc.SelectedValue = "-1";
                else cmbRelativeLoc.SelectedValue = dgData.Rows[e.RowIndex].Cells["Relative_Loc"].Value.ToString();
                
                txtLenght.Text = dgData.Rows[e.RowIndex].Cells["length"].Value.ToString();
                txtHigh.Text = dgData.Rows[e.RowIndex].Cells["High"].Value.ToString();
                txtThickness.Text = dgData.Rows[e.RowIndex].Cells["Thickness"].Value.ToString();
                txtObservations.Text = dgData.Rows[e.RowIndex].Cells["Obsevations"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LRock"].Value.ToString() == "")
                    cmbLithologyLit.SelectedValue = "-1";
                else cmbLithologyLit.SelectedValue = dgData.Rows[e.RowIndex].Cells["LRock"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LTexture"].Value.ToString() == "")
                    cmbLTextures.SelectedValue = "-1";
                else cmbLTextures.SelectedValue = dgData.Rows[e.RowIndex].Cells["LTexture"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LGSize"].Value.ToString() == "")
                    cmbLGsize.SelectedValue = "-1";
                else cmbLGsize.SelectedValue = dgData.Rows[e.RowIndex].Cells["LGSize"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LWeathering"].Value.ToString() == "")
                    cmbLWeathering.SelectedValue = "-1";
                else cmbLWeathering.SelectedValue = dgData.Rows[e.RowIndex].Cells["LWeathering"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LRocksSorting"].Value.ToString() == "")
                    cmbRSorting.SelectedValue = "-1";
                else cmbRSorting.SelectedValue = dgData.Rows[e.RowIndex].Cells["LRocksSorting"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LRocksSphericity"].Value.ToString() == "")
                    cmbRSphericity.SelectedValue = "-1";
                else cmbRSphericity.SelectedValue = dgData.Rows[e.RowIndex].Cells["LRocksSphericity"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LRocksRounding"].Value.ToString() == "")
                    cmbRounding.SelectedValue = "-1";
                else cmbRounding.SelectedValue = dgData.Rows[e.RowIndex].Cells["LRocksRounding"].Value.ToString();
                
                txtObservSedimentary.Text = dgData.Rows[e.RowIndex].Cells["LRocksObservation"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LMatrixPerc"].Value.ToString() == "")
                    cmbMatrixPorc.SelectedValue = "-1";
                else cmbMatrixPorc.SelectedValue = dgData.Rows[e.RowIndex].Cells["LMatrixPerc"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LMatrixGSize"].Value.ToString() == "")
                    cmbMatrixGSize.SelectedValue = "-1";
                else cmbMatrixGSize.SelectedValue = dgData.Rows[e.RowIndex].Cells["LMatrixGSize"].Value.ToString();
                
                txtMatrixObserv.Text = dgData.Rows[e.RowIndex].Cells["LMatrixObsevations"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LPhenoCPerc"].Value.ToString() == "")
                    cmbPhenoPerc.SelectedValue = "-1";
                else cmbPhenoPerc.SelectedValue = dgData.Rows[e.RowIndex].Cells["LPhenoCPerc"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["LPhenoCGSize"].Value.ToString() == "")
                    cmbPhenoGSize.SelectedValue = "-1";
                else cmbPhenoGSize.SelectedValue = dgData.Rows[e.RowIndex].Cells["LPhenoCGSize"].Value.ToString();
                
                txtPhenoObserv.Text = dgData.Rows[e.RowIndex].Cells["LPhenoCObsevations"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["VContactType"].Value.ToString() == "")
                    cmbContactType.SelectedValue = "-1";
                else
                    cmbContactType.SelectedValue = dgData.Rows[e.RowIndex].Cells["VContactType"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["VVeinName"].Value.ToString() == "")
                    cmbPhenoGSize.SelectedValue = "-1";
                else
                    cmbVeinName.SelectedValue = dgData.Rows[e.RowIndex].Cells["VVeinName"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["VHostRock"].Value.ToString() == "")
                    cmbHostRock.SelectedValue = "-1";
                else
                    cmbHostRock.SelectedValue = dgData.Rows[e.RowIndex].Cells["VHostRock"].Value.ToString();

                txtVeinObserv.Text = dgData.Rows[e.RowIndex].Cells["VObsevations"].Value.ToString();

                if (dgData.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString() == "")
                    cmbSamplingType.SelectedValue = "-1";
                else cmbSamplingType.SelectedValue = dgData.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString();

                txtDupOf.Text = dgData.Rows[e.RowIndex].Cells["DupOf"].Value.ToString();


                //sMineLoc
                if (dgData.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString() != "")
                {
                    if (dgData.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString().Substring(0, 1) == "O")
                    {
                        bMineLocation(false);
                        txtMineLocation.Text = dgData.Rows[e.RowIndex].Cells["Mine"].Value.ToString();
                    }
                    else
                    {
                        bMineLocation(true);
                        LoadCmbMineLocation();
                        cmbMineLocation.SelectedValue = dgData.Rows[e.RowIndex].Cells["Mine"].Value.ToString()
                            == ""
                            ? "Select an option..."
                            : dgData.Rows[e.RowIndex].Cells["Mine"].Value.ToString();
                    }
                }
                
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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

        private void TbRocks_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == (char)Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void txtCS_KeyPress(object sender, KeyPressEventArgs e)
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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                ControlsClean();
                dgData.DataSource = LoadDataRocksAll("2");//LoadDataRocks(txtSample.Text.ToString());
                dgData.Columns["SKSamplesRock"].Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ControlsClean()
        {
            try
            {
                oGCRock.iSKSamplesRock = 0;
                sEdit = "0";
                dtimeDate.Text = DateTime.Now.ToShortDateString();
                txtSample.Text = "";
                cmbTarget.SelectedValue = "-1";
                txtLocation.Text = "";
                txtProject.Text = ConfigurationSettings.AppSettings["IDProjectGC"].ToString();
                cmbGeologist.SelectedValue = "-1";
                txtHelper.Text = "";
                txtStation.Text = "";
                txtCoordE.Text = "";
                txtCoordN.Text =  "";
                txtCoordElevation.Text = "";
                cmbCS.SelectedValue= "-1";
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
                cmbMatrixPorc.SelectedValue = "-1";
                cmbMatrixGSize.SelectedValue = "-1";
                txtMatrixObserv.Text = "";
                cmbPhenoPerc.SelectedValue = "-1";
                cmbPhenoGSize.SelectedValue = "-1";
                txtPhenoObserv.Text = "";
                cmbContactType.SelectedValue = "-1";
                cmbVeinName.SelectedValue = "-1";
                cmbHostRock.SelectedValue = "-1";
                txtVeinObserv.Text = "";

                cmbMineLocation.SelectedValue = "";
                txtMineLocation.Text = "";
                cmbMineLocation.Visible = false;
                txtMineLocation.Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (MessageBox.Show("Row Delete. " + "Sample " + dgData.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Location " + dgData.Rows[e.RowIndex].Cells["Location"].Value.ToString()
                   + " Project " + dgData.Rows[e.RowIndex].Cells["Project"].Value.ToString()
                   + " Geologist " + dgData.Rows[e.RowIndex].Cells["Geologist"].Value.ToString()
                   , "Geochemistry Rocks", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oGCRock.iSKSamplesRock = int.Parse(dgData.Rows[e.RowIndex].Cells["SKSamplesRock"].Value.ToString());
                    string sRespDel = oGCRock.GCSamplesRock_Delete();
                    if (sRespDel == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Geochemistry", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dgData.DataSource = LoadDataRocksAll("2");//LoadDataRocks(txtSample.Text.ToString());
                        dgData.Columns["SKSamplesRock"].Visible = false;
                        dgLithology.DataSource = LoadDataRocksAll("2");//LoadDataRocks(txtSample.Text.ToString());
                        dgLithology.Columns["SKSamplesRock"].Visible = false;
                        ControlsClean();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtSample_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == (char)Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void LoadDataMinLith(string _sOpcion)
        {
            try
            {
                if (_sOpcion == "1")
                {
                    oGCRLith.sOpcion = _sOpcion;
                    oGCRLith.sSample = cmbSample.SelectedValue.ToString();
                    dgLithMatrix.DataSource = oGCRLith.getGCSamplesRockLithList();
                    dgLithMatrix.Columns["SKLithRock"].Visible = false;

                }
                else if (_sOpcion == "2")
                {
                    oGCRLith.sOpcion = _sOpcion;
                    oGCRLith.sSample = cmbSample.SelectedValue.ToString();
                    dgLithPheno.DataSource = oGCRLith.getGCSamplesRockLithList();
                    dgLithPheno.Columns["SKLithRock"].Visible = false;

                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddMinMat_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtSample.Text != "")
                {
                    if (sEditLithMat == "0")
                    {
                        oGCRLith.iSkLithRock = 0;
                        oGCRLith.sOpcion = "1";
                    }
                    else if (sEditLithMat == "1")
                    {
                        oGCRLith.sOpcion = "2";
                    }
                    
                    oGCRLith.sMineral = cmbMineralMt.SelectedValue.ToString();
                    oGCRLith.sSample = txtSample.Text.ToString();
                    oGCRLith.sType = "Mat";
                    string sResp = oGCRLith.GCSamplesRockLith_Add();
                    if (sResp == "OK")
                    {
                        cmbMineralMt.SelectedValue = "-1";
                        LoadDataMinLith("1");
                    }
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
                if (txtSample.Text != "")
                {
                    if (sEditLthPhe == "0")
                    {
                        oGCRLith.iSkLithRock = 0;
                        oGCRLith.sOpcion = "1";
                    }
                    else if (sEditLthPhe == "1")
                    {
                        oGCRLith.sOpcion = "2";
                    }
                   
                    oGCRLith.sMineral = cmbMineralPh.SelectedValue.ToString();
                    oGCRLith.sSample = txtSample.Text.ToString();
                    oGCRLith.sType = "Phe";
                    string sResp = oGCRLith.GCSamplesRockLith_Add();
                    if (sResp == "OK")
                    {
                        cmbMineralPh.SelectedValue = "-1";
                        LoadDataMinLith("2");
                    }
                }


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
                    oGCRLith.iSkLithRock = int.Parse(dgLithMatrix.Rows[e.RowIndex].Cells["SKLithRock"].Value.ToString());
                    oGCRLith.GCSamplesRockLith_Delete();
                    LoadDataMinLith("1");
                    sEditLithMat = "0";
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
                      , "Lithology PhenoCrystal", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                  MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oGCRLith.iSkLithRock = int.Parse(dgLithPheno.Rows[e.RowIndex].Cells["SKLithRock"].Value.ToString());
                    oGCRLith.GCSamplesRockLith_Delete();
                    LoadDataMinLith("2");
                    sEditLthPhe = "0";
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
                oGCRLith.iSkLithRock = int.Parse(dgLithMatrix.Rows[e.RowIndex].Cells["SKLithRock"].Value.ToString());

                if (dgLithMatrix.Rows[e.RowIndex].Cells["Mineral"].Value.ToString() == "")
                    cmbMineralMt.SelectedValue = "-1";
                else cmbMineralMt.SelectedValue = dgLithMatrix.Rows[e.RowIndex].Cells["Mineral"].Value.ToString();
                
                sEditLithMat = "1";
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
                oGCRLith.iSkLithRock = int.Parse(dgLithPheno.Rows[e.RowIndex].Cells["SKLithRock"].Value.ToString());

                if (dgLithPheno.Rows[e.RowIndex].Cells["Mineral"].Value.ToString() == "")
                    cmbMineralPh.SelectedValue = "-1";
                else cmbMineralPh.SelectedValue = dgLithPheno.Rows[e.RowIndex].Cells["Mineral"].Value.ToString();
                
                sEditLthPhe = "1";
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
                oAlt.sSample = cmbSample.SelectedValue.ToString();
                dgAlterations.DataSource = oAlt.getGCSamplesRockAltList();
                dgAlterations.Columns["SKSampleRAlt"].Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnAddAlt_Click(object sender, EventArgs e)
        {
            try
            {
                if (sEditAlt == "0")
                {
                    oAlt.iSkSampleRAlt = 0;
                    oAlt.sOpcion = "1";
                }
                else 
                {
                    oAlt.sOpcion = "2";
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


                string sResp = oAlt.GCSamplesRockAlt_Add();
                if (sResp == "OK")
                {
                    CleanControlsAlt();
                    sEditAlt = "0";
                    LoadDataAlterations("1");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgAlterations_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditAlt = "1";
                oAlt.iSkSampleRAlt = int.Parse(dgAlterations.Rows[e.RowIndex].Cells["SKSampleRAlt"].Value.ToString());

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
                    oAlt.iSkSampleRAlt = int.Parse(dgAlterations.Rows[e.RowIndex].Cells["SKSampleRAlt"].Value.ToString());
                    oAlt.GCSamplesRockAlt_Delete();
                    LoadDataAlterations("1");
                    sEditAlt = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private void LoadDataMineralizations(string _sOpcion)
        {
            try
            {

                oMin.sOpcion = _sOpcion;
                oMin.sSample = cmbSample.SelectedValue.ToString();
                dgSamplesRockMin.DataSource = oMin.getGCSamplesRockMinList();
                dgSamplesRockMin.Columns["SKSampleRMin"].Visible = false;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
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
                    oMin.iSKSampleRMin = 0;
                }
                else
                {
                    oMin.sOpcion = "2";
                }
                
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

                if (txtMinPerc.Text == "")
                    oMin.dMinPerc = null;
                else
                    oMin.dMinPerc = double.Parse(txtMinPerc.Text.ToString());

                if (txtObservMin.Text.ToString() == "")
                    oMin.sObservations = null;
                else
                    oMin.sObservations = txtObservMin.Text.ToString();

                string sResp = oMin.GCSamplesRockMin_Add();
                if (sResp == "OK")
                {
                    CleanControlsMin();
                    sEditMin = "0";
                    LoadDataMineralizations("1");
                }
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
                CleanControlsMin();
                sEditMin = "0";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgSamplesRockMin_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditMin = "1";
                oMin.iSKSampleRMin = int.Parse(dgSamplesRockMin.Rows[e.RowIndex].Cells["SKSampleRMin"].Value.ToString());

                if (dgSamplesRockMin.Rows[e.RowIndex].Cells["MZMin"].Value.ToString() == "")
                    cmbMineralmin.SelectedValue = "-1";
                else cmbMineralmin.SelectedValue = dgSamplesRockMin.Rows[e.RowIndex].Cells["MZMin"].Value.ToString();

                if (dgSamplesRockMin.Rows[e.RowIndex].Cells["MZStyle"].Value.ToString() == "")
                    cmbStyleM.SelectedValue = "-1";
                else cmbStyleM.SelectedValue = dgSamplesRockMin.Rows[e.RowIndex].Cells["MZStyle"].Value.ToString();

                if (dgSamplesRockMin.Rows[e.RowIndex].Cells["MZPerc"].Value.ToString() == "")
                    txtMinPerc.Text = "";
                else txtMinPerc.Text = dgSamplesRockMin.Rows[e.RowIndex].Cells["MZPerc"].Value.ToString();
                
                txtObservMin.Text = dgSamplesRockMin.Rows[e.RowIndex].Cells["Obsevations"].Value.ToString();
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
                if (MessageBox.Show("Row Delete. " + "Sample " + dgSamplesRockMin.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Mineral " + dgSamplesRockMin.Rows[e.RowIndex].Cells["MZMin"].Value.ToString()
                   + " Style " + dgSamplesRockMin.Rows[e.RowIndex].Cells["MZStyle"].Value.ToString()
                   , "Alterations ", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oMin.iSKSampleRMin = int.Parse(dgSamplesRockMin.Rows[e.RowIndex].Cells["SKSampleRMin"].Value.ToString());
                    oMin.GCSamplesRockMin_Delete();
                    LoadDataMineralizations("1");
                    sEditMin = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private void LoadDataOxides(string _sOpcion)
        {
            try
            {

                oOxid.sOpcion = _sOpcion;
                oOxid.sSample = cmbSample.SelectedValue.ToString();
                dgSamplesRockOxides.DataSource = oOxid.getGCSamplesRockOxidesList();
                dgSamplesRockOxides.Columns["SKSamplesRockOxides"].Visible = false;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnAddOxides_Click(object sender, EventArgs e)
        {
            try
            {
                if (sEditOxid == "0")
                {
                    oOxid.sOpcion = "1";
                    oOxid.iSKSamplesRockOxides = 0;
                }
                else
                {
                    oOxid.sOpcion = "2";
                }

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

                string sResp = oOxid.GCSamplesRockOxides_Add();
                if (sResp == "OK")
                {
                    sEditOxid = "0";
                    CleanControlsOxides();
                    LoadDataOxides("1");
                }

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

        private void dgSamplesRockOxides_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditOxid = "1";
                oOxid.iSKSamplesRockOxides = int.Parse(dgSamplesRockOxides.Rows[e.RowIndex].Cells["SKSamplesRockOxides"].Value.ToString());

                if (dgSamplesRockOxides.Rows[e.RowIndex].Cells["GoeStyle"].Value.ToString() == "")
                    cmbStyleGoe.SelectedValue = "-1";
                else cmbStyleGoe.SelectedValue = dgSamplesRockOxides.Rows[e.RowIndex].Cells["GoeStyle"].Value.ToString();

                if (dgSamplesRockOxides.Rows[e.RowIndex].Cells["GoePerc"].Value.ToString() == "")
                    cmbPercGoe.SelectedValue = "-1";
                else cmbPercGoe.SelectedValue = dgSamplesRockOxides.Rows[e.RowIndex].Cells["GoePerc"].Value.ToString();

                if (dgSamplesRockOxides.Rows[e.RowIndex].Cells["HemStyle"].Value.ToString() == "")
                    cmbStyleHem.SelectedValue = "-1";
                else cmbStyleHem.SelectedValue = dgSamplesRockOxides.Rows[e.RowIndex].Cells["HemStyle"].Value.ToString();

                if (dgSamplesRockOxides.Rows[e.RowIndex].Cells["HemPerc"].Value.ToString() == "")
                    cmbPercHem.SelectedValue = "-1";
                else cmbPercHem.SelectedValue = dgSamplesRockOxides.Rows[e.RowIndex].Cells["HemPerc"].Value.ToString();

                if (dgSamplesRockOxides.Rows[e.RowIndex].Cells["JarStyle"].Value.ToString() == "")
                    cmbStyleJar.SelectedValue = "-1";
                else cmbStyleJar.SelectedValue = dgSamplesRockOxides.Rows[e.RowIndex].Cells["JarStyle"].Value.ToString();

                if (dgSamplesRockOxides.Rows[e.RowIndex].Cells["JarPerc"].Value.ToString() == "")
                    cmbPercJar.SelectedValue = "-1";
                else cmbPercJar.SelectedValue = dgSamplesRockOxides.Rows[e.RowIndex].Cells["JarPerc"].Value.ToString();

                if (dgSamplesRockOxides.Rows[e.RowIndex].Cells["LimStyle"].Value.ToString() == "")
                    cmbStyleLim.SelectedValue = "-1";
                else cmbStyleLim.SelectedValue = dgSamplesRockOxides.Rows[e.RowIndex].Cells["LimStyle"].Value.ToString();

                if (dgSamplesRockOxides.Rows[e.RowIndex].Cells["LimPerc"].Value.ToString() == "")
                    cmbPercLim.SelectedValue = "-1";
                else cmbPercLim.SelectedValue = dgSamplesRockOxides.Rows[e.RowIndex].Cells["LimPerc"].Value.ToString();
                
                txtObservOxides.Text = dgSamplesRockOxides.Rows[e.RowIndex].Cells["Observations"].Value.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgSamplesRockOxides_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgSamplesRockOxides.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                   + " Goethite Style " + dgSamplesRockOxides.Rows[e.RowIndex].Cells["GoeStyle"].Value.ToString()
                   + " Hematite Style " + dgSamplesRockOxides.Rows[e.RowIndex].Cells["HemStyle"].Value.ToString()
                   , " Oxides ", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oOxid.iSKSamplesRockOxides = int.Parse(dgSamplesRockOxides.Rows[e.RowIndex].Cells["SKSamplesRockOxides"].Value.ToString());
                    oOxid.GCSamplesRockOxides_Delete();
                    LoadDataOxides("1");
                    sEditOxid = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox42_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtNumberSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtDipStr_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
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

        private void btnAddStr_Click(object sender, EventArgs e)
        {
            try
            {
                if (sEditStr == "0")
                {
                    oStr.sOpcion = "1";
                    oStr.iSKSamplesRockStr = 0;
                }
                else
                {
                    oStr.sOpcion = "2";
                }
                

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


                if (cmbFillSt.SelectedValue != null)
                {
                    if (cmbFillSt.SelectedValue.ToString() == "-1" ||
                        cmbFillSt.SelectedValue.ToString() == "")
                        oStr.sFill = null;
                    else
                        oStr.sFill = cmbFillSt.SelectedValue.ToString();
                }
                else
                { oStr.sFill = null; }
                

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


                if (cmbFillSt2.SelectedValue != null)
                {
                    if (cmbFillSt2.SelectedValue.ToString() == "-1" ||
                        cmbFillSt2.SelectedValue.ToString() == "")
                        oStr.sFill2 = null;
                    else
                        oStr.sFill2 = cmbFillSt2.SelectedValue.ToString();
                }
                else
                { oStr.sFill2 = null; }
                

                if (cmbFillSt3.SelectedValue != null)
                {
                    if (cmbFillSt3.SelectedValue.ToString() == "-1" ||
                        cmbFillSt3.SelectedValue.ToString() == "")
                        oStr.sFill3 = null;
                    else
                        oStr.sFill3 = cmbFillSt3.SelectedValue.ToString();
                }
                else
                { oStr.sFill3 = null; }
                

                string sResp = oStr.GCSamplesRockStr_Add();
                if (sResp == "OK")
                {
                    sEditStr = "0";
                    LoadDataStructures("1");
                    CleanControlsStr();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadDataStructures(string _sOpcion)
        {
            try
            {
                oStr.sOpcion = _sOpcion;
                oStr.sSample = cmbSample.SelectedValue.ToString();
                dgSamplesRockStr.DataSource = oStr.getGCSamplesRockStrList();
                dgSamplesRockStr.Columns["SKSamplesRockStr"].Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgSamplesRockStr_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditStr = "1";
                oStr.iSKSamplesRockStr = int.Parse(dgSamplesRockStr.Rows[e.RowIndex].Cells["SKSamplesRockStr"].Value.ToString());

                if (dgSamplesRockStr.Rows[e.RowIndex].Cells["StrType"].Value.ToString() == "")
                    cmbStructureTypeSt.SelectedValue = "-1";
                else cmbStructureTypeSt.SelectedValue = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrType"].Value.ToString();
                
                txtDipStr.Text = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrDip"].Value.ToString();
                txtDipAzStr.Text = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrDipAz"].Value.ToString();
                txtAppThickSt.Text = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrAThick"].Value.ToString();
                txtRThickStr.Text = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrRThick"].Value.ToString();

                if (dgSamplesRockStr.Rows[e.RowIndex].Cells["StrFill"].Value.ToString() == "")
                    cmbFillSt.SelectedValue = "-1";
                else cmbFillSt.SelectedValue = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrFill"].Value.ToString();

                if (dgSamplesRockStr.Rows[e.RowIndex].Cells["StrFill2"].Value.ToString() == "")
                    cmbFillSt2.SelectedValue = "-1";
                else cmbFillSt2.SelectedValue = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrFill2"].Value.ToString();

                if (dgSamplesRockStr.Rows[e.RowIndex].Cells["StrFill3"].Value.ToString() == "")
                    cmbFillSt3.SelectedValue = "-1";
                else cmbFillSt3.SelectedValue = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrFill3"].Value.ToString();

                txtNumberSt.Text = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrNumber"].Value.ToString();
                txtDensityStr.Text = dgSamplesRockStr.Rows[e.RowIndex].Cells["StrDensity"].Value.ToString();
                txtObservStr.Text = dgSamplesRockStr.Rows[e.RowIndex].Cells["Obsevations"].Value.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgSamplesRockStr_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Sample " + dgSamplesRockStr.Rows[e.RowIndex].Cells["Sample"].Value.ToString()
                      + " Dip " + dgSamplesRockStr.Rows[e.RowIndex].Cells["StrDip"].Value.ToString()
                      + " Fill " + dgSamplesRockStr.Rows[e.RowIndex].Cells["StrFill"].Value.ToString()
                      , "Structures", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                  MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oStr.iSKSamplesRockStr = int.Parse(dgSamplesRockStr.Rows[e.RowIndex].Cells["SKSamplesRockStr"].Value.ToString());
                    oStr.GCSamplesRockStr_Delete();
                    LoadDataStructures("1");
                    sEditStr = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtAppThickSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtRThickStr_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
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

        private void txtDupOf_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void btnCancelMinMat_Click(object sender, EventArgs e)
        {
            try
            {
                sEditLithMat = "0";
                cmbMineralMt.SelectedValue = "-1";
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
                sEditLthPhe = "0";
                cmbMineralPh.SelectedValue = "-1";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 
        }

        private void btnCancelStr_Click(object sender, EventArgs e)
        {
            CleanControlsStr();
        }

        private void cmbSample_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dgData.DataSource = LoadDataRocksAll("2"); //LoadDataRocks(txtSample.Text.ToString());
                dgData.Columns["SKSamplesRock"].Visible = false;
                
               
                LoadDataMinLith("1"); LoadDataMinLith("2"); LoadDataAlterations("1");
                LoadDataMineralizations("1"); LoadDataOxides("1"); LoadDataStructures("1");


                DataTable dtData = LoadDataRocksAll("2");
                // Query the SalesOrderHeader table for orders placed 
                // after August 8, 2001.
                IEnumerable<DataRow> query =
                    from dtDat in dtData.AsEnumerable()
                    where dtDat.Field<String>("Sample") == cmbSample.SelectedValue.ToString()
                    select dtDat;

                DataTable boundTable = new DataTable();
                if (query.Count() > 0)
                {
                    // Create a table from the query.
                    boundTable = query.CopyToDataTable<DataRow>();
                    dgLithology.DataSource = boundTable;//LoadDataRocks(txtSample.Text.ToString());
                    dgLithology.Columns["SKSamplesRock"].Visible = false;
                }
                else
                {
                    boundTable = null;
                    dgLithology.DataSource = boundTable;//LoadDataRocks(txtSample.Text.ToString());
                }

                cmbSample_Leave(null, null);
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
                frmRock ofrm = new frmRock();
                ofrm.MdiParent = this.MdiParent;
                ofrm.Show();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Para activar o desactivar los campos de mine location, combo o text.
        /// True activa el combo sMineLoc = 1, False activa el text sMineLoc = 2.  
        /// </summary>
        /// <param name="_bActive"></param>
        private void bMineLocation(bool _bActive)
        {
            try
            {
                if (_bActive == true)
                {
                    cmbMineLocation.Visible = true;
                    txtMineLocation.Visible = false;
                    txtMineLocation.Text = "";
                    sMineLoc = "1";
                }
                else
                {
                    txtMineLocation.Visible = true;
                    cmbMineLocation.Visible = false;
                    cmbMineLocation.SelectedValue = "-1";
                    sMineLoc = "2";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cmbSamplingType_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                if (cmbSamplingType.SelectedValue.ToString().Substring(0, 1) == "O")
                {
                    bMineLocation(false);   
                }
                else
                {
                    bMineLocation(true);
                    LoadCmbMineLocation();
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadCmbMineLocation()
        {
            try
            {
                DataTable dtMineEnt = new DataTable();
                dtMineEnt = oRf.getMineEntrance();
                DataRow drMineEnt = dtMineEnt.NewRow();
                drMineEnt[0] = "-1";
                drMineEnt[1] = "Select an option...";
                dtMineEnt.Rows.Add(drMineEnt);
                cmbMineLocation.DisplayMember = "cmb";
                cmbMineLocation.ValueMember = "cmb";
                cmbMineLocation.DataSource = dtMineEnt;
                cmbMineLocation.SelectedValue = "Select an option...";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgLithology_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                oGCRock.iSKSamplesRock = int.Parse(dgLithology.Rows[e.RowIndex].Cells["SKSamplesRock"].Value.ToString());
                sEdit = "1";

                DateTime dDate =
                    dgLithology.Rows[e.RowIndex].Cells["Date"].Value.ToString() == ""
                    ? DateTime.Parse("1900/01/01")
                    : DateTime.Parse(dgLithology.Rows[e.RowIndex].Cells["Date"].Value.ToString());
                dtimeDate.Value = dDate;
                dtimeDate.Text = dgLithology.Rows[e.RowIndex].Cells["Date"].Value.ToString();

                txtSample.Text = dgLithology.Rows[e.RowIndex].Cells["Sample"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["Target"].Value.ToString() == "")
                    cmbTarget.SelectedValue = "-1";
                else cmbTarget.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["Target"].Value.ToString();

                txtLocation.Text = dgLithology.Rows[e.RowIndex].Cells["Location"].Value.ToString();
                txtProject.Text = dgLithology.Rows[e.RowIndex].Cells["Project"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["Geologist"].Value.ToString() == "")
                    cmbGeologist.SelectedValue = "-1";
                else cmbGeologist.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["Geologist"].Value.ToString();

                txtHelper.Text = dgLithology.Rows[e.RowIndex].Cells["Helper"].Value.ToString();
                txtStation.Text = dgLithology.Rows[e.RowIndex].Cells["Station"].Value.ToString();
                txtCoordE.Text = dgLithology.Rows[e.RowIndex].Cells["E"].Value.ToString();
                txtCoordN.Text = dgLithology.Rows[e.RowIndex].Cells["N"].Value.ToString();
                txtCoordElevation.Text = dgLithology.Rows[e.RowIndex].Cells["Z"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["CS"].Value.ToString() == "")
                    cmbCS.SelectedValue = "-1";
                else cmbCS.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["CS"].Value.ToString();

                txtGPSEPE.Text = dgLithology.Rows[e.RowIndex].Cells["GPSepe"].Value.ToString();
                txtPhoto.Text = dgLithology.Rows[e.RowIndex].Cells["Photo"].Value.ToString();
                txtPhotoAzimuth.Text = dgLithology.Rows[e.RowIndex].Cells["Photo_azimuth"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["SampleType"].Value.ToString() == "")
                    cmbSampleType.SelectedValue = "-1";
                else cmbSampleType.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["SampleType"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["NotItSitu"].Value.ToString() == "")
                    cmbNotInSitu.SelectedValue = "-1";
                else cmbNotInSitu.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["NotItSitu"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["Porpuose"].Value.ToString() == "")
                    cmbPorpuose.SelectedValue = "-1";
                else cmbPorpuose.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["Porpuose"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["Relative_Loc"].Value.ToString() == "")
                    cmbRelativeLoc.SelectedValue = "-1";
                else cmbRelativeLoc.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["Relative_Loc"].Value.ToString();

                txtLenght.Text = dgLithology.Rows[e.RowIndex].Cells["length"].Value.ToString();
                txtHigh.Text = dgLithology.Rows[e.RowIndex].Cells["High"].Value.ToString();
                txtThickness.Text = dgLithology.Rows[e.RowIndex].Cells["Thickness"].Value.ToString();
                txtObservations.Text = dgLithology.Rows[e.RowIndex].Cells["Obsevations"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LRock"].Value.ToString() == "")
                    cmbLithologyLit.SelectedValue = "-1";
                else cmbLithologyLit.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LRock"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LTexture"].Value.ToString() == "")
                    cmbLTextures.SelectedValue = "-1";
                else cmbLTextures.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LTexture"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LGSize"].Value.ToString() == "")
                    cmbLGsize.SelectedValue = "-1";
                else cmbLGsize.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LGSize"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LWeathering"].Value.ToString() == "")
                    cmbLWeathering.SelectedValue = "-1";
                else cmbLWeathering.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LWeathering"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LRocksSorting"].Value.ToString() == "")
                    cmbRSorting.SelectedValue = "-1";
                else cmbRSorting.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LRocksSorting"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LRocksSphericity"].Value.ToString() == "")
                    cmbRSphericity.SelectedValue = "-1";
                else cmbRSphericity.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LRocksSphericity"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LRocksRounding"].Value.ToString() == "")
                    cmbRounding.SelectedValue = "-1";
                else cmbRounding.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LRocksRounding"].Value.ToString();

                txtObservSedimentary.Text = dgLithology.Rows[e.RowIndex].Cells["LRocksObservation"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LMatrixPerc"].Value.ToString() == "")
                    cmbMatrixPorc.SelectedValue = "-1";
                else cmbMatrixPorc.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LMatrixPerc"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LMatrixGSize"].Value.ToString() == "")
                    cmbMatrixGSize.SelectedValue = "-1";
                else cmbMatrixGSize.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LMatrixGSize"].Value.ToString();

                txtMatrixObserv.Text = dgLithology.Rows[e.RowIndex].Cells["LMatrixObsevations"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LPhenoCPerc"].Value.ToString() == "")
                    cmbPhenoPerc.SelectedValue = "-1";
                else cmbPhenoPerc.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LPhenoCPerc"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["LPhenoCGSize"].Value.ToString() == "")
                    cmbPhenoGSize.SelectedValue = "-1";
                else cmbPhenoGSize.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["LPhenoCGSize"].Value.ToString();

                txtPhenoObserv.Text = dgLithology.Rows[e.RowIndex].Cells["LPhenoCObsevations"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["VContactType"].Value.ToString() == "")
                    cmbContactType.SelectedValue = "-1";
                else
                    cmbContactType.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["VContactType"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["VVeinName"].Value.ToString() == "")
                    cmbPhenoGSize.SelectedValue = "-1";
                else
                    cmbVeinName.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["VVeinName"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["VHostRock"].Value.ToString() == "")
                    cmbHostRock.SelectedValue = "-1";
                else
                    cmbHostRock.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["VHostRock"].Value.ToString();

                txtVeinObserv.Text = dgLithology.Rows[e.RowIndex].Cells["VObsevations"].Value.ToString();

                if (dgLithology.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString() == "")
                    cmbSamplingType.SelectedValue = "-1";
                else cmbSamplingType.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString();

                txtDupOf.Text = dgLithology.Rows[e.RowIndex].Cells["DupOf"].Value.ToString();


                //sMineLoc
                if (dgLithology.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString() != "")
                {
                    if (dgLithology.Rows[e.RowIndex].Cells["SamplingType"].Value.ToString().Substring(0, 1) == "O")
                    {
                        bMineLocation(false);
                        txtMineLocation.Text = dgLithology.Rows[e.RowIndex].Cells["Mine"].Value.ToString();
                    }
                    else
                    {
                        bMineLocation(true);
                        LoadCmbMineLocation();
                        cmbMineLocation.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["Mine"].Value.ToString()
                            == ""
                            ? "Select an option..."
                            : dgLithology.Rows[e.RowIndex].Cells["Mine"].Value.ToString();
                    }
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnExporExcelAll_Click(object sender, EventArgs e)
        {
            try
            {
                
                if (cmbSample.SelectedValue == null)
                {
                    MessageBox.Show("Select Samples");
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
                    oGCRLith.sOpcion = _sOpcion;
                    oGCRLith.sSample = cmbSample.SelectedValue.ToString();
                    dtResp = oGCRLith.getGCSamplesRockLithList();

                }
                else if (_sOpcion == "2")
                {
                    //Phenocryst
                    oGCRLith.sOpcion = _sOpcion;
                    oGCRLith.sSample = cmbSample.SelectedValue.ToString();
                    dtResp = oGCRLith.getGCSamplesRockLithList();
                }

                return dtResp;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void ExpGeochemistry()
        {
            try
            {

                DataTable dtSample = new DataTable();
                oGCRock.sSample = cmbSample.SelectedValue.ToString();
                dtSample = oGCRock.getGCSamplesRockListBySampleReport();


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

                if (dtSample.Rows.Count > 0)
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
                    oSheet.Cells[9, 12] = dtSample.Rows[0]["E"].ToString();
                    oSheet.Cells[9, 26] = dtSample.Rows[0]["N"].ToString();
                    oSheet.Cells[9, 44] = dtSample.Rows[0]["Z"].ToString();


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

                    //oSheet.Cells[7, 71] = dtSample.Rows[0]["From"].ToString();
                    //oSheet.Cells[7, 75] = dtSample.Rows[0]["To"].ToString();

                    oSheet.Cells[51, 40] = dtSample.Rows[0]["VContactType"].ToString();
                    oSheet.Cells[53, 40] = dtSample.Rows[0]["VVeinName"].ToString();
                    oSheet.Cells[55, 40] = dtSample.Rows[0]["VHostRock"].ToString();
                    oSheet.Cells[57, 42] = dtSample.Rows[0]["VObsevations"].ToString();

                    //oSheet.Cells[7, 55] = dtSample.Rows[0]["chId"].ToString();

                    oSheet.Cells[38, 62] = dtSample.Rows[0]["SampleType"].ToString();

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
                oAlt.sSample = cmbSample.SelectedValue.ToString();
                dtAlterations = oAlt.getGCSamplesRockAltListReport();

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
                oMin.sSample = cmbSample.SelectedValue.ToString();
                dtMineralizations = oMin.getGCSamplesRockMinListReport();

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
                oOxid.sSample = cmbSample.SelectedValue.ToString();
                dtOxides = oOxid.getGCSamplesRockOxidesListReport();
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
                oStr.sSample = cmbSample.SelectedValue.ToString();
                dtStructures = oStr.getGCSamplesRockStrListReport();

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

                if (dtSample.Rows.Count > 0)
                {
                    oSheet.Cells[36, 56] = dtSample.Rows[0]["Mine"].ToString();
                    ////oSheet.Cells[38, 62] = dtSample.Rows[0]["MineEntrance"].ToString();    
                }


                oXL.Visible = true;
                oXL.UserControl = true;


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


        /// <summary>
        /// Modificación Alvaro Araujo
        /// 03/07/2019
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void cmbSample_Leave(object sender, EventArgs e)
        {
            if (swCargado)
            {
                LoadDataRocks(cmbSample.Text);

                foreach (DataGridViewRow Row in dgData.Rows)
                {
                    var strFila = Row.Index;
                    string Valor = Convert.ToString(Row.Cells[1].Value);

                    if (Valor == cmbSample.Text.ToUpper())
                    {
                        BuscarChannerPorId(strFila);
                    }
                }
            }
        }

        private void BuscarChannerPorId(int strFila)
        {
            try
            {
                oGCRock.iSKSamplesRock = int.Parse(dgData.Rows[strFila].Cells["SKSamplesRock"].Value.ToString());
                sEdit = "1";

                DateTime dDate =
                    dgData.Rows[strFila].Cells["Date"].Value.ToString() == ""
                    ? DateTime.Parse("1900/01/01")
                    : DateTime.Parse(dgData.Rows[strFila].Cells["Date"].Value.ToString());
                dtimeDate.Value = dDate;
                dtimeDate.Text = dgData.Rows[strFila].Cells["Date"].Value.ToString();

                txtSample.Text = dgData.Rows[strFila].Cells["Sample"].Value.ToString();

                if (dgData.Rows[strFila].Cells["Target"].Value.ToString() == "")
                    cmbTarget.SelectedValue = "-1";
                else cmbTarget.SelectedValue = dgData.Rows[strFila].Cells["Target"].Value.ToString();

                txtLocation.Text = dgData.Rows[strFila].Cells["Location"].Value.ToString();
                txtProject.Text = dgData.Rows[strFila].Cells["Project"].Value.ToString();

                if (dgData.Rows[strFila].Cells["Geologist"].Value.ToString() == "")
                    cmbGeologist.SelectedValue = "-1";
                else cmbGeologist.SelectedValue = dgData.Rows[strFila].Cells["Geologist"].Value.ToString();

                txtHelper.Text = dgData.Rows[strFila].Cells["Helper"].Value.ToString();
                txtStation.Text = dgData.Rows[strFila].Cells["Station"].Value.ToString();
                txtCoordE.Text = dgData.Rows[strFila].Cells["E"].Value.ToString();
                txtCoordN.Text = dgData.Rows[strFila].Cells["N"].Value.ToString();
                txtCoordElevation.Text = dgData.Rows[strFila].Cells["Z"].Value.ToString();

                if (dgData.Rows[strFila].Cells["CS"].Value.ToString() == "")
                    cmbCS.SelectedValue = "-1";
                else cmbCS.SelectedValue = dgData.Rows[strFila].Cells["CS"].Value.ToString();

                txtGPSEPE.Text = dgData.Rows[strFila].Cells["GPSepe"].Value.ToString();
                txtPhoto.Text = dgData.Rows[strFila].Cells["Photo"].Value.ToString();
                txtPhotoAzimuth.Text = dgData.Rows[strFila].Cells["Photo_azimuth"].Value.ToString();

                if (dgData.Rows[strFila].Cells["SampleType"].Value.ToString() == "")
                    cmbSampleType.SelectedValue = "-1";
                else cmbSampleType.SelectedValue = dgData.Rows[strFila].Cells["SampleType"].Value.ToString();

                if (dgData.Rows[strFila].Cells["NotItSitu"].Value.ToString() == "")
                    cmbNotInSitu.SelectedValue = "-1";
                else cmbNotInSitu.SelectedValue = dgData.Rows[strFila].Cells["NotItSitu"].Value.ToString();

                if (dgData.Rows[strFila].Cells["Porpuose"].Value.ToString() == "")
                    cmbPorpuose.SelectedValue = "-1";
                else cmbPorpuose.SelectedValue = dgData.Rows[strFila].Cells["Porpuose"].Value.ToString();

                if (dgData.Rows[strFila].Cells["Relative_Loc"].Value.ToString() == "")
                    cmbRelativeLoc.SelectedValue = "-1";
                else cmbRelativeLoc.SelectedValue = dgData.Rows[strFila].Cells["Relative_Loc"].Value.ToString();

                txtLenght.Text = dgData.Rows[strFila].Cells["length"].Value.ToString();
                txtHigh.Text = dgData.Rows[strFila].Cells["High"].Value.ToString();
                txtThickness.Text = dgData.Rows[strFila].Cells["Thickness"].Value.ToString();
                txtObservations.Text = dgData.Rows[strFila].Cells["Obsevations"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LRock"].Value.ToString() == "")
                    cmbLithologyLit.SelectedValue = "-1";
                else cmbLithologyLit.SelectedValue = dgData.Rows[strFila].Cells["LRock"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LTexture"].Value.ToString() == "")
                    cmbLTextures.SelectedValue = "-1";
                else cmbLTextures.SelectedValue = dgData.Rows[strFila].Cells["LTexture"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LGSize"].Value.ToString() == "")
                    cmbLGsize.SelectedValue = "-1";
                else cmbLGsize.SelectedValue = dgData.Rows[strFila].Cells["LGSize"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LWeathering"].Value.ToString() == "")
                    cmbLWeathering.SelectedValue = "-1";
                else cmbLWeathering.SelectedValue = dgData.Rows[strFila].Cells["LWeathering"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LRocksSorting"].Value.ToString() == "")
                    cmbRSorting.SelectedValue = "-1";
                else cmbRSorting.SelectedValue = dgData.Rows[strFila].Cells["LRocksSorting"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LRocksSphericity"].Value.ToString() == "")
                    cmbRSphericity.SelectedValue = "-1";
                else cmbRSphericity.SelectedValue = dgData.Rows[strFila].Cells["LRocksSphericity"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LRocksRounding"].Value.ToString() == "")
                    cmbRounding.SelectedValue = "-1";
                else cmbRounding.SelectedValue = dgData.Rows[strFila].Cells["LRocksRounding"].Value.ToString();

                txtObservSedimentary.Text = dgData.Rows[strFila].Cells["LRocksObservation"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LMatrixPerc"].Value.ToString() == "")
                    cmbMatrixPorc.SelectedValue = "-1";
                else cmbMatrixPorc.SelectedValue = dgData.Rows[strFila].Cells["LMatrixPerc"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LMatrixGSize"].Value.ToString() == "")
                    cmbMatrixGSize.SelectedValue = "-1";
                else cmbMatrixGSize.SelectedValue = dgData.Rows[strFila].Cells["LMatrixGSize"].Value.ToString();

                txtMatrixObserv.Text = dgData.Rows[strFila].Cells["LMatrixObsevations"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LPhenoCPerc"].Value.ToString() == "")
                    cmbPhenoPerc.SelectedValue = "-1";
                else cmbPhenoPerc.SelectedValue = dgData.Rows[strFila].Cells["LPhenoCPerc"].Value.ToString();

                if (dgData.Rows[strFila].Cells["LPhenoCGSize"].Value.ToString() == "")
                    cmbPhenoGSize.SelectedValue = "-1";
                else cmbPhenoGSize.SelectedValue = dgData.Rows[strFila].Cells["LPhenoCGSize"].Value.ToString();

                txtPhenoObserv.Text = dgData.Rows[strFila].Cells["LPhenoCObsevations"].Value.ToString();

                if (dgData.Rows[strFila].Cells["VContactType"].Value.ToString() == "")
                    cmbContactType.SelectedValue = "-1";
                else
                    cmbContactType.SelectedValue = dgData.Rows[strFila].Cells["VContactType"].Value.ToString();

                if (dgData.Rows[strFila].Cells["VVeinName"].Value.ToString() == "")
                    cmbPhenoGSize.SelectedValue = "-1";
                else
                    cmbVeinName.SelectedValue = dgData.Rows[strFila].Cells["VVeinName"].Value.ToString();

                if (dgData.Rows[strFila].Cells["VHostRock"].Value.ToString() == "")
                    cmbHostRock.SelectedValue = "-1";
                else
                    cmbHostRock.SelectedValue = dgData.Rows[strFila].Cells["VHostRock"].Value.ToString();

                txtVeinObserv.Text = dgData.Rows[strFila].Cells["VObsevations"].Value.ToString();

                if (dgData.Rows[strFila].Cells["SamplingType"].Value.ToString() == "")
                    cmbSamplingType.SelectedValue = "-1";
                else cmbSamplingType.SelectedValue = dgData.Rows[strFila].Cells["SamplingType"].Value.ToString();

                txtDupOf.Text = dgData.Rows[strFila].Cells["DupOf"].Value.ToString();


                //sMineLoc
                if (dgData.Rows[strFila].Cells["SamplingType"].Value.ToString() != "")
                {
                    if (dgData.Rows[strFila].Cells["SamplingType"].Value.ToString().Substring(0, 1) == "O")
                    {
                        bMineLocation(false);
                        txtMineLocation.Text = dgData.Rows[strFila].Cells["Mine"].Value.ToString();
                    }
                    else
                    {
                        bMineLocation(true);
                        LoadCmbMineLocation();
                        cmbMineLocation.SelectedValue = dgData.Rows[strFila].Cells["Mine"].Value.ToString()
                            == ""
                            ? "Select an option..."
                            : dgData.Rows[strFila].Cells["Mine"].Value.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
