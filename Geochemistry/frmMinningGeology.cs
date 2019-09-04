using Geochemistry.Emun;
using System;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Geochemistry
{
    public partial class frmMinningGeology : Form
    {
        private clsRf oRf = new clsRf();
        private clsCHChannels oCh = new clsCHChannels();
        private clsCHSamples oCHSamp = new clsCHSamples();
        private bool swCombo = false;
        private int sampleExtradido = 0;
        private int cantMuestras = 0;
        private int conteoMuestras = 0;
        private double sFrom = 0;
        private double sTo = 0;
        private string sSampleSelect = "";
        private bool swActualizarRegistro = false;
        private int indexRegistroGrid = 0;

        private string minaSeleccionada = string.Empty;
        private string geologoSeleccionado = string.Empty;
        private string tipoCanalSeleccionado = string.Empty;
        private string sampleSeleccionado = string.Empty;
        private string sampleFaltante = string.Empty;

        private string channelPrimero = string.Empty;
        private double chLength = 0;
        private string sEditCh = "0";
        private int wSKCHChannels = 0;
        private int valorFinalMasUno = 0;

        public frmMinningGeology()
        {
            InitializeComponent();
        }

        #region Validadores
        private void ValidarControles(GroupBox groupbox)
        {
            foreach (Control control in groupbox.Controls)
            {
                if (control.GetType().Equals(typeof(TextBox)))
                {
                    if (control.Text == string.Empty)
                    {
                        MessageBox.Show(string.Concat("The Field ", control.Tag, " it is obligatory."), "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                else if (control.GetType().Equals(typeof(ComboBox)))
                {
                    if (control.Text == string.Empty)
                    {
                        MessageBox.Show(string.Concat("The Field ", control.Tag, " it is obligatory."), "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
            }
        }

        private bool ValidarValores()
        {
            if (Convert.ToInt32(cmbMineEntrance.SelectedValue) == -1)
            {
                MessageBox.Show("You must select the mine to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbMineEntrance.Focus();
                return false;
            }

            if (Convert.ToInt32(cmbGeologist.SelectedValue) == -1 || cmbGeologist.SelectedValue == null)
            {
                MessageBox.Show("You must select a Geologist to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbGeologist.Focus();
                return false;
            }

            if (cmbChannelType.Text == string.Empty)
            {
                MessageBox.Show("You must select the sampling instrument to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbChannelType.Focus();
                return false;
            }

            if (!txtChId.Text.Contains("ES_MI") && !txtChId.Text.Contains("PV_MI") && !txtChId.Text.Contains("SK_MI"))
            {
                MessageBox.Show("You must enter the channel identifier to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtChId.Focus();
                return false;
            }

            if (txtChId.Text == string.Empty)
            {
                MessageBox.Show("You must enter the channel identifier to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtChId.Focus();
                return false;
            }

            if (txtSample.Text == string.Empty)
            {
                MessageBox.Show("You must enter the number sample to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSample.Focus();
                return false;
            }

            return true;
        }
        #endregion

        #region Limpiadores
        private void LimpiarControles()
        {
            cmbMineEntrance.SelectedValue = "-1";
            cmbGeologist.SelectedValue = "-1";
            cmbChannelType.SelectedValue = "-1";
            cmbLithology.SelectedValue = "-1";

            txtChId.Text = string.Empty;
            txtChId.Enabled = true;
            txtSample.Text = string.Empty;
            txtSample.Enabled = true;
            cmbSampleType.SelectedValue = "ORIGINAL";
            cmbSamplingType.SelectedValue = "-1";
            txtMTS.Text = string.Empty;
            txtFrom.Text = string.Empty;
            txtTo.Text = string.Empty;
            
            cmbVeinName.SelectedValue = string.Empty;
            txtDescription.Text = string.Empty;
            swActualizarRegistro = false;

            minaSeleccionada = string.Empty;
            geologoSeleccionado = string.Empty;
            tipoCanalSeleccionado = string.Empty;
            channelPrimero = string.Empty;
            chLength = 0;
            wSKCHChannels = 0;
            valorFinalMasUno = 0;
            sampleFaltante = string.Empty;
            cmbChannelId.DataSource = null;
            cmbSample.DataSource = null;
            CargarComboChannelBusqueda();
            CargarComboSampleBusqueda();
        }

        private void ReiniciarControles()
        {
            cmbMineEntrance.SelectedValue = "-1";
            cmbGeologist.SelectedValue = "-1";
            cmbChannelType.SelectedValue = "-1";
            cmbLithology.SelectedValue = "-1";

            txtChId.Text = string.Empty;
            txtChId.Enabled = true;
            txtSample.Text = string.Empty;
            txtSample.Enabled = true;
            cmbSampleType.SelectedValue = "ORIGINAL";
            cmbSamplingType.SelectedValue = "-1";
            txtMTS.Text = string.Empty;
            txtFrom.Text = string.Empty;
            txtTo.Text = string.Empty;
            
            cmbVeinName.SelectedValue = string.Empty;
            txtDescription.Text = string.Empty;
            cmbChannelId.DataSource = null;
            cmbSample.DataSource = null;
            CargarComboChannelBusqueda();
            CargarComboSampleBusqueda();
        }

        private void LimpiarControlesEditable()
        {
            cmbMineEntrance.SelectedValue = "-1";
            cmbGeologist.SelectedValue = "-1";
            cmbChannelType.SelectedValue = "-1";
            cmbLithology.SelectedValue = "-1";

            txtChId.Text = string.Empty;
            txtChId.Enabled = true;
            txtSample.Text = string.Empty;
            txtSample.Enabled = true;
            cmbSampleType.SelectedValue = "ORIGINAL";
            cmbSamplingType.SelectedValue = "-1";
            txtMTS.Text = string.Empty;
            txtFrom.Text = string.Empty;
            txtTo.Text = string.Empty;
            
            cmbVeinName.SelectedValue = string.Empty;
            txtDescription.Text = string.Empty;
        }


        private void LimpiarTextboxDinamico(GroupBox groupbox)
        {
            foreach (Control control in groupbox.Controls)
            {
                if (control.GetType().Equals(typeof(TextBox)))
                {
                    control.Text = string.Empty;
                }
            }
        }
        #endregion

        #region Metodos Privados
        private void Loadcmb()
        {
            DataTable dtMineEnt = new DataTable();
            dtMineEnt = oRf.getMineEntranceExplora();
            DataRow drMineEnt = dtMineEnt.NewRow();
            drMineEnt[0] = "-1";
            drMineEnt[1] = "Select an option...";
            dtMineEnt.Rows.Add(drMineEnt);
            cmbMineEntrance.DisplayMember = "cmb";
            cmbMineEntrance.ValueMember = "MineID";
            cmbMineEntrance.DataSource = dtMineEnt;
            cmbMineEntrance.SelectedValue = "-1";
            swCombo = true;

            DataTable dtUsers = new DataTable();
            dtUsers = oRf.getUsers("-99");
            DataRow dr2 = dtUsers.NewRow();
            dr2[0] = "-1";
            dr2[7] = "Select an option..";
            dtUsers.Rows.Add(dr2);
            cmbGeologist.DisplayMember = "cmb";
            cmbGeologist.ValueMember = "id";
            cmbGeologist.DataSource = dtUsers;
            cmbGeologist.SelectedValue = "-1";

            DataSet dtSampleT = new DataSet();
            dtSampleT = oRf.getRfTypeSampleDataSet();
            DataRow drSTy = dtSampleT.Tables[1].NewRow();
            drSTy[0] = "-1";
            drSTy[1] = "Select an option..";
            dtSampleT.Tables[1].Rows.Add(drSTy);
            cmbSampleType.DisplayMember = "Comb";
            cmbSampleType.ValueMember = "Code";
            cmbSampleType.DataSource = dtSampleT.Tables[1];
            cmbSampleType.SelectedValue = "ORIGINAL";

            DataRow drT2 = dtSampleT.Tables[2].NewRow();
            drT2[0] = "-1";
            drT2[1] = "Select an option..";
            dtSampleT.Tables[2].Rows.Add(drT2);
            cmbSamplingType.DisplayMember = "Comb";
            cmbSamplingType.ValueMember = "Code";
            cmbSamplingType.DataSource = dtSampleT.Tables[2];
            cmbSamplingType.SelectedValue = -1;


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

            DataTable dtLithology = new DataTable();
            dtLithology = oRf.getDsRfLithology().Tables[1];
            cmbLithology.DisplayMember = "Comb";
            cmbLithology.ValueMember = "Code";
            cmbLithology.DataSource = dtLithology;
            cmbLithology.SelectedValue = "-1";

            DataRow drChn = dtSampleT.Tables[3].NewRow();
            drChn[0] = "-1";
            drChn[1] = "Select an option..";
            dtSampleT.Tables[3].Rows.Add(drChn);
            cmbChannelType.DisplayMember = "Comb";
            cmbChannelType.ValueMember = "Code";
            cmbChannelType.DataSource = dtSampleT.Tables[3];
            cmbChannelType.SelectedValue = "-1";

            CargarComboChannelBusqueda();
            CargarComboSampleBusqueda();
        }

        private void CargarComboChannelBusqueda()
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
            cmbChannelId.SelectedValue = "Select an option..";
        }

        private void CargarComboSampleBusqueda()
        {
            oCHSamp.sOpcion = "2";
            oCHSamp.sChId = string.Empty;
            DataTable dtSamp = new DataTable();
            dtSamp = oCHSamp.getCHSamplesByChid();
            DataRow drsample = dtSamp.NewRow();
            drsample["Sample"] = "Select an option..";
            dtSamp.Rows.Add(drsample);
            cmbSample.DisplayMember = "Sample";
            cmbSample.ValueMember = "Sample";
            cmbSample.DataSource = dtSamp;
            cmbSample.SelectedValue = "Select an option..";

            if (sSampleSelect != "" && sSampleSelect != "Select an option..")
            {
                cmbSample.SelectedValue = sSampleSelect.ToString();
                sSampleSelect = "";
            }
        }

        private void InhabilitarColumnasDataGrid()
        {
            dgData.Columns[0].ReadOnly = true;
            dgData.Columns[1].ReadOnly = true;
            dgData.Columns[2].ReadOnly = true;
            dgData.Columns[3].ReadOnly = true;
            dgData.Columns[4].ReadOnly = true;
            dgData.Columns[5].ReadOnly = true;
            dgData.Columns[6].ReadOnly = true;
            dgData.Columns[7].ReadOnly = true;
            dgData.Columns[8].ReadOnly = true;
            dgData.Columns[9].ReadOnly = true;
            dgData.Columns[10].ReadOnly = true;
            dgData.Columns[11].ReadOnly = true;
            dgData.Columns[11].Visible = false;
        }

        private void CleanControls()
        {
            if (dgData.DataSource != null)
            {
                ColumnasGrid();
            }
            else
            {
                dgData.Rows.Clear();
            }

            LimpiarControles();
        }

        private void ColumnasGrid()
        {
            dgData.DataSource = null;
            dgData.Columns.Add("chId", "Channel");
            dgData.Columns.Add("Sample", "Num. Samples");
            dgData.Columns.Add("SampleType", "Sample Type");
            dgData.Columns.Add("MTS", "MTS");
            dgData.Columns.Add("From", "From");
            dgData.Columns.Add("To", "To");
            dgData.Columns.Add("LRock", "Lithology");
            dgData.Columns.Add("VVeinName", "Cutting Width (m)");
            dgData.Columns.Add("SamplingLocation", "Sampling Location");
            dgData.Columns.Add("LRockObservations", "Description");
            dgData.Columns.Add("DateChann", "Date Channel");
            dgData.Columns.Add("ID", "ID");
            dgData.AllowUserToDeleteRows = false;
            InhabilitarColumnasDataGrid();
        }
        #endregion

        private void MinningGeology_Load(object sender, EventArgs e)
        {
            Loadcmb();
            ColumnasGrid();
            cmbMineEntrance.Focus();
        }

        private void btnCancel_Click(object sender, EventArgs e)
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

        private void btnAgregate_Click(object sender, EventArgs e)
        {
            if (!ValidarValores())
            {
                return;
            }

            minaSeleccionada = cmbMineEntrance.SelectedValue.ToString();
            geologoSeleccionado = cmbGeologist.SelectedValue.ToString();
            tipoCanalSeleccionado = cmbChannelType.SelectedValue.ToString();

            dgData.Rows.Add(txtChId.Text, txtSample.Text, cmbSampleType.SelectedValue.ToString() == "-1" ? string.Empty : cmbSampleType.SelectedValue,
                            txtMTS.Text, txtFrom.Text, txtTo.Text, cmbLithology.SelectedValue == null ? string.Empty : cmbLithology.SelectedValue,
                            cmbVeinName.SelectedValue.ToString() == "-1" ? string.Empty : cmbVeinName.SelectedValue,
                            cmbSamplingType.SelectedValue.ToString() == "-1" ? string.Empty : cmbSamplingType.SelectedValue, txtDescription.Text, dtimeDate.Text);

        }

        private void cmbMineEntrance_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(cmbMineEntrance.SelectedValue) != -1 && swCombo)
            {
                if (sampleFaltante != string.Empty)
                {
                    txtChId.Text = channelPrimero;
                    txtSample.Text = sampleFaltante;

                    txtChId.Enabled = false;
                    txtSample.Enabled = false;

                    cmbMineEntrance.SelectedValue = minaSeleccionada;
                    cmbGeologist.SelectedValue = geologoSeleccionado == string.Empty ? "-1" : geologoSeleccionado;
                    cmbChannelType.SelectedValue = tipoCanalSeleccionado;
                }
                else
                {
                    if (sEditCh == "0")
                    {
                        switch (Convert.ToInt32(cmbMineEntrance.SelectedValue))
                        {
                            case (int)Mines.ES:
                                txtChId.Text = string.Concat(Mines.ES.ToString(), "_MI");
                                break;

                            case (int)Mines.SK:
                                txtChId.Text = string.Concat(Mines.SK.ToString(), "_MI");
                                break;

                            case (int)Mines.PV:
                                txtChId.Text = string.Concat(Mines.PV.ToString(), "_MI");
                                break;
                        }
                    }

                    if (sEditCh == "1")
                    {
                        var concecutivo = validarConcecutivo();

                        valorFinalMasUno = (validarUltimoValor() + 1);
                        txtChId.Text = channelPrimero;
                        txtChId_Leave(null, null);
                        txtSample.Text = string.Concat("R", valorFinalMasUno);
                        txtSample_Leave(null, null);
                    }
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                if (sEditCh == "0")
                {
                    oCh.sOpcion = "1";
                    oCh.iSKCHChannels = 0;
                }
                else if (sEditCh == "1")
                {
                    oCh.sOpcion = "2";
                    oCh.iSKCHChannels = wSKCHChannels;
                }

                oCh.sChId = channelPrimero;
                oCh.dEast = null;
                oCh.dNorth = null;
                oCh.dElevation = null;
                oCh.dLenght = chLength;
                oCh.sProjection = string.Empty;
                oCh.sDatum = string.Empty;
                oCh.sProject = "GZC";
                oCh.sClaim = string.Empty;

                oCh.sStartDate = Convert.ToDateTime(dtimeDate.Text).ToShortDateString();
                oCh.sFinalDate = Convert.ToDateTime(dtimeDate.Text).ToShortDateString();

                oCh.sStorage = string.Empty;
                oCh.sSource = string.Empty;

                oCh.sComments = txtDescription.Text;
                oCh.sMineID = minaSeleccionada;
                oCh.sType = tipoCanalSeleccionado;
                oCh.sInstrument = tipoCanalSeleccionado;

                if (cantMuestras == 0)
                {
                    oCh.iSamplesTotal = null;
                }
                else
                {
                    oCh.iSamplesTotal = cantMuestras;
                }

                string sRespAdd = oCh.CH_Collars_Add();
                if (sRespAdd == "OK")
                {
                    for (int i = 0; i < dgData.RowCount - 1; i++)
                    {
                        if (sEditCh == "0")
                        {
                            oCHSamp.sOpcion = "1";
                            oCHSamp.iSKCHSamples = 0;
                        }
                        else if (sEditCh == "1")
                        {
                            if ((dgData.Rows[i].Cells[11].Value == null ? 0 : Convert.ToInt32(dgData.Rows[i].Cells[11].Value)) == 0)
                            {
                                oCHSamp.sOpcion = "1";
                                oCHSamp.iSKCHSamples = 0;
                            }
                            else
                            {
                                oCHSamp.sOpcion = "2";
                                oCHSamp.iSKCHSamples = Convert.ToInt32(dgData.Rows[i].Cells[11].Value);
                            }
                        }

                        oCHSamp.sChId = dgData.Rows[i].Cells[0].Value.ToString();
                        oCHSamp.sSample = dgData.Rows[i].Cells[1].Value.ToString();
                        oCHSamp.dFrom = double.Parse(dgData.Rows[i].Cells[4].Value.ToString());
                        oCHSamp.dTo = double.Parse(dgData.Rows[i].Cells[5].Value.ToString());
                        oCHSamp.sTarget = dgData.Rows[i].Cells[7].Value.ToString();
                        oCHSamp.sProject = "GZC";
                        oCHSamp.sGeologist = geologoSeleccionado;
                        oCHSamp.sHelper = string.Empty;
                        oCHSamp.sStation = string.Empty;

                        DateTime dDate = DateTime.Parse(dgData.Rows[i].Cells[10].Value.ToString());
                        string sDate = dDate.Year.ToString().PadLeft(4, '0') + dDate.Month.ToString().PadLeft(2, '0') +
                            dDate.Day.ToString().PadLeft(2, '0');

                        oCHSamp.sDate = sDate.ToString();
                        oCHSamp.dE = null;
                        oCHSamp.dN = null;
                        oCHSamp.dZ = null;
                        oCHSamp.dE2 = null;
                        oCHSamp.dN2 = null;
                        oCHSamp.dZ2 = null;
                        oCHSamp.sCS = null;
                        oCHSamp.dGPSEpe = null;
                        oCHSamp.sPhoto = null;
                        oCHSamp.sPhotoAzimuth = null;
                        oCHSamp.sSampleType = dgData.Rows[i].Cells[2].Value.ToString();
                        oCHSamp.sSamplingType = dgData.Rows[i].Cells[8].Value.ToString() == string.Empty ? null : dgData.Rows[i].Cells[8].Value.ToString();
                        oCHSamp.sPorpouse = null;
                        oCHSamp.sRelativeLoc = null;
                        oCHSamp.dLenght = dgData.Rows[i].Cells[5].Value.ToString() == "-99" ? 0 : (double.Parse(dgData.Rows[i].Cells[5].Value.ToString()) - double.Parse(dgData.Rows[i].Cells[4].Value.ToString()));

                        oCHSamp.dHigh = null;
                        oCHSamp.sThickness = null;
                        oCHSamp.sObservations = dgData.Rows[i].Cells[9].Value.ToString();
                        oCHSamp.sLRock = dgData.Rows[i].Cells[6].Value == null ? string.Empty : dgData.Rows[i].Cells[6].Value.ToString();
                        oCHSamp.sLTexture = null;
                        oCHSamp.sLGSize = null;
                        oCHSamp.sLWeathering = null;
                        oCHSamp.sLRockSorting = null;
                        oCHSamp.sLRockSphericity = null;
                        oCHSamp.sLRockRounding = null;
                        oCHSamp.sLRockObservation = null;
                        oCHSamp.sLMatrixGSize = null;
                        oCHSamp.sLMatrixObservations = null;
                        oCHSamp.sLMatrixPerc = null;
                        oCHSamp.sLPhenoCPerc = null;
                        oCHSamp.sLPhenoCGSize = null;
                        oCHSamp.sLPhenoCObservations = null;
                        oCHSamp.sVContactType = null;
                        oCHSamp.sVVeinName = dgData.Rows[i].Cells[8].Value.ToString() == string.Empty ? null : dgData.Rows[i].Cells[8].Value.ToString();
                        oCHSamp.sVHostRock = null;
                        oCHSamp.sVObservations = null;
                        oCHSamp.sDupOf = null;
                        oCHSamp.bValited = false;
                        oCHSamp.iSampleCont = null;

                        string sResp = oCHSamp.CH_Samples_Add();
                        if (sResp != "OK")
                        {
                            MessageBox.Show("Save Error: " + sResp.ToString(), "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    MessageBox.Show("Channels saved successfully.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CleanControls();
                    sampleExtradido = 0;
                    cantMuestras = 0;
                    conteoMuestras = 0;
                    swActualizarRegistro = false;
                    indexRegistroGrid = 0;
                    sFrom = 0;
                    sTo = 0;
                    sEditCh = "0";
                    cmbMineEntrance.Focus();
                }
                else
                {
                    MessageBox.Show("We have problems saving your information.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    sEditCh = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Concat("Error: ", ex.Message), "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                sEditCh = "0";
            }
        }

        private void txtChId_Leave(object sender, EventArgs e)
        {
            if (txtChId.Text != string.Empty)
            {
                if (txtChId.Text.Length <= 5)
                {
                    MessageBox.Show("You must enter the channel identifier to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtChId.Focus();
                    return;
                }
                else
                {
                    txtChId.Enabled = false;
                    channelPrimero = txtChId.Text;
                    int i = 0;
                    string respuesta = string.Empty;

                    if (sEditCh == "0")
                    {
                        respuesta = Microsoft.VisualBasic.Interaction.InputBox("Enter the number of samples: ", "Minning Geology", string.Empty);
                        if (!int.TryParse(respuesta, out i))
                        {
                            MessageBox.Show("You must enter numbers only to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            respuesta = Microsoft.VisualBasic.Interaction.InputBox("Enter the number of samples: ", "Minning Geology", string.Empty);
                        }

                        if (respuesta == string.Empty)
                        {
                            if (!int.TryParse(respuesta, out i))
                            {
                                MessageBox.Show("You must enter numbers only to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                respuesta = Microsoft.VisualBasic.Interaction.InputBox("Enter the number of samples: ", "Minning Geology", string.Empty);
                            }

                            if (respuesta == string.Empty)
                            {
                                MessageBox.Show("A single sample has been assigned for the channel.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                cantMuestras = 1;
                            }
                            else
                            {
                                cantMuestras = Convert.ToInt32(respuesta);
                            }
                        }
                        else
                        {
                            cantMuestras = Convert.ToInt32(respuesta);
                        }
                    }

                    Match val = Regex.Match(txtChId.Text, "(\\d+)");
                    txtSample.Text = string.Concat("R", Convert.ToInt32(val.Value));
                    conteoMuestras = 1;
                }
            }
        }

        private void txtSample_Leave(object sender, EventArgs e)
        {
            if (txtSample.Text != string.Empty)
            {
                if (txtSample.Text.Length <= 1)
                {
                    MessageBox.Show("You must enter the number sample to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtSample.Enabled = true;
                    txtSample.Focus();
                    return;
                }
            }

            txtSample.Enabled = false;
        }

        private void txtMTS_Leave(object sender, EventArgs e)
        {
            if (!ValidarValores())
            {
                return;
            }

            if (cmbSampleType.SelectedValue.ToString() != "-1")
            {
                if (cmbSampleType.SelectedValue.ToString() == "ORIGINAL")
                {
                    if (txtMTS.Text == string.Empty)
                    {
                        return;
                    }

                    if (!swActualizarRegistro)
                    {
                        if (conteoMuestras == 1)
                        {
                            if (sEditCh == "1")
                            {
                                txtFrom.Text = UltimoTo().ToString();
                                sTo = UltimoTo();
                                sFrom = Convert.ToDouble(txtFrom.Text);
                                txtTo.Text = (sTo + Convert.ToDouble(txtMTS.Text)).ToString();
                                sTo = Convert.ToInt32(txtTo.Text);
                            }
                            else
                            {
                                txtFrom.Text = sFrom.ToString();
                                txtTo.Text = txtMTS.Text;
                                sTo = Convert.ToDouble(txtMTS.Text);
                            }
                        }
                        else
                        {
                            txtFrom.Text = sTo.ToString();
                            sFrom = Convert.ToDouble(txtFrom.Text);
                            txtTo.Text = (sTo + Convert.ToDouble(txtMTS.Text)).ToString();
                            sTo = Convert.ToDouble(txtTo.Text);
                        }
                    }
                    else
                    {
                        ActualizarRegistroDataGrid(indexRegistroGrid);
                        ReiniciarControles();
                        return;
                    }
                }
                else
                {
                    txtFrom.Text = "-99";
                    txtTo.Text = "-99";
                }

                if (!swActualizarRegistro)
                {
                    btnAgregate_Click(null, null);
                    txtFrom.Text = string.Empty;
                    txtTo.Text = string.Empty;
                    txtMTS.Text = string.Empty;
                    cmbSampleType.SelectedValue = "-1";
                    cmbSampleType.Focus();

                    if (sampleFaltante != string.Empty)
                    {
                        ReiniciarControles();
                        return;
                    }

                    if (sampleExtradido == 0)
                    {
                        Match val = Regex.Match(txtSample.Text, "(\\d+)");
                        sampleExtradido = Convert.ToInt32(val.Value);
                        sampleExtradido++;
                    }
                    else
                    {
                        sampleExtradido++;
                    }

                    if (cantMuestras == conteoMuestras) //&& sEditCh == "0")
                    {
                        chLength = sTo;
                        btnAgregate.Enabled = false;
                        return;
                    }

                    txtSample.Text = string.Concat("R", sampleExtradido);
                    conteoMuestras++;
                }
                else
                {
                    ActualizarRegistroDataGrid(indexRegistroGrid);
                    ReiniciarControles();
                    return;
                }
            }
        }

        private int UltimoTo()
        {
            int encontrado = 0;
            int[] valoresSamples = new int[dgData.RowCount - 1];

            for (int i = 0; i < dgData.RowCount - 1; i++)
            { 
                valoresSamples[i] = Convert.ToInt32(dgData.Rows[i].Cells[5].Value.ToString().Replace(".00",""));
            }

            for (int y = 0; y <= valoresSamples.Length - 1; y++)
            {
                if (y == (valoresSamples.Length - 1))
                {
                    if (valoresSamples[y] == -99)
                        encontrado = valoresSamples[y - 1];
                    else
                        encontrado = valoresSamples[y];
                }
            }

            return encontrado;
        }


        private void ActualizarRegistroDataGrid(int index)
        {
            if (dgData.Rows.Count > 1)
            {
                if (cmbSampleType.SelectedValue.ToString() == "ORIGINAL")
                {
                    txtFrom.Text = UltimoTo().ToString();
                    sTo = UltimoTo();
                    sFrom = Convert.ToDouble(txtFrom.Text);
                    txtTo.Text = (sTo + Convert.ToDouble(txtMTS.Text)).ToString();
                    sTo = Convert.ToInt32(txtTo.Text);
                }

                dgData.Rows[index].Cells[0].Value = txtChId.Text;
                dgData.Rows[index].Cells[1].Value = txtSample.Text;
                dgData.Rows[index].Cells[2].Value = cmbSampleType.SelectedValue;
                dgData.Rows[index].Cells[3].Value = txtMTS.Text;
                dgData.Rows[index].Cells[4].Value = txtFrom.Text;
                dgData.Rows[index].Cells[5].Value = txtTo.Text;
                dgData.Rows[index].Cells[6].Value = cmbLithology.SelectedValue == null ? string.Empty : cmbLithology.Text;
                dgData.Rows[index].Cells[7].Value = cmbVeinName.SelectedValue.ToString() == "-1" ? string.Empty : cmbVeinName.Text;
                dgData.Rows[index].Cells[8].Value = cmbSamplingType.SelectedValue.ToString() == "-1" ? string.Empty : cmbSamplingType.Text;
                dgData.Rows[index].Cells[9].Value = txtDescription.Text;

                swActualizarRegistro = false;
                indexRegistroGrid = 0;

                txtChId.Enabled = true;
                txtSample.Enabled = true;
            }
        }

        private void cmbSampleType_Leave(object sender, EventArgs e)
        {
            if (cmbSampleType.SelectedValue.ToString() != "ORIGINAL")
            {
                txtMTS.Enabled = false;
                txtMTS_Leave(null, null);
            }
            else
            {
                txtMTS.Enabled = true;
            }
        }

        private void cmbSample_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                if (cmbSample.SelectedValue.ToString() != "Select an option..")
                {
                    sampleSeleccionado = cmbSample.SelectedValue.ToString();
                    oCHSamp.sSamp1 = cmbSample.SelectedValue.ToString();
                    oCHSamp.sSamp2 = cmbSample.SelectedValue.ToString();
                    oCHSamp.sChId = string.Empty;

                    DataTable dgSamp = oCHSamp.getCHSamplesListchIdSample();

                    oCHSamp.sOpcion = "1";
                    oCHSamp.sChId = dgSamp.Rows[0][1].ToString();
                    channelPrimero = dgSamp.Rows[0][1].ToString();
                    DataTable dgListaSamples = oCHSamp.getCHSamplesByChid();

                    oCh.sChId = dgSamp.Rows[0][1].ToString();
                    oCh.sOpcion = "2";
                    DataTable dtCollar = oCh.getCH_Collars();

                    wSKCHChannels = Convert.ToInt32(dtCollar.Rows[0][0]);
                    chLength = Convert.ToInt32(dtCollar.Rows[0][2]);
                    minaSeleccionada = dtCollar.Rows[0][16].ToString();
                    geologoSeleccionado = dgSamp.Rows[0][7].ToString();
                    tipoCanalSeleccionado = dtCollar.Rows[0][17].ToString();

                    cantMuestras = Convert.ToInt32(dtCollar.Rows[0][20]);

                    foreach (DataRow lista in dgListaSamples.Rows)
                    {
                        dgData.Rows.Add(lista[1].ToString(), lista[2].ToString(), lista[21].ToString(), lista[3].ToString() == "-99.00" ? string.Empty : (Convert.ToDouble(lista[4]) - Convert.ToDouble(lista[3])).ToString(),
                            lista[3].ToString(), lista[4].ToString(), lista[31].ToString(), lista[46].ToString(), lista[22].ToString(), lista[30].ToString(), lista[10].ToString(), lista[0].ToString());
                    }

                    DataRow[] myRow = dgListaSamples.Select(@"Sample = '" + cmbSample.SelectedValue.ToString() + "'");
                    int rowindex = dgListaSamples.Rows.IndexOf(myRow[0]);

                    dgData.Rows[rowindex].Selected = true;
                    dgData.CurrentCell = dgData.Rows[rowindex].Cells[1];
                    sEditCh = "1";

                    sampleFaltante = validarConcecutivo() == 0 ? string.Empty : string.Concat("R", validarConcecutivo());
                    valorFinalMasUno = (validarUltimoValor() + 1);
                    cmbMineEntrance.SelectedValue = minaSeleccionada;
                    cmbGeologist.SelectedValue = geologoSeleccionado == string.Empty ? "-1" : geologoSeleccionado;
                    cmbChannelType.SelectedValue = tipoCanalSeleccionado;
                    cmbSampleType.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Concat("An error has occurred: ", ex.Message), "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void cmbChannelId_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                if (cmbChannelId.SelectedValue.ToString() != "Select an option..")
                {
                    oCh.sChId = cmbChannelId.SelectedValue.ToString();
                    channelPrimero = cmbChannelId.SelectedValue.ToString();
                    oCh.sOpcion = "2";
                    DataTable dtCollar = oCh.getCH_Collars();

                    oCHSamp.sOpcion = "1";
                    oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
                    DataTable dgListaSamples = oCHSamp.getCHSamplesByChid();

                    wSKCHChannels = Convert.ToInt32(dtCollar.Rows[0][0]);
                    minaSeleccionada = dtCollar.Rows[0][16].ToString();
                    geologoSeleccionado = dgListaSamples.Rows[0][7].ToString();
                    tipoCanalSeleccionado = dtCollar.Rows[0][17].ToString();

                    cantMuestras = Convert.ToInt32(dtCollar.Rows[0][20]);

                    foreach (DataRow lista in dgListaSamples.Rows)
                    {
                        dgData.Rows.Add(lista[1].ToString(), lista[2].ToString(), lista[21].ToString(), lista[3].ToString() == "-99.00" ? string.Empty : (Convert.ToDouble(lista[4]) - Convert.ToDouble(lista[3])).ToString(),
                            lista[3].ToString(), lista[4].ToString(), lista[31].ToString(), lista[46].ToString(), lista[22].ToString(), lista[30].ToString(), lista[10].ToString(), lista[0].ToString());
                    }

                    sEditCh = "1";
                }

                sampleFaltante = string.Concat("R", validarConcecutivo());
                valorFinalMasUno = (validarUltimoValor() + 1);
                cmbMineEntrance.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Concat("An error has occurred: ", ex.Message), "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void dgData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgData.Rows.Count > 1)
            {
                cmbMineEntrance.SelectedValue = minaSeleccionada;
                cmbGeologist.SelectedValue = geologoSeleccionado == string.Empty ? "-1" : geologoSeleccionado;
                cmbChannelType.SelectedValue = tipoCanalSeleccionado;

                txtChId.Text = dgData.Rows[e.RowIndex].Cells[0].Value.ToString();
                txtSample.Text = dgData.Rows[e.RowIndex].Cells[1].Value.ToString();
                cmbSampleType.Text = dgData.Rows[e.RowIndex].Cells[2].Value.ToString();
                cmbLithology.Text = dgData.Rows[e.RowIndex].Cells[6].Value == null ? string.Empty : dgData.Rows[e.RowIndex].Cells[6].Value.ToString();
                cmbVeinName.Text = dgData.Rows[e.RowIndex].Cells[7].Value == null ? string.Empty : dgData.Rows[e.RowIndex].Cells[7].Value.ToString();
                cmbSamplingType.Text = dgData.Rows[e.RowIndex].Cells[8].Value == null ? string.Empty : dgData.Rows[e.RowIndex].Cells[8].Value.ToString();
                txtDescription.Text = dgData.Rows[e.RowIndex].Cells[9].Value.ToString();
                dtimeDate.Text = Convert.ToDateTime(dgData.Rows[e.RowIndex].Cells[10].Value).ToShortDateString();
                txtFrom.Text = dgData.Rows[e.RowIndex].Cells[4].Value.ToString();
                txtTo.Text = dgData.Rows[e.RowIndex].Cells[5].Value.ToString();
                wSKCHChannels = Convert.ToInt32(dgData.Rows[e.RowIndex].Cells[11].Value);

                if (cmbSampleType.Text != "ORIGINAL")
                {
                    txtMTS.Enabled = false;
                }
                else
                {
                    txtMTS.Enabled = true;
                }

                sEditCh = "1";
                swActualizarRegistro = true;
                indexRegistroGrid = e.RowIndex;

                txtChId.Enabled = false;
                txtSample.Enabled = false;
            }
        }

        private void dgData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (MessageBox.Show("Do you really want to delete the sample?", "Minning Geology", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                oCHSamp.iSKCHSamples = Convert.ToInt32(dgData.Rows[e.RowIndex].Cells[11].Value);
                oCHSamp.CH_Samples_Delete();
                MessageBox.Show("Channels deleted successfully.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LlenarDataGrid();
                LimpiarControlesEditable();
                cmbChannelId.DataSource = null;
                cmbSample.DataSource = null;
                CargarComboChannelBusqueda();
                CargarComboSampleBusqueda();
            }
        }
               
        private void LlenarDataGrid()
        {
            oCh.sChId = channelPrimero;
            oCh.sOpcion = "2";
            DataTable dtCollar = oCh.getCH_Collars();

            oCHSamp.sOpcion = "1";
            oCHSamp.sChId = channelPrimero;
            DataTable dgListaSamples = oCHSamp.getCHSamplesByChid();

            wSKCHChannels = Convert.ToInt32(dtCollar.Rows[0][0]);
            minaSeleccionada = dtCollar.Rows[0][16].ToString();
            geologoSeleccionado = dgListaSamples.Rows[0][7].ToString();
            tipoCanalSeleccionado = dtCollar.Rows[0][17].ToString();

            cantMuestras = Convert.ToInt32(dtCollar.Rows[0][20]);

            dgData.Rows.Clear();
            foreach (DataRow lista in dgListaSamples.Rows)
            {
                dgData.Rows.Add(lista[1].ToString(), lista[2].ToString(), lista[21].ToString(), lista[3].ToString() == "-99.00" ? string.Empty : (Convert.ToDouble(lista[4]) - Convert.ToDouble(lista[3])).ToString(),
                    lista[3].ToString(), lista[4].ToString(), lista[31].ToString(), lista[46].ToString(), lista[22].ToString(), lista[30].ToString(), lista[10].ToString(), lista[0].ToString());
            }

            sEditCh = "1";
        }

        private int validarConcecutivo()
        {
            int encontrado = 0;
            int[] valoresSamples = new int[dgData.RowCount - 1];

            for (int i = 0; i < dgData.RowCount - 1; i++)
            {
                Match val = Regex.Match(dgData.Rows[i].Cells[1].Value.ToString(), "(\\d+)");
                valoresSamples[i] = Convert.ToInt32(val.Value);
            }

            for (int y = 0; y < valoresSamples.Length; y++)
            {
                if (y != 0)
                {
                    if ((valoresSamples[y] - 1) != valoresSamples[y - 1])
                    {
                        encontrado = valoresSamples[y] - 1;
                    }
                }
            }

            cmbMineEntrance.Focus();
            return encontrado;
        }

        private int validarUltimoValor()
        {
            int encontrado = 0;
            int[] valoresSamples = new int[dgData.RowCount - 1];

            for (int i = 0; i < dgData.RowCount - 1; i++)
            {
                Match val = Regex.Match(dgData.Rows[i].Cells[1].Value.ToString(), "(\\d+)");
                valoresSamples[i] = Convert.ToInt32(val.Value);
            }

            for (int y = 0; y <= valoresSamples.Length - 1; y++)
            {
                if (y == (valoresSamples.Length - 1))
                {
                    encontrado = valoresSamples[y];
                }
            }

            return encontrado;
        }

        private void cmbChannelType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbChannelType.SelectedValue != null && swCombo)
            {
                tipoCanalSeleccionado = cmbChannelType.SelectedValue.ToString();
            }
        }

        private void cmbGeologist_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbGeologist.SelectedValue != null && swCombo)
            {
                geologoSeleccionado = cmbGeologist.SelectedValue.ToString();
            }
        }

        private void txtMTS_KeyPress(object sender, KeyPressEventArgs e)
        {
            CultureInfo cc = System.Threading.Thread.CurrentThread.CurrentCulture;

            if (e.KeyChar >= 48 && e.KeyChar <= 57)
            {
                e.Handled = false;
            }
            else if (e.KeyChar == 8)
            {
                e.Handled = false;
            }
            else if (e.KeyChar == 13)
            {
                e.Handled = false;
            }
            else if (e.KeyChar.ToString() == cc.NumberFormat.NumberDecimalSeparator)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void btnExporPDFAll_Click(object sender, EventArgs e)
        {
            var dataGrid = GetContentAsDataTable(dgData);
            FrmRptMinningGeology reporte = new FrmRptMinningGeology(dataGrid);
            reporte.Show();
        }

        public DataTable GetContentAsDataTable(DataGridView dgv)
        {
            try
            {
                if (dgv.ColumnCount == 0) return null;
                DataTable dtSource = new DataTable();
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (col.Name == string.Empty) continue;
                    dtSource.Columns.Add(col.Name);
                    dtSource.Columns[col.Name].Caption = col.HeaderText;
                }

                dtSource.Columns.Add("NombreMina");
                dtSource.Columns.Add("NombreGeologo");
                dtSource.Columns["NombreMina"].Caption = "NombreMina";
                dtSource.Columns["NombreGeologo"].Caption = "NombreGeologo";
                
                if (dtSource.Columns.Count == 0) return null;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    DataRow drNewRow = dtSource.NewRow();
                    foreach (DataColumn col in dtSource.Columns)
                    {
                        if (col.ColumnName != "NombreMina" && col.ColumnName != "NombreGeologo")
                            drNewRow[col.ColumnName] = row.Cells[col.ColumnName].Value;

                        if (col.ColumnName == "NombreMina")
                            drNewRow[col.ColumnName] = cmbMineEntrance.Text.ToUpper();

                        if (col.ColumnName == "NombreGeologo")
                            drNewRow[col.ColumnName] = cmbGeologist.Text.ToUpper();
                    }

                    dtSource.Rows.Add(drNewRow);
                }
                return dtSource;
            }
            catch { return null; }
        }

    }
}
