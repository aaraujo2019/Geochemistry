using Geochemistry.Emun;
using System;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Geochemistry
{
    public partial class frmMinningGeology : Form
    {
        private clsRf oRf = new clsRf();
        private clsCHChannels oCh = new clsCHChannels();
        clsCHSamples oCHSamp = new clsCHSamples();
        private bool swCombo = false;
        private int sampleExtradido = 0;
        private int cantMuestras = 0;
        private int conteoMuestras = 0;
        private int sFrom = 0;
        private int sTo = 0;
        private string sSampleSelect = "";
        private bool swActualizarRegistro = false;
        private int indexRegistroGrid = 0;

        private string minaSeleccionada = string.Empty;
        private string geologoSeleccionado = string.Empty;
        private string tipoCanalSeleccionado = string.Empty;

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
            bool valor = false;

            if (Convert.ToInt32(cmbMineEntrance.SelectedValue) == -1)
            {
                MessageBox.Show("You must select the mine to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbMineEntrance.Focus();
                valor = true;
            }

            if (Convert.ToInt32(cmbGeologist.SelectedValue) == -1)
            {
                MessageBox.Show("You must select a Geologist to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbMineEntrance.Focus();
                valor = true;
            }

            if (cmbChannelType.Text == string.Empty)
            {
                MessageBox.Show("You must select the sampling instrument to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbMineEntrance.Focus();
                valor = true;
            }

            if (!txtChId.Text.Contains("ES_MI") && !txtChId.Text.Contains("PV_MI") && txtChId.Text.Contains("SK_MI"))
            {
                MessageBox.Show("You must enter the channel identifier to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbMineEntrance.Focus();
                valor = true;
            }

            if (txtChId.Text == string.Empty)
            {
                MessageBox.Show("You must enter the channel identifier to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtChId.Focus();
                valor = true;
            }

            if (txtSample.Text == string.Empty)
            {
                MessageBox.Show("You must enter the number sample to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSample.Focus();
                valor = true;
            }

            return valor;
        }
        #endregion

        #region Limpiadores
        private void LimpiarControles()
        {
            cmbMineEntrance.SelectedValue = "-1";
            cmbGeologist.SelectedValue = "-1";
            cmbChannelType.SelectedValue = "-1";
            txtChId.Text = string.Empty;
            txtChId.Enabled = true;
            txtSample.Text = string.Empty;
            txtSample.Enabled = true;
            cmbSampleType.SelectedValue = "ORIGINAL";
            txtSamplePlace.Text = string.Empty;
            txtMTS.Text = string.Empty;
            txtFrom.Text = string.Empty;
            txtTo.Text = string.Empty;
            cmbLithology.SelectedValue = "-1";
            cmbVeinName.SelectedValue = string.Empty;
            txtDescription.Text = string.Empty;
        }

        private void LimpiarTextboxDinamico(GroupBox groupbox)
        {
            foreach (Control control in groupbox.Controls)
            {
                if (control.GetType().Equals(typeof(TextBox))) control.Text = string.Empty;
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

            oCHSamp.sOpcion = "1";
            oCHSamp.sChId = cmbChannelId.SelectedValue.ToString();
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
        }

        private void CleanControls()
        {
            if (dgData.DataSource != null)
            {
                ColumnasGrid();
            }
            else dgData.Rows.Clear();

            LimpiarControles();
        }

        private void ColumnasGrid()
        {
            dgData.DataSource = null;
            dgData.Columns.Add("chId", "Channel");
            dgData.Columns.Add("Sample", "Num. Samples");
            dgData.Columns.Add("SampleType", "Sample Type");
            dgData.Columns.Add("SamplingPlace", "Sampling Place");
            dgData.Columns.Add("MTS", "MTS");
            dgData.Columns.Add("From", "From");
            dgData.Columns.Add("To", "To");
            dgData.Columns.Add("LRock", "Lithology");
            dgData.Columns.Add("VVeinName", "Cutting Width (m)");
            dgData.Columns.Add("LRockObservations", "Description");
            dgData.Columns.Add("DateChann", "Date Channel");
            dgData.AllowUserToDeleteRows = false;
            InhabilitarColumnasDataGrid();
        }
        #endregion

        private void MinningGeology_Load(object sender, EventArgs e)
        {
            Loadcmb();
            cmbMineEntrance.Focus();
            ColumnasGrid();
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
            if (ValidarValores()) return;

            minaSeleccionada = cmbMineEntrance.SelectedValue.ToString();
            geologoSeleccionado = cmbGeologist.SelectedValue.ToString();
            tipoCanalSeleccionado = cmbChannelType.SelectedValue.ToString();

            dgData.Rows.Add(txtChId.Text, txtSample.Text, cmbSampleType.SelectedValue, txtSamplePlace.Text, 
                            txtMTS.Text, txtFrom.Text, txtTo.Text, cmbLithology.SelectedValue, 
                            cmbVeinName.SelectedValue.ToString() == "-1" ? string.Empty : cmbVeinName.SelectedValue, 
                            txtDescription.Text, dtimeDate.Text);
        }
        
        private void cmbMineEntrance_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(cmbMineEntrance.SelectedValue) != -1 && swCombo)
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
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            sampleExtradido = 0;
            cantMuestras = 0;
            conteoMuestras = 0;
            swActualizarRegistro = false;
            indexRegistroGrid = 0;
            sFrom = 0;
            sTo = 0;
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
                    int i = 0;
                    string respuesta = string.Empty;

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
                        else cantMuestras = Convert.ToInt32(respuesta);
                    }
                    else cantMuestras = Convert.ToInt32(respuesta);

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
                    txtChId.Enabled = true;
                    txtChId.Focus();
                    return;
                }
            }

            txtSample.Enabled = false;
        }

        private void txtMTS_Leave(object sender, EventArgs e)
        {
            if (cmbSampleType.SelectedValue.ToString() != "-1")
            {
                if (cmbSampleType.SelectedValue.ToString() == "ORIGINAL")
                {
                    if (txtMTS.Text == string.Empty)
                    {
                        MessageBox.Show("The MTS field cannot be empty.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        txtMTS.Focus();
                        return;
                    }

                    if (!swActualizarRegistro)
                    {
                        if (conteoMuestras == 1)
                        {
                            txtFrom.Text = sFrom.ToString();
                            txtTo.Text = txtMTS.Text;
                            sTo = Convert.ToInt32(txtMTS.Text);
                        }
                        else
                        {
                            txtFrom.Text = sTo.ToString();
                            sFrom = Convert.ToInt32(txtFrom.Text);
                            txtTo.Text = (sTo + Convert.ToInt32(txtMTS.Text)).ToString();
                            sTo = Convert.ToInt32(txtTo.Text);
                        }
                    }
                    else
                    {
                        ActualizarRegistroDataGrid(indexRegistroGrid);
                        LimpiarControles();
                        return;
                    }
                }
                else
                {
                    txtFrom.Text = "-99";
                    txtTo.Text = "-99";
                }

                btnAgregate_Click(null, null);
                txtFrom.Text = string.Empty;
                txtTo.Text = string.Empty;
                txtMTS.Text = string.Empty;
                cmbSampleType.SelectedValue = "-1";
                cmbSampleType.Focus();

                if (sampleExtradido == 0)
                {
                    Match val = Regex.Match(txtSample.Text, "(\\d+)");
                    sampleExtradido = Convert.ToInt32(val.Value);
                    sampleExtradido++;
                }
                else sampleExtradido++; 

                if (cantMuestras == conteoMuestras)
                {
                    btnAgregate.Enabled = false;
                    LimpiarControles();
                    return;
                }

                txtSample.Text = string.Concat("R", sampleExtradido);
                conteoMuestras++;
            }           
        }

        private void ActualizarRegistroDataGrid(int index)
        {
            if (dgData.Rows.Count > 1)
            {
                txtFrom.Text = sTo.ToString();
                sFrom = Convert.ToInt32(txtFrom.Text);
                txtTo.Text = (sTo + Convert.ToInt32(txtMTS.Text)).ToString();
                sTo = Convert.ToInt32(txtTo.Text);

                dgData.Rows[index].Cells[0].Value = txtChId.Text;
                dgData.Rows[index].Cells[1].Value = txtSample.Text;
                dgData.Rows[index].Cells[2].Value = cmbSampleType.SelectedValue;
                dgData.Rows[index].Cells[3].Value = txtSamplePlace.Text;
                dgData.Rows[index].Cells[4].Value = txtMTS.Text;
                dgData.Rows[index].Cells[5].Value = txtFrom.Text;
                dgData.Rows[index].Cells[6].Value = txtTo.Text;
                dgData.Rows[index].Cells[7].Value = cmbLithology.SelectedValue;
                dgData.Rows[index].Cells[8].Value = cmbVeinName.SelectedValue;
                dgData.Rows[index].Cells[9].Value = txtDescription.Text;

                swActualizarRegistro = false;
                indexRegistroGrid = 0;

                txtChId.Enabled = true;
                txtSample.Enabled = true;
            }
        }

        private void dgData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgData.Rows.Count > 1)
            {
                cmbMineEntrance.SelectedValue = minaSeleccionada;
                cmbGeologist.SelectedValue = geologoSeleccionado;
                cmbChannelType.SelectedValue = tipoCanalSeleccionado;

                txtChId.Text = dgData.Rows[e.RowIndex].Cells[0].Value.ToString();
                txtSample.Text = dgData.Rows[e.RowIndex].Cells[1].Value.ToString();
                cmbSampleType.Text = dgData.Rows[e.RowIndex].Cells[2].Value.ToString();
                txtSamplePlace.Text = dgData.Rows[e.RowIndex].Cells[3].Value.ToString();
                cmbLithology.Text = dgData.Rows[e.RowIndex].Cells[7].Value == null ? string.Empty : dgData.Rows[e.RowIndex].Cells[7].Value.ToString();
                cmbVeinName.Text = dgData.Rows[e.RowIndex].Cells[8].Value == null ? string.Empty : dgData.Rows[e.RowIndex].Cells[8].Value.ToString();
                txtDescription.Text = dgData.Rows[e.RowIndex].Cells[9].Value.ToString();
                dtimeDate.Text = Convert.ToDateTime(dgData.Rows[e.RowIndex].Cells[10].Value).ToShortDateString();

                swActualizarRegistro = true;
                indexRegistroGrid = e.RowIndex;

                txtChId.Enabled = false;
                txtSample.Enabled = false;
            }
        }

        private void cmbSample_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                if (cmbSample.SelectedValue.ToString() != "Select an option..")
                {
                    DataTable dgSamp = (DataTable)dgData.DataSource;
                    DataRow[] myRow = dgSamp.Select(@"Sample = '" + cmbSample.SelectedValue.ToString() + "'");
                    int rowindex = dgSamp.Rows.IndexOf(myRow[0]);
                    dgData.Rows[rowindex].Selected = true;
                    dgData.CurrentCell = dgData.Rows[rowindex].Cells[1];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cmbSampleType_Leave(object sender, EventArgs e)
        {
            if (cmbSampleType.SelectedValue.ToString() != "ORIGINAL")
            {
                txtMTS.Enabled = false;
                txtMTS_Leave(null, null);
            }
            else txtMTS.Enabled = true;
        }
    }
}
