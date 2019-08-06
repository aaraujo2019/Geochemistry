using Geochemistry.Emun;
using System;
using System.Data;
using System.Windows.Forms;

namespace Geochemistry
{
    public partial class frmMinningGeology : Form
    {
        private clsRf oRf = new clsRf();
        private bool swCombo = false;

        public frmMinningGeology()
        {
            InitializeComponent();
        }

        private void MinningGeology_Load(object sender, EventArgs e)
        {
            Loadcmb();

            dgData.DataSource = null;
            dgData.Columns.Add("chId", "Channel");
            dgData.Columns.Add("Sample", "Num. Samples");
            dgData.Columns.Add("SampleType", "Sample Type");
            dgData.Columns.Add("MTS", "MTS");
            dgData.Columns.Add("From", "From");
            dgData.Columns.Add("To", "To");
            dgData.Columns.Add("LRock", "Lithology");
            dgData.Columns.Add("VVeinName", "Cutting Width (m)");
            dgData.Columns.Add("LRockObservations", "Description");
        }

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
            cmbGeologist.SelectedValue = -1;


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
            cmbLithology.SelectedValue = -1;


            DataRow drChn = dtSampleT.Tables[3].NewRow();
            drChn[0] = "-1";
            drChn[1] = "Select an option..";
            dtSampleT.Tables[3].Rows.Add(drChn);
            cmbChannelType.DisplayMember = "Comb";
            cmbChannelType.ValueMember = "Code";
            cmbChannelType.DataSource = dtSampleT.Tables[3];
            cmbChannelType.SelectedValue = -1;
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

        private void CleanControls()
        {
            if (dgData.DataSource != null)
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
                dgData.Columns.Add("LRockObservations", "Description");
            }
            else
            {
                dgData.Rows.Clear();
            }
        }

        private void btnAgregate_Click(object sender, EventArgs e)
        {
            if (ValidarValores())
            {
                return;
            }

            dgData.Rows.Add(txtChId.Text, txtSample.Text, cmbSampleType.SelectedValue, txtMTS.Text, txtFrom.Text, txtTo.Text, cmbLithology.SelectedValue, cmbVeinName.SelectedValue, txtDescription.Text);
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


            if (Convert.ToInt32(cmbChannelType.SelectedValue) == -1)
            {
                MessageBox.Show("You must select the sampling instrument to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbMineEntrance.Focus();
                valor = true;
            }


            if (!txtChId.Text.Contains("ES_MI") || !txtChId.Text.Contains("PV_MI") || txtChId.Text.Contains("SK_MI"))
            {
                MessageBox.Show("You must enter the channel identifier to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbMineEntrance.Focus();
                valor = true;
            }

            var valCanal = txtChId.Text.Substring(5, txtChId.TextLength);

            if (txtChId.Text.Length <= 5)
            {
                MessageBox.Show("You must enter the channel identifier to continue.", "Minning Geology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbMineEntrance.Focus();
                valor = true;
            }
            
            return valor;
        }


        private void cmbMineEntrance_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(cmbMineEntrance.SelectedValue) != -1 && swCombo)
            {
                switch (Convert.ToInt32(cmbMineEntrance.SelectedValue))
                {
                    case (int)Mines.ES:
                        txtChId.Text = string.Concat(Mines.ES.ToString(), "_MI");
                        txtChId.Focus();
                        break;

                    case (int)Mines.SK:
                        txtChId.Text = string.Concat(Mines.SK.ToString(), "_MI");
                        txtChId.Focus();
                        break;

                    case (int)Mines.PV:
                        txtChId.Text = string.Concat(Mines.PV.ToString(), "_MI");
                        txtChId.Focus();
                        break;
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            
        }
    }
}
