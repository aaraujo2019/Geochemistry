using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace Geochemistry
{
    public partial class frmPpal : Form
    {
        clsRf oRf = new clsRf();
        DataTable dtFormsAllowed = new DataTable();


        public frmPpal()
        {
            InitializeComponent();
        }

        private void soilToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //DataRow[] dato = dtFormsAllowed.Select("nombre_Real_Form = 'ControlSampling'");
            //if (dato.Length > 0)
            //{
            frmSoil oSoil = new frmSoil();
            oSoil.MdiParent = this;
            oSoil.Show();
            //}
            //else
            //{
            //    MessageBox.Show("Form is not allowed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void rockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow[] dato = dtFormsAllowed.Select("nombre_Real_Form = 'frmRock'");
            if (dato.Length > 0)
            {
                frmRock oRock = new frmRock();
                oRock.MdiParent = this;
                oRock.Show();
            }
            else
            {
                MessageBox.Show("Form is not allowed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void sedimentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //DataRow[] dato = dtFormsAllowed.Select("nombre_Real_Form = 'ControlSampling'");
            //if (dato.Length > 0)
            //{
            frmSediments oSed = new frmSediments();
            oSed.MdiParent = this;
            oSed.Show();
            //}
            //else
            //{
            //    MessageBox.Show("Form is not allowed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void passwordChangeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //DataRow[] dato = dtFormsAllowed.Select("nombre_Real_Form = 'ControlSampling'");
            //if (dato.Length > 0)
            //{
            frmChangeLoggin oChP= new frmChangeLoggin();
            oChP.MdiParent = this;
            oChP.Show();
            //}
            //else
            //{
            //    MessageBox.Show("Form is not allowed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void logOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void frmPpal_Load(object sender, EventArgs e)
        {


            dtFormsAllowed = oRf.getFormsByGrupo(clsRf.sIdGrupo, ConfigurationSettings.AppSettings["IDProject"].ToString());
            clsRf.dsPermisos = oRf.getFormsByGrupoAll(clsRf.sIdGrupo);

            MdiClient ctlMDI = default(MdiClient);
            foreach (Control ctl in this.Controls)
            {
                try
                {
                    ctlMDI = (MdiClient)ctl;
                    ctlMDI.BackColor = Color.White;
                }
                catch (InvalidCastException ex)
                {
                    //throw new Exception(ex.Message);
                }
            }

        }

        private void frmPpal_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void channelsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow[] dato = dtFormsAllowed.Select("nombre_Real_Form = 'frmChannels'");
            if (dato.Length > 0)
            {
                frmChannels oCh = new frmChannels();
                oCh.MdiParent = this;
                oCh.Show();
            }
            else
            {
                MessageBox.Show("Form is not allowed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

           
        }
    }
}
