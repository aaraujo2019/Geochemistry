using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Configuration;
using System.Diagnostics;

namespace Geochemistry
{
    public partial class frmLogin : Form
    {

        clsRf oRf = new clsRf();
        bool bAct = false;

        public frmLogin()
        {
            InitializeComponent();
        }

        private void btnAceptar_Click(object sender, EventArgs e)
        {
            try
            {
                //MessageBox.Show(GetSHA1(txtPwd.Text.ToString()));

                string sPwd = GetSHA1(txtPwd.Text.ToString());
                //DataTable dtRfWorker = new DataTable();
                //dtRfWorker = oRf.getRfWorkerCred(txtUser.Text.ToString(), sPwd.ToString());
                //if (dtRfWorker.Rows.Count > 0)
                //if (txtUser.Text.ToString() == "" && txtPwd.Text.ToString() == "")

                DataTable dtUser = new DataTable();
                dtUser = oRf.getUsersPortal(txtUser.Text.ToString());
                if (dtUser.Rows.Count > 0)
                {

                    if (bool.Parse(dtUser.Rows[0]["activo_User"].ToString()) == false)
                    {
                        MessageBox.Show("Disabled User", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (dtUser.Rows[0]["login_User"].ToString().ToUpper() == txtUser.Text.ToString().ToUpper() &&
                        dtUser.Rows[0]["passwd_User"].ToString().ToUpper() == sPwd.ToString().ToUpper())
                    {

                        clsRf.sUser = txtUser.Text.ToString();
                        //clsRf.sIdentification = dtRfWorker.Rows[0]["Identification"].ToString();
                        clsRf.sIdentification = dtUser.Rows[0]["id_User"].ToString();
                        clsRf.sIdGrupo = dtUser.Rows[0]["idGrupo_User"].ToString();

                        //FrmPpal oPpal = new FrmPpal();
                        //oPpal.Show();
                        frmSplash oSplash = new frmSplash();
                        oSplash.Show();
                        this.Hide();
                        //this.Dispose();
                    }
                    else
                    {
                        MessageBox.Show("Credentials failed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Credentials failed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        public static string GetSHA1(String texto)
        {
            try
            {
                SHA1 sha1 = SHA1CryptoServiceProvider.Create();
                Byte[] textOriginal = ASCIIEncoding.Default.GetBytes(texto);
                Byte[] hash = sha1.ComputeHash(textOriginal);
                StringBuilder cadena = new StringBuilder();
                foreach (byte i in hash)
                {
                    cadena.AppendFormat("{0:x2}", i);
                }
                return cadena.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            try
            {
                //ConfigurationSettings.AppSettings["IDProject"].ToString();
                oRf.iIdProject = int.Parse(ConfigurationSettings.AppSettings["IDProject"].ToString());
                DataTable dtVers = oRf.getVersionProject();
                if (double.Parse(dtVers.Rows[0]["version"].ToString()) >
                    double.Parse(ConfigurationSettings.AppSettings["Version"].ToString()))
                {
                    bAct = true;
                    MessageBox.Show("Actualizar Versión");
                    this.Close();
                    Process.Start(@"\\mdesvrfs01\Publica\Aplicaciones\Actualizaciones\Geochemistry.bat");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void frmLogin_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (bAct == true)
            {
                bAct = false;
            }
        }
    }
}
