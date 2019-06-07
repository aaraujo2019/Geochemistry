using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;

namespace Geochemistry
{
    public partial class frmChangeLoggin : Form
    {
        clsRf oRf = new clsRf();

        public frmChangeLoggin()
        {
            InitializeComponent();
        }

        private void frmChangeLoggin_Load(object sender, EventArgs e)
        {

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

        private void btnAccept_Click(object sender, EventArgs e)
        {
            try
            {
                //Valida que la contraseña nueva sea la correcta
                if (txtNewPass.Text.ToString() != txtRepPass.Text.ToString())
                {
                    MessageBox.Show("Different New Password and Repeat Password", "Geochemistry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string sPwd = GetSHA1(txtNewPass.Text.ToString());
                string sPwdOld = GetSHA1(txtOldPass.Text.ToString());

                string sResp = oRf.UpdatePass(sPwdOld.ToString(),
                    sPwd.ToString(),
                    clsRf.sUser.ToString());

                MessageBox.Show(sResp);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
