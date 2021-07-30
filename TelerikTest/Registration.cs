using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using FoxLearn.License;

namespace TelerikTest
{
    public partial class Registration : Telerik.WinControls.UI.RadForm
    {
        const int ProductCode = 1;
        public Registration()
        {
            InitializeComponent();
            id.Text = ComputerInfo.GetComputerId();
        }

        private void activate_Click(object sender, EventArgs e)
        {
            KeyManager km = new KeyManager(id.Text);
            string productKey = key.Text;
            //Check valid
            if (km.ValidKey(ref productKey))
            {
                KeyValuesClass kv = new KeyValuesClass();
                //Decrypt license key
                if (km.DisassembleKey(productKey, ref kv))
                {
                    LicenseInfo lic = new LicenseInfo();
                    lic.ProductKey = productKey;
                    lic.FullName = "FastOffice";
                    //Save license key to file
                    km.SaveSuretyFile(string.Format(@"{0}\Key.lic", Application.StartupPath), lic);
                    MessageBox.Show(@"Вы зарегистрировали свою копию программы.", @"Активация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Application.Restart();
                }
            }
            else
                MessageBox.Show(@"Ключ активации недействителен", @"Активация", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void key_TextChanged(object sender, EventArgs e)
        {
            activate.Enabled = true;
        }

        private void Registration_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            this.WindowState = FormWindowState.Normal;
            this.Focus(); this.Show();
        }
    }
}
