using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using FoxLearn.License;

namespace TelerikTest
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            SplashForm.ShowSplashScreen();



            // LICENSE CHECK
            /*
             var id = ComputerInfo.GetComputerId();
            KeyManager km = new KeyManager(id);
            LicenseInfo lic = new LicenseInfo();
            //Get license information from license file
            int value = km.LoadSuretyFile(string.Format(@"{0}\Key.lic", Application.StartupPath), ref lic);
            string productKey = lic.ProductKey;
            //Check valid
            if (km.ValidKey(ref productKey))
            {
                RadForm1 mainForm = new RadForm1(); //this takes ages
                SplashForm.CloseForm();
                Application.Run(mainForm);
               // Application.Run(new RadForm1());
            }
            else
            {
                Registration mainForm = new Registration(); //this takes ages
                SplashForm.CloseForm();
                Application.Run(mainForm);
                //Application.Run(new Registration());
            }
            */
            //END LICENSE CHECK
            RadForm1 mainForm = new RadForm1(); //this takes ages
            SplashForm.CloseForm();
            Application.Run(mainForm);
        }
    }
}
