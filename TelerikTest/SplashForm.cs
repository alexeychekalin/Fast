using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace TelerikTest
{
    public partial class SplashForm : CSWinFormLayeredWindow.PerPixelAlphaForm
    {
        public SplashForm()
        {
            InitializeComponent();
            this.SelectBitmap(Properties.Resources.LoadingNoticeOrig);
        }

        //Delegate for cross thread call to close
        private delegate void CloseDelegate();

        //The type of form to be displayed as the splash screen.
        private static SplashForm splashForm;

        static public void ShowSplashScreen()
        {
            // Make sure it is only launched once.    
            if (splashForm != null) return;
            splashForm = new SplashForm();
            Thread thread = new Thread(new ThreadStart(SplashForm.ShowForm));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            //Thread.Sleep(3000);
        }

        static private void ShowForm()
        {
           // if (splashForm != null) Application.Run(splashForm);
            splashForm?.Invoke(new MethodInvoker(() => { Application.Run(splashForm); } ));

        }

        static public void CloseForm()
        {
            splashForm?.Invoke(new CloseDelegate(SplashForm.CloseFormInternal));
        }

        static private void CloseFormInternal()
        {
            if (splashForm != null)
            {
                splashForm.Close();
                splashForm = null;
            };
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
