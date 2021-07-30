using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using MailKit;
using MailKit.Net.Imap;
using MimeKit;
using Telerik.WinControls;
using Telerik.WinForms.RichTextEditor;
using SmtpClient = MailKit.Net.Smtp.SmtpClient;

namespace TelerikTest
{
    public partial class RadForm2 : Telerik.WinControls.UI.RadForm
    {
        private string ConnectedMail;
        private ImapClient mailclient;
        private string pass;
        private int Stat = 0;

        public RadForm2()
        {
            InitializeComponent();
        }
        public RadForm2(IMailFolder inbox, int num, string connMail, ImapClient mclient, string password, int status)
        {
            InitializeComponent();
            pass = password;
            mailclient = mclient;
            Stat = status;
            var message = inbox.GetMessage(num);
            if (status == 0)
            {
                sender.Text = message.From.OfType<MailboxAddress>().First().Address;
                ConnectedMail = connMail; ;
                subject.Text = "Re: " + message.Subject;
            }
            else
            {
                sender.Visible = false;
                radMultiColumnComboBox1.Visible = true;
                ConnectedMail = connMail;



                radMultiColumnComboBox1.DataSource = File.ReadAllLines("MyMail.txt").Select(x => new { StrValue = x }).ToList(); ;
                radMultiColumnComboBox1.Columns[0].HeaderText = "Сохраненные Адресаты";
                radMultiColumnComboBox1.Columns[0].Width = radMultiColumnComboBox1.Width-25;
            }


        }
        public RadForm2(IMailFolder inbox, int num, string connMail, ImapClient mclient, string password, int status,string TO)
        {
            InitializeComponent();
            TopMost = true;
            pass = password;
            mailclient = mclient;
            Stat = status;
            var message = inbox.GetMessage(num);
            if (status == 0)
            {
                sender.Text = TO;
                ConnectedMail = connMail; ;
               // subject.Text = "Re: " + message.Subject;
            }
            else
            {
                sender.Visible = false;
                radMultiColumnComboBox1.Visible = true;
                ConnectedMail = connMail;



                radMultiColumnComboBox1.DataSource = File.ReadAllLines("MyMail.txt").Select(x => new { StrValue = x }).ToList(); ;
                radMultiColumnComboBox1.Columns[0].HeaderText = "Сохраненные Адресаты";
                radMultiColumnComboBox1.Columns[0].Width = radMultiColumnComboBox1.Width-25;
            }


        }

        private void RadButton1_Click(object sender, EventArgs e)
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(ConnectedMail));
            var ss = this.sender.Text.Split(';');
            foreach (var d in ss)
            {
                message.To.Add(new MailboxAddress(d));
            }
          
            message.Subject = subject.Text;
            

            var builder = new BodyBuilder();
            builder.TextBody = body.Text;

            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                builder.Attachments.Add(listBox1.Items[i] + "\\" + radListView1.Items[i].Text);

            }

            message.Body = builder.ToMessageBody();
            


            using (var client = new SmtpClient())
            {
                client.Connect("smtp." + ConnectedMail.Substring(ConnectedMail.IndexOf("@") + 1), 587, false);


                // Note: only needed if the SMTP server requires authentication
                client.Authenticate(ConnectedMail, pass);

                client.Send(message);
                client.Disconnect(true);
            }

            mailclient.GetFolder(SpecialFolder.Sent).Append(FormatOptions.Default, message);

            var s = File.ReadAllLines("MyMail.txt");
            if (!s.Contains(this.sender.Text))
            {
                File.AppendAllText("Mymail.txt", this.sender.Text + "\n");
            }

            MessageBox.Show("Письмо отправлено");
            Close();
        }

        private void Body_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);


            foreach (var s in files)
            {
                radListView1.Items.Add(s.Substring(s.LastIndexOf("\\") + 1));
                listBox1.Items.Add(s.Substring(0, s.LastIndexOf("\\")));
            }
        }

        private void Body_DragEnter(object sender, DragEventArgs e)
        {
            /*string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);


            foreach (var s in files)
            {
                radListView1.Items.Add(s.Substring(s.LastIndexOf("\\") + 1));
                listBox1.Items.Add(s.Substring(0, s.LastIndexOf("\\")));
            }*/
            e.Effect = DragDropEffects.All;

        }

        private void RadMultiColumnComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.sender.Text = radMultiColumnComboBox1.Text;
        }

        private void RadListView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);


            foreach (var s in files)
            {
                radListView1.Items.Add(s.Substring(s.LastIndexOf("\\") + 1));
                listBox1.Items.Add(s.Substring(0, s.LastIndexOf("\\")));
            }
        }

        private void RadListView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }
    }
}
