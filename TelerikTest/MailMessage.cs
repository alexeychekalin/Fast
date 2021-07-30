﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using MailKit;
using MailKit.Net.Imap;
using MimeKit;
using MimeKit.Text;
using Telerik.WinControls;

namespace TelerikTest
{
    public partial class MailMessage : Telerik.WinControls.UI.RadForm
    {
        private IMailFolder inbox;
        private int Numbermessage;
        private string ConnectedMail;
        private ImapClient mclient;
        private string pass;
        public MailMessage(ImapClient mailclient, int num, string ConnMAil, string password)
        {
            InitializeComponent();
            pass = password;
            mclient = mailclient;
            ConnectedMail = ConnMAil;
            Numbermessage = num;
            inbox = mailclient.Inbox;
            MimeMessage message;
            try
            {
                 message = inbox.GetMessage(num);
            }
            catch
            {
                inbox = mailclient.GetFolder(SpecialFolder.Sent);
                try
                {
                     message = inbox.GetMessage(num);
                }
                catch (Exception e)
                {
                    RadMessageBox.Show(e.Message);
                    return;
                }
            }
           

            subject.Text = message.Subject;
            try
            {
                body.Text = message.GetTextBody(TextFormat.Plain);
            }
            catch
            {
                File.WriteAllText("temp.html", message.HtmlBody, Encoding.UTF8);
                WebBrowser wb = new WebBrowser();
                var html = Application.StartupPath + @"\temp.html";
                wb.Navigate(html);
                while (wb.ReadyState != WebBrowserReadyState.Complete)
                {
                    Thread.Sleep(200);
                    Application.DoEvents();
                }

                wb.Document.ExecCommand("SelectAll", false, null);
                wb.Document.ExecCommand("Copy", false, null);
                body.Paste();
            }

            sender.Text = message.From.OfType<MailboxAddress>().First().Address;


            foreach (var att in message.Attachments)
            {
                //radListView1.Items.Add(att.ContentType.Name);

                listView1.Items.Add(att.ContentType.Name);

            }


        }

        private void RadListView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
              //  radContextMenu1.Show(radListView1, new Point(e.X, e.Y));
            }
        }

        private void radMenuItem1_Click(object sender, System.EventArgs e)
        {
            var message = inbox.GetMessage(Numbermessage);

            foreach (var attachment in message.Attachments)
            {
                if (listView1.SelectedItems[0].Text != attachment.ContentType.Name)
                    continue;
                FolderBrowserDialog fb = new FolderBrowserDialog();
                if (fb.ShowDialog() == DialogResult.OK)
                {
                    using (var stream = File.Create(fb.SelectedPath + "\\" + attachment.ContentType.Name))
                    {
                        if (attachment is MessagePart)
                        {
                            var part = (MessagePart)attachment;

                            part.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment;

                            part.Content.DecodeTo(stream);
                        }
                    }
                }

            }
        }

        private void listView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                radContextMenu1.Show(listView1, new Point(e.X, e.Y));
            }
            else
            {
                string s = listView1.SelectedItems.ToString();
                
            }
        }

        private void listView1_ItemDrag(object sender, ItemDragEventArgs e)
        {
            DoDragDrop(e.Item, DragDropEffects.All);
        }

        private void listView1_DragLeave(object sender, EventArgs e)
        {
            DragDropEffects dde1 = DoDragDrop(listView1.SelectedItems.ToString(), DragDropEffects.All);
            //Close();
        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            var message = inbox.GetMessage(Numbermessage);

            foreach (var attachment in message.Attachments)
            {
                if (listView1.SelectedItems[0].Text != attachment.ContentType.Name)
                    continue;

                using (var stream = File.Create("tempfordrop" + attachment.ContentType.Name.Substring(attachment.ContentType.Name.LastIndexOf("."))))
                {
                    if (attachment is MessagePart)
                    {
                        var part = (MessagePart)attachment;

                        part.Message.WriteTo(stream);
                    }
                    else
                    {
                        var part = (MimePart)attachment;

                        part.Content.DecodeTo(stream);
                    }
                }
            }
            DragDropEffects dde1 = DoDragDrop(listView1.SelectedItems.ToString(), DragDropEffects.All);
        }

        private void RadButton1_Click(object sender, EventArgs e)
        {
            var r = new RadForm2(inbox, Numbermessage, ConnectedMail, mclient,pass,0);
            r.ShowDialog();
        }

        private void ListView1_DragEnter(object sender, DragEventArgs e)
        {
           

        }
    }
}
