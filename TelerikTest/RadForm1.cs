using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Win32;
using Telerik.WinControls.UI;
using Telerik.WinForms.Documents.FormatProviders.OpenXml.Docx;
using Telerik.WinForms.Documents.Model;
using TelerikTest.Properties;
using Awesomium.Windows.Forms;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using b2xtranslator.WordprocessingMLMapping;
using MailKit;
using MailKit.Net.Imap;
using Microsoft.Office.Interop.Word;
using MimeKit;
using Telerik.WinControls;
using Telerik.WinControls.Enumerations;
using static b2xtranslator.OpenXmlLib.OpenXmlPackage;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.WinForms.Documents.FormatProviders.Rtf;
using Telerik.WinForms.Documents.FormatProviders.Txt;
using Telerik.WinForms.Documents.RichTextBoxCommands;
using Telerik.WinForms.RichTextEditor;
using Telerik.WinForms.Spreadsheet;
using Point = System.Drawing.Point;
using Rectangle = System.Drawing.Rectangle;
using Word = Microsoft.Office.Interop.Word;
using IDataObject_Com = System.Runtime.InteropServices.ComTypes.IDataObject;


namespace TelerikTest
{


    public partial class RadForm1 : Telerik.WinControls.UI.RadForm
    {
        private int[] change = new int[100];
        private List<string> _foldersList = new List<string>();
        private List<string> _filesList = new List<string>();
        private string ConnectedMail = "";

        Thread receiveThread;
        static string userName;
        static string key;
        private const string host = "95.165.142.183";
        private const int port = 8888;
        static TcpClient client;
        static NetworkStream stream;

        private ImapClient mailclient = new ImapClient();
        private string pass = "";

        private readonly string[] _formats = new[]
        {
            "rtf", "txt", "doc", "docm", "docx", "dot", "dotm", "dotx", "htm", "html", "mht", "mhtml", "odt", "xls",
            "xlsx", "xlsm", "pdf"
        };

        public RadForm1()
        {
            RichTextBoxLocalizationProvider.CurrentProvider = RichTextBoxLocalizationProvider.FromFile(@"localization.word.xml");
            SpreadsheetLocalizationProvider.CurrentProvider = SpreadsheetLocalizationProvider.FromFile(@"localization.excel.xml");
            InitializeComponent();
            imageList3.Images.Add(new Bitmap("вложение.png"));
           //Подтягиваем сохраненные email
           var s = File.ReadAllLines("SavedMail.txt");
            try
            {
                loginbox.DataSource = s.Select(x => new { StrValue = x.Substring(0, x.IndexOf(" ")) }).ToList(); ;
                loginbox.Columns[0].HeaderText = "Мои аккаунты";
                
                loginbox.Columns[0].Width = loginbox.Width-25;
                loginbox.Text = "";
                passbox.Text = "";
            }
            catch
            {

            }
            //Подтягиваем ключи
            CheckForIllegalCrossThreadCalls = false;
            s = File.ReadAllLines("keys.txt");
            try
            {
                keybox.DataSource = s.Select(x => new { StrValue = x }).ToList(); ;
                keybox.Columns[0].HeaderText = "Сохраненные ключи";
                keybox.Columns[0].Width = keybox.Width;
                keybox.Text = "";
            }
            catch
            {

            }
            // ДОБАВЛЕНИЕ ПАПКИ ЗАГРУЗКИ
            _foldersList.Add(Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString());
            imageList1.Images.Add(Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString(), Resources.download);
            foldersPanel.Items.Add("Загрузки", imageList1.Images.Count - 1);
            foldersPanel.Items[0].Tag = "dowload";
            // END ДОБАВЛЕНИЕ ПАПКИ ЗАГРУЗКИ
            // 
            // восстановление папок
            var folders = File.ReadAllLines(@"foldersPanel.txt");
            foreach (var folder in folders)
            {
                if (!Directory.Exists(folder))
                {
                    //var f = new FileInfo(file);
                    imageList1.Images.Add(folder, ChangeOpacity(Resources.folder, 0.7f));
                }
                else
                {
                    imageList1.Images.Add(folder, Resources.folder);
                }
                _foldersList.Add(folder);
                foldersPanel.Items.Add(folder.Split(Path.DirectorySeparatorChar).Last(), imageList1.Images.Count - 1);
            }
            // end восстановление папок

            // ДОБАВЛЕНИЕ КОРЗИНЫ
            _filesList.Add("Trash");
            imageList2.Images.Add(Resources.trash_files);
            filesPanel.Items.Add("Корзина", imageList2.Images.Count - 1);
            filesPanel.Items[0].Tag = "trash";
            // END ДОБАВЛЕНИЕ КОРЗИНЫ 

            // восстановление файлов
            var files = File.ReadAllLines(@"filesPanel.txt");
            foreach (var file in files)
            {
                if (!File.Exists(file))
                {
                    var f = new FileInfo(file);
                    imageList2.Images.Add(file, (Image)Image.FromFile(@"imgFile\" + Path.GetExtension(f.Name).Remove(0, 1) + ".ico").Clone());

                    imageList2.Images[imageList2.Images.Count-1] =
                       ChangeOpacity(
                            imageList2.Images[imageList2.Images.Count - 1], 0.7f);
                    //imageList2.Images[imageList2.Images.Count - 1].Tag = "deleted";
                }
                else
                {
                    imageList2.Images.Add(file, (Icon)Icon.ExtractAssociatedIcon(file).Clone());
                }
                FileInfo fi = new FileInfo(file);
                _filesList.Add(fi.FullName);
                filesPanel.Items.Add(Path.GetFileNameWithoutExtension(fi.Name), imageList2.Images.Count - 1);
            }

            // end восстановление файлов
           
        }

        // FOLDER панель
        private void foldersPanel_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (Directory.Exists(files[0]))
                {
                    imageList1.Images.Add(files[0], Resources.folder);
                    _foldersList.Add(files[0]);
                    foldersPanel.Items.Add(files[0].Split(Path.DirectorySeparatorChar).Last(), imageList1.Images.Count-1);
                    File.AppendAllText(@"foldersPanel.txt", files[0] + System.Environment.NewLine);
                }
            }
            else
            {
                var pos = foldersPanel.PointToClient(new Point(e.X, e.Y));
                var hit = foldersPanel.HitTest(pos);
                var dragItem = (ListViewItem) e.Data.GetData(typeof(ListViewItem));
                if (dragItem.ListView.Name == "listView1")
                {
                    File.Copy("tempfordrop" + dragItem.Text.Substring(dragItem.Text.LastIndexOf(".")),
                        _foldersList[hit.Item.Index] + @"\" + dragItem.Text, true);
                }
            }
        }
        private void foldersPanel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.Move;
            }
        }
        private void foldersPanel_ItemDrag(object sender, ItemDragEventArgs e)
        {
            DoDragDrop(e.Item, DragDropEffects.All);
        }
        private void foldersPanel_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (Directory.Exists(_foldersList[foldersPanel.SelectedItems[0].Index]))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    Arguments = _foldersList[foldersPanel.SelectedItems[0].Index],
                    FileName = "explorer.exe"
                };

                Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show($@"Директории ""{_foldersList[foldersPanel.SelectedItems[0].Index]}"" не сущетсвует!");
            }
        }
        private void foldersPanel_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected)
            {
                pathToFileFolder.Text = _foldersList[foldersPanel.SelectedItems[0].Index];
                nameFileFolder.Text = foldersPanel.SelectedItems[0].Text;
                if (!Directory.Exists(_foldersList[foldersPanel.FocusedItem.Index]))
                {
                    imageList1.Images[imageList1.Images.IndexOfKey(_foldersList[foldersPanel.FocusedItem.Index])] =
                        ChangeOpacity(
                            imageList1.Images[imageList1.Images.IndexOfKey(_foldersList[foldersPanel.FocusedItem.Index])], 0.7f);
                }
            }
            else
            {
                pathToFileFolder.Text = "";
                nameFileFolder.Text = "";
            }

        }
        // END FOLDER панель

        // FILES панель
        private void filesPanel_DragDrop(object sender, DragEventArgs e)
        {
            string[] file = (string[])e.Data.GetData(DataFormats.FileDrop);
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                if (File.Exists(file[0]) && _formats.Any(Path.GetExtension(file[0]).ToLower().Contains))
                {
                    radWaitingBar2.AssociatedControl = filesPanel;
                    radWaitingBar2.StartWaiting();
                    switch (Path.GetExtension(file[0])?.ToLower())
                    {
                        case ".doc":
                            StructuredStorageReader reader = new StructuredStorageReader(file[0]);
                            WordDocument doc = new WordDocument(reader);
                            WordprocessingDocument docx = WordprocessingDocument.Create(file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx", DocumentType.Document);
                            Converter.Convert(doc, docx);
                            file[0] = file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx";
                            break;
                        case ".docm":
                        case ".dot":
                        case ".dotm":
                        case ".dotx":
                        case ".odt":
                            Convert(file[0], file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx",
                                WdSaveFormat.wdFormatDocumentDefault);
                            file[0] = file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx";
                            break;
                        
                        case ".xls":
                            var xlsConverter = new XlsToXlsx();
                            xlsConverter.ConvertToXlsxFile(file[0], file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").xlsx");
                            file[0] = file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").xlsx";
                            break;
                    }

                    imageList2.Images.Add(file[0], Icon.ExtractAssociatedIcon(file[0]));
                    FileInfo fi = new FileInfo(file[0]);
                    _filesList.Add( fi.FullName);
                    filesPanel.Items.Add(Path.GetFileNameWithoutExtension(fi.Name), imageList2.Images.Count - 1);

                    File.AppendAllText(@"filesPanel.txt", fi.FullName + System.Environment.NewLine);

                    var img = imageList2.Images[file[0]];
                    img.Save(@"imgFile\"+Path.GetExtension(fi.Name).Remove(0,1)+".ico");

                    radWaitingBar2.StopWaiting();
                    radWaitingBar2.AssociatedControl = null;
                }
            }
            else if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                // //УДАЛЕНИЕ item
                var pos = filesPanel.PointToClient(new Point(e.X, e.Y));
                var hit = filesPanel.HitTest(pos);
                if (hit.Item != null && hit.Item.Tag != null)
                {
                    var dragItem = (ListViewItem)e.Data.GetData(typeof(ListViewItem));
                    if (dragItem.ListView.Name == "foldersPanel")
                    {
                        foreach (ListViewItem deleted in foldersPanel.SelectedItems)
                        {
                            var del = _foldersList[deleted.Index];
                            _foldersList.RemoveAt(deleted.Index);
                            foldersPanel.Items.RemoveAt(deleted.Index);
                            File.WriteAllLines("foldersPanel.txt", File.ReadLines("foldersPanel.txt").Where(l => l != del).ToList());
                        }
                    }
                    else
                    {
                        foreach (ListViewItem deletedf in filesPanel.SelectedItems)
                        {
                            var del = _filesList[deletedf.Index];
                            _filesList.RemoveAt(deletedf.Index);
                            filesPanel.Items.RemoveAt(deletedf.Index);
                            File.WriteAllLines("filesPanel.txt", File.ReadLines("filesPanel.txt").Where(l => l != del).ToList());
                        }
                        //var deleted = _filesList[filesPanel.SelectedItems[0].Index];
                        //_filesList.RemoveAt(filesPanel.SelectedItems[0].Index);
                        //filesPanel.Items.RemoveAt(filesPanel.SelectedItems[0].Index);

                    }

                }
            }
            radWaitingBar2.StopWaiting();
            // filesPanel.Items[0].Position = new Point(0, 2);
            // filesPanel.Items[0].Position = new Point(filesPanel.Width - 21, 2);
        }
        private void filesPanel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        private void filesPanel_ItemDrag(object sender, ItemDragEventArgs e)
        {
            DoDragDrop(e.Item, DragDropEffects.All);
        }

        private void filesPanel_DoubleClick(object sender, EventArgs e)
        {
            if (File.Exists(_filesList[filesPanel.SelectedItems[0].Index]))
            {
                Process.Start(_filesList[filesPanel.SelectedItems[0].Index]);
            }
            else if (filesPanel.SelectedItems[0].Tag != null)
            {
                // do nothing
            }
            else
            {
                MessageBox.Show($"Файл \"{_filesList[filesPanel.SelectedItems[0].Index]}\" удалён или перемещён!");
            }
        }
        private void filesPanel_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected)
            {
                pathToFileFolder.Text = _filesList[filesPanel.SelectedItems[0].Index];
                nameFileFolder.Text = filesPanel.SelectedItems[0].Text;
                if (!File.Exists(_filesList[filesPanel.SelectedItems[0].Index]))
                {
                    imageList2.Images[imageList2.Images.IndexOfKey(_filesList[filesPanel.FocusedItem.Index])] =
                        ChangeOpacity(
                            imageList2.Images[imageList2.Images.IndexOfKey(_filesList[filesPanel.FocusedItem.Index])], 0.7f);
                   // imageList2.Images[_filesList[filesPanel.SelectedItems[0].Index]].Tag =
                   //     "deleted";
                  //  imageList2.Images[3].Tag = (object)@"fast";
                }
            }
            else
            {
                pathToFileFolder.Text = "";
                nameFileFolder.Text = "";
            }
        }
        // END FILES панель

        // MAIN WINDOW
        private void radPageView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] file = (string[])e.Data.GetData(DataFormats.FileDrop);
           // var dragItem = (ListViewItem)e.Data.GetData(typeof(ListViewItem));
            var dragItem = (ListViewItem)e.Data.GetData(typeof(ListViewItem));

            //if (!e.Data.GetDataPresent(typeof(ListViewItem))) return;

            // radWaitingBar2.StartWaiting();
            try
            {
               if (dragItem.ListView.Name == "filesPanel" || dragItem.ListView.Name == "listView1")
                {
                    if (dragItem.ListView.Name == "listView1")
                    {
                        File.Copy("tempfordrop" + dragItem.Text.Substring(dragItem.Text.LastIndexOf(".")), _foldersList[0] + @"\" + dragItem.Text, true);
                        var newFile = _foldersList[0] + @"\" + dragItem.Text;
                        try
                        {
                            using (FileStream stream = File.OpenRead(newFile))
                            {
                                stream.Close();
                            }
                        }
                        catch (IOException)
                        {
                            MessageBox.Show(@"Файл занят другим приложением, закройте файл и повторите попытку", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (File.Exists(newFile) && _formats.Any(Path.GetExtension(newFile).ToLower().Contains))
                        {
                            radWaitingBar2.AssociatedControl = filesPanel;
                            radWaitingBar2.StartWaiting();
                            switch (Path.GetExtension(newFile)?.ToLower())
                            {
                                case ".doc":
                                    StructuredStorageReader reader = new StructuredStorageReader(file[0]);
                                    WordDocument doc = new WordDocument(reader);
                                    WordprocessingDocument docx = WordprocessingDocument.Create(file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx", DocumentType.Document);
                                    Converter.Convert(doc, docx);
                                    newFile = newFile.Remove(newFile.LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx";
                                    break;
                                case ".docm":
                                case ".dot":
                                case ".dotm":
                                case ".dotx":
                                case ".odt":
                                    Convert(newFile, newFile.Remove(newFile.LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx",
                                        WdSaveFormat.wdFormatDocumentDefault);
                                    newFile = newFile.Remove(newFile.LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx";
                                    break;

                                case ".xls":
                                    var xlsConverter = new XlsToXlsx();
                                    xlsConverter.ConvertToXlsxFile(newFile, newFile.Remove(newFile.LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").xlsx");
                                    newFile = newFile.Remove(newFile.LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").xlsx";
                                    break;
                            }

                            imageList2.Images.Add(newFile, Icon.ExtractAssociatedIcon(newFile));
                            FileInfo fi = new FileInfo(newFile);
                            _filesList.Add(fi.FullName);
                            filesPanel.Items.Add(Path.GetFileNameWithoutExtension(fi.Name), imageList2.Images.Count - 1);

                            File.AppendAllText(@"filesPanel.txt", fi.FullName + System.Environment.NewLine);

                            var img = imageList2.Images[newFile];
                            img.Save(@"imgFile\" + Path.GetExtension(fi.Name).Remove(0, 1) + ".ico");

                            radWaitingBar2.StopWaiting();
                            radWaitingBar2.AssociatedControl = null;
                        }
                    }
                    var filePath = dragItem.ListView.Name == "filesPanel" ? _filesList[dragItem.Index] : _filesList[_filesList.Count - 1];
                    if (!File.Exists(filePath))
                    {
                        try
                        {
                            imageList2.Images[imageList2.Images.IndexOfKey(_filesList[filesPanel.SelectedItems[0].Index])] =
                                ChangeOpacity(
                                    imageList2.Images[imageList2.Images.IndexOfKey(_filesList[filesPanel.SelectedItems[0].Index])], 0.7f);
                            //imageList2.Images[imageList2.Images.IndexOfKey(_filesList[filesPanel.SelectedItems[0].Index])].Tag =
                            //   "deleted";
                            return;
                        }
                        catch (Exception exception)
                        {
                           // Console.WriteLine(exception);
                           // throw;
                        }
                        
                    }
                    //radPageView1.Pages.(dragItem.Text);

                    StartView(filePath, dragItem.Text);
                }
            }
           catch (Exception exception)
           {
               if (_formats.Any(Path.GetExtension(file[0]).ToLower().Contains))
               {
                   switch (Path.GetExtension(file[0])?.ToLower())
                   {
                       case ".doc":
                           StructuredStorageReader reader = new StructuredStorageReader(file[0]);
                           WordDocument doc = new WordDocument(reader);
                           WordprocessingDocument docx = WordprocessingDocument.Create(file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx", DocumentType.Document);
                           Converter.Convert(doc, docx);
                           file[0] = file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx";
                            break;
                       case ".docm":
                       case ".dot":
                       case ".dotm":
                       case ".dotx":
                       case ".odt":
                           Convert(file[0], file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx",
                               WdSaveFormat.wdFormatDocumentDefault);
                           file[0] = file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").docx";
                            break;

                       case ".xls":
                           var xlsConverter = new XlsToXlsx();
                           xlsConverter.ConvertToXlsxFile(file[0], file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").xlsx");
                           file[0] = file[0].Remove(file[0].LastIndexOf(".", StringComparison.Ordinal)) + "(" + file[0].Split('.')[1] + ").xlsx";
                            break;
                   }
                    StartView(file[0], file[0].Split('\\').Last().Split('.')[0]);
               }
           }
        }

        private void StartView(string filePath, string name)
        {
            try
            {
                using (FileStream stream = File.OpenRead(filePath))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                MessageBox.Show(@"Файл занят другим приложением, закройте файл и повторите попытку", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var page = new RadPageViewPage(name);
            radPageView1.Pages.Add(page);
            radPageView1.SelectedPage = page;
            page.KeyUp += new KeyEventHandler(radPageView1_KeyDown);

            radWaitingBar2.AssociatedControl = page;
            radWaitingBar2.StartWaiting();
            switch (Path.GetExtension(filePath).ToLower())
            {
                case ".docx":
                case ".txt":
                case ".rtf":
                    var editor = new RadRichTextEditor();
                    editor.LayoutMode = DocumentLayoutMode.Paged;
                    editor.ThemeName = "visualStudio2012LightTheme1";
                    editor.KeyUp += new KeyEventHandler(radPageView1_KeyDown);

                    switch (Path.GetExtension(filePath).ToLower())
                    {
                        case ".docx":
                            var providerd = new DocxFormatProvider();
                            using (FileStream inputStream = File.OpenRead(filePath))
                            {
                                editor.Document = providerd.Import(inputStream);
                            }
                            break;
                        case ".txt":
                            string encoding;
                            using (var reader = new StreamReader(filePath, Encoding.Default))
                            {
                               // reader.Peek(); // you need this!
                                //encoding = reader.CurrentEncoding.BodyName;
                                encoding = reader.ReadToEnd();
                            }
                            var providert = new TxtFormatProvider();
                           // using (Stream inputStream = File.OpenRead(filePath))
                           // {
                                editor.Document = providert.Import(encoding);
                           // }
                            break;
                        case ".rtf":
                            var providerr = new RtfFormatProvider();
                            using (Stream inputStream = File.OpenRead(filePath))
                            {
                                editor.Document = providerr.Import(inputStream);
                            }
                            break;
                    }

                    var ruler = new RadRichTextEditorRuler();
                    ruler.Dock = DockStyle.Fill;
                    ruler.AssociatedRichTextBox = editor;
                    ruler.Controls.Add(editor);
                    ruler.ThemeName = "visualStudio2012LightTheme1";

                    var ribbon = new CustomRichTextEditorRibbonBar();
                    ribbon.AssociatedRichTextEditor = editor;
                    ribbon.OpenedFileName = filePath;
                    ribbon.Controls[1].Visible = false;
                    ribbon.Text = Path.GetFileName(filePath);
                    ((Telerik.WinControls.UI.RichTextEditorRibbonUI.RichTextEditorRibbonTab)(ribbon.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(1))).Visibility = Telerik.WinControls.ElementVisibility.Visible;
                    ((Telerik.WinControls.UI.RichTextEditorRibbonUI.RichTextEditorRibbonTab)(ribbon.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(2))).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    //  ((Telerik.WinControls.UI.RichTextEditorRibbonUI.RichTextEditorRibbonTab)(ribbon.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(3))).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    ((Telerik.WinControls.UI.RichTextEditorRibbonUI.RichTextEditorRibbonTab)(ribbon.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(4))).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    ((Telerik.WinControls.UI.RichTextEditorRibbonUI.RichTextEditorRibbonTab)(ribbon.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(5))).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    ((Telerik.WinControls.UI.RichTextEditorRibbonUI.RichTextEditorRibbonTab)(ribbon.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(6))).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    ribbon.ExpandButton = new RadToggleButtonElement();
                    //ribbon.
                    page.Controls.Add(ruler);
                    page.Controls.Add(ribbon);
                    editor.DocumentContentChanged += EditorOnDocumentChanged;
                    editor.CommandExecuting += EditorOnCommandExecuting;
                    break;
                case ".htm":
                case ".html":
                case ".pdf":
                    WebControl webControl1 = new WebControl();
                    webControl1.Dock = DockStyle.Fill;
                    webControl1.Source = new Uri(filePath);
                    webControl1.KeyUp += new KeyEventHandler(radPageView1_KeyDown);
                    //radButton12.Visible = true;
                    page.Controls.Add(webControl1);
                    break;
                case ".mht":
                case ".mhtml":
                    var webBrowser = new WebBrowser();
                    webBrowser.Navigate(filePath);
                    webBrowser.Dock = DockStyle.Fill;
                    webBrowser.KeyUp += new KeyEventHandler(radPageView1_KeyDown);

                    page.Controls.Add(webBrowser);
                    break;
                case ".xlsx":
                    var excel = new RadSpreadsheet();
                    excel.Dock = DockStyle.Fill;
                    excel.ThemeName = "visualStudio2012LightTheme1";
                    excel.KeyUp += new KeyEventHandler(radPageView1_KeyDown);
                    

                    var ribonExcel = new RadSpreadsheetRibbonBar
                    {
                        AssociatedSpreadsheet = excel,
                        ThemeName = "visualStudio2012LightTheme1",
                        CloseButton = false,
                        MaximizeButton = false,
                        MinimizeButton = false,
                        LayoutMode = RibbonLayout.Simplified,

                        Text = Path.GetFileName(filePath)
                    };
                    
                    //ribonExcel.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(1).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    //ribonExcel.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(2).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    //ribonExcel.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(3).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    //ribonExcel.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(4).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    //ribonExcel.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(5).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                    //ribonExcel.GetChildAt(0).GetChildAt(4).GetChildAt(0).GetChildAt(0).GetChildAt(6).Visibility = Telerik.WinControls.ElementVisibility.Collapsed;

                    var formatProvider = new XlsxFormatProvider();
                    using (Stream input = new FileStream(filePath, FileMode.Open))
                    {
                        excel.Workbook = formatProvider.Import(input);
                    }
                    page.Controls.Add(excel);
                    page.Controls.Add(ribonExcel);
                    break;
            }

            radWaitingBar2.StopWaiting();
            radWaitingBar2.AssociatedControl = null;

            if (radPageView1.Pages.Count > 1)
            {
                foreach (var temPage in radPageView1.Pages)
                {
                    temPage.Tag = temPage.ItemSize;
                    temPage.ItemSize = new System.Drawing.SizeF(100, 26);
                    temPage.Description = temPage.Text;
                    if (temPage.Text.Length > 8 && temPage.Text.Contains("*")) temPage.Text = temPage.Text.Substring(0, 8) + @"...*";
                    else
                    {
                        if(temPage.Text.Length > 8)
                            temPage.Text = temPage.Text.Substring(0, 8) + @"...";
                    }
                }
            }
        }

        private void EditorOnCommandExecuting(object sender, CommandExecutingEventArgs e)
        {
            var t = (RadRichTextBox) sender;
            if (e.Command.ToString() == t.Commands.SaveCommand.ToString())
            {
                change[radPageView1.Pages.IndexOf(radPageView1.SelectedPage)] = 0;
                radPageView1.SelectedPage.Description = radPageView1.SelectedPage.Description.Split('*')[0];
                radPageView1.SelectedPage.Text = radPageView1.SelectedPage.Text.Split('*')[0];

            }
        }

        private void EditorOnDocumentChanged(object sender, EventArgs e)
        {
            
            var t = (RadRichTextBox) sender;

            if (change[radPageView1.Pages.IndexOf(radPageView1.SelectedPage)] == 1 && !radPageView1.SelectedPage.Text.Contains("*"))
            {
                if (radPageView1.SelectedPage.Description != "")
                { 
                    radPageView1.SelectedPage.Description = radPageView1.SelectedPage.Description + "*";
                }
                else
                {
                    radPageView1.SelectedPage.Description = radPageView1.SelectedPage.Text + "*";
                }
                radPageView1.SelectedPage.Text = radPageView1.SelectedPage.Text + @"*";
            }
            else 
            {
                change[radPageView1.Pages.IndexOf(radPageView1.SelectedPage)] = 1;
            }
            
        }

        private void radPageView1_DragEnter(object sender, DragEventArgs e)
        {
           // e.Effect = e.Data.GetDataPresent(typeof(ListViewItem)) ? DragDropEffects.Move : DragDropEffects.None;
            e.Effect = DragDropEffects.Move;
        }

        private void radPageView1_PageRemoving(object sender, RadPageViewCancelEventArgs e)
        {
            //var closePage = (RadPageViewPage) e.Page;
            change[radPageView1.Pages.IndexOf(e.Page)] = 0;

            try
            {
                for (int i = e.Page.Controls.Count - 1; i >= 0; i--)
                {
                    e.Page.Controls[i].Dispose();
                }
            }
            catch (Exception exception)
            {
                //Console.WriteLine(exception);
               // throw;
            }
           
        }
        private void radPageView1_PageRemoved(object sender, RadPageViewEventArgs e)
        {
            e.Page.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.WaitForFullGCApproach();
            GC.WaitForFullGCComplete();
            GC.Collect();

            if (radPageView1.Pages.Count == 1)
            {
                radPageView1.Pages[0].Text = radPageView1.Pages[0].Description;
                radPageView1.Pages[0].ItemSize = (System.Drawing.SizeF)radPageView1.Pages[0].Tag;
            }
        }
        //END MAIN WINDOW

        //HELPERS
        public static void Convert(string input, string output, WdSaveFormat format)
        {
            _Application oWord = new Word.Application
            {
                Visible = false
            };

            // Interop requires objects.
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = true;     // Does not cause any word dialog to show up
            //object readOnly = false;  // Causes a word object dialog to show at the end of the conversion
            object oInput = input;
            object oOutput = output;
            object oFormat = format;
            Object falseObj = false;
            _Document oDoc = null;

            try
            {
                oDoc = oWord.Documents.Open(
                    ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

                // Make this document the active document.
                oDoc.Activate();

                // Save this document using Word
                oDoc.SaveAs2(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

                // Always close Word.exe.
                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                oDoc = null;
                oWord = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(@"Ошибка конвертирования файлов", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                oDoc.Close(ref falseObj, ref oMissing, ref oMissing);
                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                oDoc = null;
                oWord = null;
                throw;
            }
            
        }
        public static Bitmap ChangeOpacity(Image img, float opacityvalue)
        {
            Bitmap bmp = new Bitmap(img.Width, img.Height); // Determining Width and Height of Source Image
            Graphics graphics = Graphics.FromImage(bmp);
            ColorMatrix colormatrix = new ColorMatrix();
            colormatrix.Matrix33 = opacityvalue;
            ImageAttributes imgAttribute = new ImageAttributes();
            imgAttribute.SetColorMatrix(colormatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
            graphics.DrawImage(img, new Rectangle(0, 0, bmp.Width, bmp.Height), 0, 0, img.Width, img.Height, GraphicsUnit.Pixel, imgAttribute);
            graphics.Dispose();  // Releasing all resource used by graphics 
            return bmp;
        }
        //END HELPERS

        //MAIL AND CHAT
        private void RadButton1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            if (keybox.Text == "")
            {
                RadMessageBox.Show("Необходимо ввести ключ");
                return;
            }
            userName = chatname.Text;

            client = new TcpClient();
            try
            {
                client.Connect(host, port); //подключение клиента

                stream = client.GetStream(); // получаем поток
                string message = userName;
                byte[] data = Encoding.Unicode.GetBytes(message);
                stream.Write(data, 0, data.Length);
                key = keybox.Text;

                // запускаем новый поток для получения данных
                receiveThread = new Thread(new ThreadStart(ReceiveMessage));
                receiveThread.Start(); //старт потока

                listBox1.Items.Add("Добро пожаловать, " + userName);

                messagebox.Text = key;
                SendMessage();
                messagebox.Text = "";
                var s = File.ReadAllLines("keys.txt");
                if (!s.Contains(key))
                    File.AppendAllText("keys.txt", key + "\n");
                radButton8.Visible = true;
                radButton1.Visible = false;

            }
            catch (Exception ex)
            {
                RadMessageBox.Show(ex.Message);
            }
        }

        void SendMessage()
        {
            string message = messagebox.Text;
            byte[] data = Encoding.Unicode.GetBytes(message);
            stream.Write(data, 0, data.Length);
        }
        
        void ReceiveMessage()
        {
            while (true)
            {
                try
                {
                    byte[] data = new byte[64]; // буфер для получаемых данных
                    StringBuilder builder = new StringBuilder();
                    int bytes = 0;
                    do
                    {
                        bytes = stream.Read(data, 0, data.Length);
                        builder.Append(Encoding.Unicode.GetString(data, 0, bytes));
                    }
                    while (stream.DataAvailable);

                    string message = builder.ToString();
                    listBox1.Items.Add(message);//вывод сообщения
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message);
                    //listBox1.Items.Add("Подключение прервано!"); //соединение было прервано
                    //Console.ReadLine();
                    Disconnect();
                }
            }
        }

        static void Disconnect()
        {
            if (stream != null)
                stream.Close();//отключение потока
            if (client != null)
            {
                client.Dispose();
                client.Close(); //отключение клиента
                client = null;
            }
        }

        private void RadButton2_Click(object sender, EventArgs e)
        {
            SendMessage();
            messagebox.Text = "";
        }

        private void RadButton4_Click(object sender, EventArgs e)
        {
            try
            {

           
            //radWaitingBar1.AssociatedControl = radListView1;
            // radWaitingBar1.StartWaiting();
            IMailFolder inbox ;
            if (radListView1.Columns[1].Text == "Отправитель")
                inbox = mailclient.Inbox;
            else
                inbox = mailclient.GetFolder(SpecialFolder.Sent);
            inbox.Open(FolderAccess.ReadOnly);

            inbox.Open(FolderAccess.ReadOnly);

            //passbox.Text = "Total messages:" + inbox.Count;
            //Console.WriteLine();
            //Console.WriteLine("Recent messages: {0}", inbox.Recent);
            radListView1.Items.Clear();

            radListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.None);
            listBox2.Items.Clear();
            radListView1.Columns[1].Text = "Отправитель";
            for (int i = inbox.Count - 1; i >= Math.Max(inbox.Count - 40, 0); i--)
            {
                var message = inbox.GetMessage(i);

                if (message.Subject == null)
                {
                    if (message.Attachments.Count() > 0)
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]{"",message.From.OfType<MailboxAddress>().First().Name,
                            " ", message.Date.ToString("dd.MM.yyyy")});
                        listViewItem.ImageIndex = 0;

                        radListView1.Items.Add(listViewItem);


                    }
                    else
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]{"",message.From.OfType<MailboxAddress>().First().Name,
                            " ", message.Date.ToString("dd.MM.yyyy")});


                        radListView1.Items.Add(listViewItem);
                    }

                }

                else
                {
                    if (message.Attachments.Count() > 0)
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]{"",message.From.OfType<MailboxAddress>().First().Name,
                            message.Subject, message.Date.ToString("dd.MM.yyyy")});
                        listViewItem.ImageIndex = 0;

                        radListView1.Items.Add(listViewItem);

                    }
                    else
                    {
                        ListViewItem listViewItem = new ListViewItem(
                        new string[]{"",message.From.OfType<MailboxAddress>().First().Name.ToString(),
                            message.Subject.ToString(), message.Date.ToString("dd.MM.yyyy")});

                        radListView1.Items.Add(listViewItem);
                    }

                }



                listBox2.Items.Add(i);



            }
            }
            catch (Exception exception)
            {
                RadMessageBox.Show(exception.Message);
            }
        }

        private void RadButton5_Click(object sender, EventArgs e)
        {
            try
            {



                RadForm2 rf = new RadForm2(mailclient.Inbox, 0, ConnectedMail, mailclient, pass, 1);
                rf.Show();
            }
            catch (Exception exception)
            {
                RadMessageBox.Show(exception.Message);
            }
        }

        private void Loginbox_TextChanged(object sender, EventArgs e)
        {
            var s = File.ReadAllLines("SavedMail.txt");
            foreach (var t in s)
            {
                if (t.Substring(0, t.IndexOf(" ")) == loginbox.Text)
                {
                    passbox.Text = t.Substring(t.IndexOf(" ") + 1);
                    break;
                }
            }
        }

        private void Disconnectbutton_Click(object sender, EventArgs e)
        {
            mailclient.Disconnect(true);
            radListView1.Items.Clear();
            listBox2.Items.Clear();
            disconnectbutton.Visible = false;
            //radButton5.Visible = false;
            //radButton6.Visible = false;
            //findbox.Visible = false;
            //radButton4.Visible = false;
            //radButton7.Visible = false;
            radButton3.Visible = true;

            //outcomeMessage.Visible = false;
            //incomeMessage.Visible = false;
            //radButton5.Visible = false;
            //button1.Visible = false;
            //button2.Visible = false;
        }

        private void RadListView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            MailMessage m = new MailMessage(mailclient, System.Convert.ToInt32(listBox2.Items[radListView1.SelectedIndices[0]]), ConnectedMail, pass);
            m.Show();
        }

        private void RadButton6_Click(object sender, EventArgs e)
        {
            if (findbox.Text != "")
            {
                var inbox = mailclient.Inbox;
                inbox.Open(FolderAccess.ReadOnly);

                radListView1.Items.Clear();
                listBox2.Items.Clear();
                for (int i = 0; i < inbox.Count; i++)
                {
                    var message = inbox.GetMessage(i);
                    if ((message.Subject != null && message.Subject.Contains(findbox.Text)) || (message.TextBody != null && message.TextBody.Contains(findbox.Text)))
                    {
                        radListView1.Items.Add(message.Subject);
                        if (message.Attachments.Count() > 0)
                        {
                            //radListView1.Items[radListView1.Items.Count - 1].Image = new Bitmap("вложение.png");
                        }

                        listBox2.Items.Add(i);
                    }
                }
            }
            else
            {
                var inbox = mailclient.Inbox;
                inbox.Open(FolderAccess.ReadOnly);

                radListView1.Items.Clear();
                listBox2.Items.Clear();
                for (int i = inbox.Count - 10; i < inbox.Count; i++)
                {
                    var message = inbox.GetMessage(i);
                    listBox2.Items.Add(i);
                    radListView1.Items.Add(message.Subject);
                    if (message.Attachments.Count() != 0)
                    {
                       // radListView1.Items[radListView1.Items.Count - 1].Image = new Bitmap("вложение.png");
                    }
                }
            }
        }

        private void RadButton7_Click(object sender, EventArgs e)
        {
            try
            {
                IMailFolder inbox;
                if(radListView1.Columns[1].Text=="Отправитель")
                 inbox = mailclient.Inbox;
                else
                    inbox = mailclient.GetFolder(SpecialFolder.Sent);

                inbox.Open(FolderAccess.ReadWrite );
                //var message = inbox.GetMessage(System.Convert.ToInt32 (radListView1.Items[radListView1.SelectedIndices[0]]));

                inbox.AddFlags(System.Convert.ToInt32( listBox2.Items[System.Convert.ToInt32(radListView1.Items[radListView1.SelectedIndices[0].ToString()])].ToString())
                    , MessageFlags.Deleted, true);

               RadButton4_Click(sender,e);
            
            }
            catch (Exception exception)
            {
                RadMessageBox.Show(exception.Message);
            }
        }

        private void RadForm1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (splitPanel2.Collapsed)
            {
                e.Cancel = true; // отменяем закрытие формы
                splitPanel2.Collapsed = false;
                radSplitContainer2.Visible = true;
            }
            else
            {
                mailclient.Disconnect(true);
            }
        }
        
        private void RadButton8_Click(object sender, EventArgs e)
        {
            receiveThread.Abort();
            receiveThread.Join(500);

            Disconnect();

            radButton1.Visible = true;
            radButton8.Visible = false;

        }

        private void radButton3_Click_1(object sender, EventArgs e)
        {
           
            

           // radWaitingBar1.StartWaiting();
            try
            {
                ConnectedMail = loginbox.Text;
                var pos = loginbox.Text.LastIndexOf("@");
                //pos = loginbox.Text.LastIndexOf(".", pos - 1);
                mailclient.Connect("imap." + loginbox.Text.Substring(pos + 1), 993, true);
                mailclient.Authenticate(loginbox.Text, passbox.Text);
                pass = passbox.Text;
            }
            catch
            {
               
                RadMessageBox.Show("Проверьте правильность ввода имя пользователя и пароля");
                return;
            }

           
            // The Inbox folder is always available on all IMAP servers...
            var inbox = mailclient.Inbox;
            inbox.Open(FolderAccess.ReadOnly);

            //passbox.Text = "Total messages:" + inbox.Count;
            //Console.WriteLine();
            //Console.WriteLine("Recent messages: {0}", inbox.Recent);
            radListView1.Items.Clear();

           
            listBox2.Items.Clear();
            
           
            for (int i = inbox.Count-1; i >= Math.Max(inbox.Count-40,0); i--)
            {
                var message = inbox.GetMessage(i);
                
                if (message.Subject==null)
                {
                    
                    if (message.Attachments.Count() > 0)
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]{"",message.From.OfType<MailboxAddress>().First().Name,
                            " ", message.Date.ToString("dd.MM.yyyy")});
                        listViewItem.ImageIndex = 0;

                        radListView1.Items.Add(listViewItem);
                       

                    }
                    else
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]
                        {
                            "", message.From.OfType<MailboxAddress>().First().Name,
                            " ", message.Date.ToString("dd.MM.yyyy")
                        });


                        radListView1.Items.Add(listViewItem);
                    }

                }
                
                else
                {
                    if (message.Attachments.Count() > 0)
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]{"",message.From.OfType<MailboxAddress>().First().Name,
                            message.Subject, message.Date.ToString("dd.MM.yyyy")});
                        listViewItem.ImageIndex = 0;

                        radListView1.Items.Add(listViewItem);
                      

                    }
                    else
                    {
                        ListViewItem listViewItem = new ListViewItem(
                        new string[]{"",message.From.OfType<MailboxAddress>().First().Name.ToString(),
                            message.Subject.ToString(), message.Date.ToString("dd.MM.yyyy")});
                      
                        radListView1.Items.Add(listViewItem);
                       
                    }

                }
              

               
                listBox2.Items.Add(i);

              
              
            }



                radListView1.Columns[0].Width = 20;
                radListView1.Columns[1].Width = 100;
                radListView1.Columns[2].Width = 120;
                radListView1.Columns[3].Width = 78;


                var s = File.ReadAllLines("SavedMail.txt");
            if (!s.Contains(loginbox.Text + " " + passbox.Text))
                File.AppendAllText("SavedMail.txt", loginbox.Text + " " + passbox.Text + "\n");
            passbox.Text = "";
            disconnectbutton.Visible = true;
            
        }

        private void OutcomeMessage_Click(object sender, EventArgs e)
        {
            try
            {


                // radWaitingBar1.AssociatedControl = radListView1;
                radListView1.Columns[1].Text = "Получатель";
                var inbox = mailclient.GetFolder(SpecialFolder.Sent);
                inbox.Open(FolderAccess.ReadOnly);



                //passbox.Text = "Total messages:" + inbox.Count;
                //Console.WriteLine();
                //Console.WriteLine("Recent messages: {0}", inbox.Recent);
                radListView1.Items.Clear();


                listBox2.Items.Clear();
                radListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.None);
                for (int i = inbox.Count - 1; i >= Math.Max(inbox.Count - 40, 0); i--)
                {
                    var message = inbox.GetMessage(i);

                    if (message.Subject == null)
                    {
                        if (message.Attachments.Count() > 0)
                        {
                            ListViewItem listViewItem = new ListViewItem(new string[]
                            {
                                "", message.From.OfType<MailboxAddress>().First().Name,
                                " ", message.Date.ToString("dd.MM.yyyy")
                            });
                            listViewItem.ImageIndex = 0;

                            radListView1.Items.Add(listViewItem);


                        }
                        else
                        {
                            ListViewItem listViewItem = new ListViewItem(new string[]
                            {
                                "", message.From.OfType<MailboxAddress>().First().Name,
                                " ", message.Date.ToString("dd.MM.yyyy")
                            });


                            radListView1.Items.Add(listViewItem);
                        }

                    }

                    else
                    {
                        if (message.Attachments.Count() > 0)
                        {
                            ListViewItem listViewItem = new ListViewItem(new string[]
                            {
                                "", message.From.OfType<MailboxAddress>().First().Name,
                                message.Subject, message.Date.ToString("dd.MM.yyyy")
                            });
                            listViewItem.ImageIndex = 0;

                            radListView1.Items.Add(listViewItem);

                        }
                        else
                        {
                            ListViewItem listViewItem = new ListViewItem(
                                new string[]
                                {
                                    "", message.From.OfType<MailboxAddress>().First().Name.ToString(),
                                    message.Subject.ToString(), message.Date.ToString("dd.MM.yyyy")
                                });

                            radListView1.Items.Add(listViewItem);
                        }

                    }



                    listBox2.Items.Add(i);



                }

                // radWaitingBar1.AssociatedControl = null;

            }
            catch (Exception ex)
            {
                RadMessageBox.Show(ex.Message);
            }

            //client.Disconnect(true);
        }

        private void IncomeMessage_Click(object sender, EventArgs e)
        {
            try
            {

           
            var inbox = mailclient.Inbox;
            inbox.Open(FolderAccess.ReadOnly);

            //passbox.Text = "Total messages:" + inbox.Count;
            //Console.WriteLine();
            //Console.WriteLine("Recent messages: {0}", inbox.Recent);
            radListView1.Items.Clear();

            radListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.None);
            listBox2.Items.Clear();
            radListView1.Columns[1].Text = "Отправитель";
            for (int i = inbox.Count - 1; i >= Math.Max(inbox.Count - 40, 0); i--)
            {
                var message = inbox.GetMessage(i);

                if (message.Subject == null)
                {
                    if (message.Attachments.Count() > 0)
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]{"",message.From.OfType<MailboxAddress>().First().Name,
                            " ", message.Date.ToString("dd.MM.yyyy")});
                        listViewItem.ImageIndex = 0;

                        radListView1.Items.Add(listViewItem);


                    }
                    else
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]{"",message.From.OfType<MailboxAddress>().First().Name,
                            " ", message.Date.ToString("dd.MM.yyyy")});


                        radListView1.Items.Add(listViewItem);
                    }

                }

                else
                {
                    if (message.Attachments.Count() > 0)
                    {
                        ListViewItem listViewItem = new ListViewItem(new string[]{"",message.From.OfType<MailboxAddress>().First().Name,
                            message.Subject, message.Date.ToString("dd.MM.yyyy")});
                        listViewItem.ImageIndex = 0;

                        radListView1.Items.Add(listViewItem);

                    }
                    else
                    {
                        ListViewItem listViewItem = new ListViewItem(
                        new string[]{"",message.From.OfType<MailboxAddress>().First().Name.ToString(),
                            message.Subject.ToString(), message.Date.ToString("dd.MM.yyyy")});

                        radListView1.Items.Add(listViewItem);
                    }

                }



                listBox2.Items.Add(i);



            }

            }
            catch (Exception exception)
            {
                RadMessageBox.Show(exception.Message);
            }

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {

           
            var r = new RadForm2(mailclient.Inbox, System.Convert.ToInt32(listBox2.Items[radListView1.SelectedIndices[0]]), ConnectedMail, mailclient, pass, 0);
            r.ShowDialog();
            }
            catch (Exception exception)
            {
                RadMessageBox.Show(exception.Message);
            }

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {

           
            string s = "";
            MimeMessage message;
            IMailFolder inbox;
            try
            {
                 message = mailclient.Inbox.GetMessage(System.Convert.ToInt32(listBox2.Items[radListView1.SelectedIndices[0]]));
                 inbox = mailclient.Inbox;
            }
            catch (Exception exception)
            {
                 message = mailclient.GetFolder(SpecialFolder.Sent).GetMessage(System.Convert.ToInt32(listBox2.Items[radListView1.SelectedIndices[0]]));
                 inbox = mailclient.GetFolder(SpecialFolder.Sent);
            }
            
            foreach (var d in message.To.Mailboxes)
            {
                s +=d.Address + ";";
            }
            foreach (var d in message.From.Mailboxes)
            {
                s += d.Address + ";";
            }





            s =s.Remove(s.Length-1);
            var r = new RadForm2(inbox, System.Convert.ToInt32(listBox2.Items[radListView1.SelectedIndices[0]]), ConnectedMail, mailclient, pass, 0,s);
            r.ShowDialog();
            }
            catch (Exception exception)
            {
                RadMessageBox.Show(exception.Message);
            }
        }

        private void RadPageView1_SelectedPageChanged(object sender, EventArgs e)
        {

        }
        
        private void RadLabel4_Click(object sender, EventArgs e)
        {

        }

        private void RadButton10_Click(object sender, EventArgs e)
        {
            if(passbox.PasswordChar=='*')
                passbox.PasswordChar='\0';
            else
            {
                passbox.PasswordChar = '*';
            }
        }

        private void radButton11_Click(object sender, EventArgs e)
        {
            splitPanel2.Collapsed = true;
            radSplitContainer2.Visible = false;
           // FormBorderStyle = FormBorderStyle.None;
        }

        private void radPageView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (FormBorderStyle == FormBorderStyle.None)
            {
                if (e.KeyCode == Keys.Escape)
                {
                    splitPanel2.Collapsed = false;
                    radSplitContainer2.Visible = true;
                    FormBorderStyle = FormBorderStyle.Sizable;
                }
            }
            
        }

        private void filesPanel_Resize(object sender, EventArgs e)
        {
            //var x = filesPanel.Width;
           // filesPanel.Items[0].Position = new Point(0, 2);
           // filesPanel.Items[0].Position = new Point(x-21, 2);
        }

        private void RadForm1_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            this.WindowState = FormWindowState.Maximized;
            this.Focus(); this.Show();
        }

        private void foldersPanel_Resize(object sender, EventArgs e)
        {
            var x = foldersPanel.Width;
           // foldersPanel.Items[0].Position = new Point(0, 2);
           // foldersPanel.Items[0].Position = new Point(x-21, 2);

        }

        private void RadListView1_ColumnWidthChanging(object sender, ColumnWidthChangingEventArgs e)
        {
            e.Cancel = true;
            e.NewWidth = radListView1.Columns[e.ColumnIndex].Width;
        }

        private void foldersPanel_DragOver(object sender, DragEventArgs e)
        {
            //e.Effect = DragDropEffects.Move;
        }

        private void foldersPanel_GiveFeedback(object sender, GiveFeedbackEventArgs e)
        {
           // e.UseDefaultCursors = false;
           // Cursor = new Cursor(imageList1.Images[0])
            //e.Effe = DragDropEffects.Move;
        }

        private static void ChangeToStar()
        {
            //radPageView1.SelectedPage.Text = radPageView1.SelectedPage.Text + @" * ";
        }
    }

    public class CustomRichTextEditorRibbonBar : RichTextEditorRibbonBar
    {
        protected override void Initialize()
        {
            base.Initialize();
           // buttonSaveHTML.Visible = false;
           // buttonSavePlain.Visible = false;
            //buttonSaveRich.Visible = false;
            buttonXAML.Visible = false;
            CloseButton = false;
            MaximizeButton = false;
            MinimizeButton = false;
            LayoutMode = RibbonLayout.Simplified;
            BuiltInStylesVersion = Telerik.WinForms.Documents.Model.Styles.BuiltInStylesVersion.Office2013;
            ThemeName = "Office2013Light";
        }
    }
}
