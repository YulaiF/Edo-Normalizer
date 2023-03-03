using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using Application = System.Windows.Forms.Application;
using File = System.IO.File;

namespace Edo_Normalizer
{
    public partial class Form1 : Form
    {
        string lastProcceesFilePath;
        readonly string separator = "=======================================================" + Environment.NewLine + Environment.NewLine;

        

        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);

            richTextBox1.AllowDrop = true;
            richTextBox1.DragEnter += new DragEventHandler(richTextBox1_DragEnter);
            richTextBox1.DragDrop += new DragEventHandler(richTextBox1_DragDrop);
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                ProccesFile(file);
            }
        }

        void richTextBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void richTextBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData("FileDrop");
            foreach (string file in files)
            {
                ProccesFile(file);
            }
        }

       

        private void Form1_Load(object sender, EventArgs e)
        {
            Text = Application.ProductName + " " + Application.ProductVersion;
        }


        /// <summary>
        /// Обработка входящего XML/ZIP
        /// </summary>
        /// <param name="FILENAME"></param>
        private void ProccesFile(string FILENAME)
        {
            var splExt = FILENAME.Split('.');
            var extension = splExt[splExt.Count() - 1];

            switch (extension)
            {
                case "zip":
                    var tmpFile = Path.GetTempFileName();
                    var isXMLFileFind = false;
                    richTextBox1.AppendText("Обработка файла... "
                                        + Environment.NewLine
                                        + FILENAME
                                        + Environment.NewLine
                                        + Environment.NewLine);
                    using (ZipArchive archive = ZipFile.OpenRead(FILENAME))
                    {
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            if (entry.Name.Contains("ON_NSCHFDOPPR"))
                            {
                                isXMLFileFind = true;
                                entry.ExtractToFile(tmpFile, true);
                                MakeZaebic(tmpFile, Path.Combine(Path.GetDirectoryName(FILENAME), entry.Name));
                            }
                        }
                    }
                    if (!isXMLFileFind)
                    {
                        richTextBox1.AppendText("Файл \'ON_NSCHFDOPPR*.xml\' в архиве не обнаружен"
                                        + Environment.NewLine
                                        + separator, Color.Red);
                    }
                    File.Delete(tmpFile);

                    break;
                case "xml":
                    if (FILENAME.Contains("ON_NSCHFDOPPR"))
                    {
                        richTextBox1.AppendText("Обработка файла... "
                                        + Environment.NewLine
                                        + FILENAME
                                        + Environment.NewLine
                                        + Environment.NewLine);
                        MakeZaebic(FILENAME, FILENAME);
                    }
                    else
                    {
                        //richrichTextBox1.fore
                        richTextBox1.AppendText("Не поддерживаемый тип xml, файл не содержит \'ON_NSCHFDOPPR\' в имени"
                        + Environment.NewLine
                        + Environment.NewLine
                        + separator,Color.Red);
                    }

                    break;
                default:
                    if (FILENAME != "")
                    {
                        richTextBox1.AppendText("Неподдерживаемый тип файла: "
                                            + extension
                                            + Environment.NewLine
                                            + separator, Color.Red);
                    }
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var FILENAME = "";
            var fileDialog = new OpenFileDialog();
            fileDialog.Filter = "ZIP,XML|*ON_NSCHFDOPPR*.zip;*ON_NSCHFDOPPR*.xml|Все ZIP|*.zip|Все XML|*.xml";
            fileDialog.AddExtension = true;
            fileDialog.ShowDialog();
            if (fileDialog.FileName != "")
            {
                FILENAME = fileDialog.FileName;
                ProccesFile(FILENAME);
            }
        }

        private bool MakeZaebic(string pathToXMLFile, string pathToSaveXMLFile)
        {
            try
            {
                var doc = new XmlDocument();

                doc.Load(pathToXMLFile);
                foreach (XmlElement node in doc.GetElementsByTagName("СвУчДокОбор"))
                {
                    for (int i = 0; i < node.Attributes.Count; i++)
                    {
                        var newValue = node.Attributes.Item(i).Value.Replace("_", "-");
                        node.Attributes.Item(i).Value = newValue;
                    }

                }
                richTextBox1.AppendText("Исправленный файл сохранён в "
                    + pathToSaveXMLFile
                    + Environment.NewLine
                    + Environment.NewLine
                    + separator, Color.Green);
                doc.Save(pathToSaveXMLFile);
                lastProcceesFilePath = pathToSaveXMLFile;
                return true;
            }
            catch (Exception ex)
            {
                richTextBox1.AppendText(ex.Message + Environment.NewLine + Environment.NewLine + separator, Color.Red);
                return false;
                //throw;
            }
        }

        /// <summary>
        /// Показать файл в Проводнике
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool ExploreFile(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                return false;
            }
            //Clean up file path so it can be navigated OK
            filePath = System.IO.Path.GetFullPath(filePath);
            System.Diagnostics.Process.Start("explorer.exe", string.Format("/select,\"{0}\"", filePath));
            return true;
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var AboutBox = new AboutBox1();
            AboutBox.ShowDialog();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void выбратьФайлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1_Click(sender, e);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (lastProcceesFilePath != "")
            {
                if (File.Exists(lastProcceesFilePath))
                {
                    ExploreFile(lastProcceesFilePath);
                }
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            // set the current caret position to the end
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            // scroll it automatically
            richTextBox1.ScrollToCaret();
        }
    }

    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;

            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }
    }
}
