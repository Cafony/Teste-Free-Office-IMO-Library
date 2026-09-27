using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Teste_Office_IMO_Library
{
    internal class OpenDocuments
    {

        public void OpenFile(string filePath, RichTextBox richTextBox)
        {
            if (String.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path is empty.");

            if (!File.Exists(filePath))
                throw new FileNotFoundException("File not found.", filePath);

            string extension = Path.GetExtension(filePath).ToLowerInvariant();

            if (extension == ".odt")
            {
                OpenOdt(filePath, richTextBox);
            }
            else if (extension == ".docx" )
            {
                OpenDocx(filePath, richTextBox);
            }
            else if (extension == ".doc")
            {
                LoadDocToText(filePath, richTextBox);
            }
            else if (extension == ".txt" || extension == ".rtf")
            {
                OpenTextFile(filePath, richTextBox);
            }
            else if(extension == ".pdf")
            {
                OpenPdfFilesUglyToad.OpenPdfFiles(filePath, richTextBox);
            }   
            else
            {
                throw new NotSupportedException(
                    "Unsupported file type: " + extension);
            }
        }

        public void OpenTextFile(string filePath, RichTextBox richTextBox)
        {
            string extension = Path.GetExtension(filePath).ToLower();

            if (extension == ".rtf")
            {
                richTextBox.LoadFile(
                    filePath,
                    RichTextBoxStreamType.RichText
                );
            }
            else if (extension == ".txt")
            {
                richTextBox.Text = File.ReadAllText(
                    filePath,
                    Encoding.GetEncoding(1252)
                );
            }
        }



        // ============================================================
        // ODT
        // ============================================================

        public void OpenOdt(string filePath, RichTextBox richTextBox)
        {
            using (ZipArchive archive =
                ZipFile.OpenRead(filePath))
            {
                ZipArchiveEntry entry =
                    archive.GetEntry("content.xml");

                if (entry == null)
                    throw new Exception(
                        "Invalid ODT file. content.xml was not found.");

                using (Stream stream = entry.Open())
                {
                    XmlDocument xml = new XmlDocument();

                    XmlReaderSettings settings =
                        new XmlReaderSettings();

                    settings.DtdProcessing =
                        DtdProcessing.Ignore;

                    using (XmlReader reader =
                        XmlReader.Create(stream, settings))
                    {
                        xml.Load(reader);
                    }

                    StringBuilder text =
                        new StringBuilder();

                    XmlNodeList paragraphs =
                        xml.GetElementsByTagName(
                            "text:p");

                    foreach (XmlNode paragraph in paragraphs)
                    {
                        AddNodeText(paragraph, text);

                        text.AppendLine();
                    }

                    XmlNodeList headings =
                        xml.GetElementsByTagName(
                            "text:h");

                    if (paragraphs.Count == 0)
                    {
                        foreach (XmlNode heading in headings)
                        {
                            AddNodeText(heading, text);
                            text.AppendLine();
                        }
                    }

                    richTextBox.Text = text.ToString();
                }
            }
        }


        // ============================================================
        // DOCX
        // ============================================================

        public void OpenDocx(string filePath, RichTextBox richTextBox)
        {
            if (String.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path is empty.");

            if (!File.Exists(filePath))
                throw new FileNotFoundException(
                    "DOCX file not found.",
                    filePath);

            if (richTextBox == null)
                throw new ArgumentNullException(
                    "richTextBox");

            using (FileStream fileStream =
                File.OpenRead(filePath))
            {
                using (ZipArchive archive =
                    new ZipArchive(
                        fileStream,
                        ZipArchiveMode.Read))
                {
                    ZipArchiveEntry entry =
                        archive.GetEntry(
                            "word/document.xml");

                    if (entry == null)
                    {
                        throw new Exception(
                            "Invalid DOCX file. " +
                            "word/document.xml was not found.");
                    }

                    using (Stream documentStream =
                        entry.Open())
                    {
                        XmlDocument xml =
                            new XmlDocument();

                        XmlReaderSettings settings =
                            new XmlReaderSettings();

                        settings.DtdProcessing =
                            DtdProcessing.Ignore;

                        using (XmlReader reader =
                            XmlReader.Create(
                                documentStream,
                                settings))
                        {
                            xml.Load(reader);
                        }

                        XmlNamespaceManager ns =
                            new XmlNamespaceManager(
                                xml.NameTable);

                        ns.AddNamespace(
                            "w",
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                        StringBuilder text =
                            new StringBuilder();

                        XmlNodeList paragraphs =
                            xml.SelectNodes(
                                "//w:body/w:p",
                                ns);

                        if (paragraphs != null)
                        {
                            foreach (XmlNode paragraph
                                in paragraphs)
                            {
                                XmlNodeList nodes =
                                    paragraph.SelectNodes(
                                        ".//w:t | .//w:tab | .//w:br",
                                        ns);

                                if (nodes != null)
                                {
                                    foreach (XmlNode node
                                        in nodes)
                                    {
                                        if (node.LocalName == "t")
                                        {
                                            text.Append(
                                                node.InnerText);
                                        }
                                        else if (
                                            node.LocalName == "tab")
                                        {
                                            text.Append("\t");
                                        }
                                        else if (
                                            node.LocalName == "br")
                                        {
                                            text.AppendLine();
                                        }
                                    }
                                }

                                text.AppendLine();
                            }
                        }

                        richTextBox.Text =
                            text.ToString();
                    }
                }
            }

        }

        
        public void OpenDocx2(string filePath, RichTextBox richTextBox)
        {

        }

        public void LoadDocToText(string filePath, RichTextBox rtb)
        {
            try
            {
                var sb = new StringBuilder();

                using (WordDocument document = WordDocument.Load(filePath))
                {
                    foreach (var paragraph in document.Paragraphs)
                    {
                        sb.AppendLine(paragraph.Text);
                    }
                }

                rtb.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                rtb.Text = $"Failed to load document:{Environment.NewLine}{ex.Message}";
            }
        }





        // ============================================================
        // Read text from ODT XML
        // ============================================================

        private void AddNodeText(
            XmlNode node,
            StringBuilder text)
        {
            foreach (XmlNode child in node.ChildNodes)
            {
                if (child.NodeType ==
                    XmlNodeType.Text)
                {
                    text.Append(child.Value);
                }
                else if (child.Name == "text:tab")
                {
                    text.Append("\t");
                }
                else if (child.Name == "text:line-break")
                {
                    text.AppendLine();
                }
                else if (child.Name == "text:s")
                {
                    text.Append(" ");
                }
                else if (child.HasChildNodes)
                {
                    AddNodeText(child, text);
                }
            }
        }
    }




}

