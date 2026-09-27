using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Teste_Office_IMO_Library
{
    internal class OpenPdfFilesUglyToad
    {
        public static void OpenPdfFiles1(string filePath, RichTextBox richTextBox)
        {
            StringBuilder text = new StringBuilder();

            using (PdfDocument document = PdfDocument.Open(filePath))
            {
                foreach (var page in document.GetPages())
                {
                    // Get words from the page
                    var words = page.GetWords()
                        .OrderByDescending(w => w.BoundingBox.Top)
                        .ThenBy(w => w.BoundingBox.Left)
                        .ToList();

                    if (words.Count == 0)
                        continue;

                    double currentY = words[0].BoundingBox.Top;

                    foreach (var word in words)
                    {
                        // Difference in vertical position
                        double difference = Math.Abs(currentY - word.BoundingBox.Top);

                        // New visual line
                        if (difference > 3)
                        {
                            text.AppendLine();
                            currentY = word.BoundingBox.Top;
                        }
                        else
                        {
                            // Space between words
                            if (text.Length > 0 &&
                                text[text.Length - 1] != '\n' &&
                                text[text.Length - 1] != ' ')
                            {
                                text.Append(" ");
                            }
                        }

                        text.Append(word.Text);
                    }

                    // Separate PDF pages
                    text.AppendLine();
                    text.AppendLine();
                }
            }

            richTextBox.Text = text.ToString();
        }

        public static void OpenPdfFiles(string filePath, RichTextBox richTextBox)
        {
            try
            {
                var document = PdfDocument.Open(filePath);
                var sb = new StringBuilder();

                foreach (var page in document.GetPages())
                {
                    string pageText = ContentOrderTextExtractor.GetText(page);
                    sb.AppendLine(pageText);
                    sb.AppendLine();
                }

                richTextBox.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open PDF: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



    }
}
