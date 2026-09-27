using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Teste_Office_IMO_Library
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //load doc files
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "OpenDocument Text (*.odt)|*.odt|" +
                                    "Word .docx (*.docx)|*.docx|" +
                                    "Word .doc (*.doc)|*.doc|" +
                                    "Pdf Files (*.pdf)|*.pdf|" +
                                    "All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                OpenDocuments openDocuments = new OpenDocuments();
                try
                {
                    openDocuments.OpenFile(filePath, richTextBox1);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening file: " + ex.Message);
                }
            }


        }
    }
}
