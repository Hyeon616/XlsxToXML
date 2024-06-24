using ClosedXML.Excel;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ConvertXlsxToXML
{
    public partial class Form1 : Form
    {
        private string excelFilePath;

        public Form1()
        {
            InitializeComponent();
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
                openFileDialog.Title = "Select an Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog.FileName;
                    lblStatus.Text = "File Uploaded: " + openFileDialog.SafeFileName;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Please upload an Excel file first.");
                return;
            }

            try
            {
                var workbook = new XLWorkbook(excelFilePath);
                var worksheet = workbook.Worksheets.First();

                var xml = new XElement("root",
                    from row in worksheet.RowsUsed().Skip(1)
                    let cells = row.CellsUsed().Select(c => new XElement(worksheet.Cell(1, c.Address.ColumnNumber).GetValue<string>(), c.GetValue<string>()))
                    select new XElement("row", cells)
                );

                string xmlFilePath = System.IO.Path.ChangeExtension(excelFilePath, ".xml");
                xml.Save(xmlFilePath);
                lblStatus.Text = "Converted to XML: " + xmlFilePath;
                MessageBox.Show("Excel file has been converted to XML.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
