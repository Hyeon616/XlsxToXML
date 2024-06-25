using ClosedXML.Excel;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using MySql.Data.MySqlClient;
using System.IO;
using System.Collections.Generic;

namespace ConvertXlsxToXML
{
    public partial class Form1 : Form
    {
        private string excelFilePath;
        private string connectionString = "server=127.0.0.1;user=root;password=1234;database=wjhdb;port=3306";

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

                string xmlFilePath = Path.ChangeExtension(excelFilePath, ".xml");
                xml.Save(xmlFilePath);
                lblStatus.Text = "Converted to XML: " + xmlFilePath;

                // 파일 이름에서 테이블 이름 생성
                string tableName = Path.GetFileNameWithoutExtension(excelFilePath);

                // XML 데이터를 MariaDB에 추가하는 기능 호출
                LoadXmlToMariaDb(xml, tableName);

                MessageBox.Show("Excel file has been converted to XML and loaded to the database.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void LoadXmlToMariaDb(XElement xml, string tableName)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();

                    // 기존 테이블 삭제
                    DropTableIfExists(conn, tableName);

                    // 테이블 생성
                    CreateTableFromXml(conn, tableName, xml);

                    // 데이터 삽입
                    InsertDataFromXml(conn, tableName, xml);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while loading XML to the database: " + ex.Message);
            }
        }

        private void DropTableIfExists(MySqlConnection conn, string tableName)
        {
            string dropTableQuery = $"DROP TABLE IF EXISTS `{tableName}`;";
            using (MySqlCommand cmd = new MySqlCommand(dropTableQuery, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }

        private void CreateTableFromXml(MySqlConnection conn, string tableName, XElement xml)
        {
            string createTableQuery = $"CREATE TABLE `{tableName}` (";

            // 컬럼 정의 추출
            var firstRow = xml.Elements("row").First();
            var columns = new HashSet<string>(); // 중복된 컬럼 이름 방지
            foreach (var element in firstRow.Elements())
            {
                string columnName = element.Name.LocalName;
                if (!columns.Add(columnName))
                {
                    MessageBox.Show($"Duplicate column name found: {columnName}");
                    return;
                }
                string columnType = "VARCHAR(255)"; // 기본적으로 VARCHAR로 설정, 필요에 따라 변경 가능
                createTableQuery += $"`{columnName}` {columnType}, ";
            }

            // 마지막 쉼표 제거 및 쿼리 종료
            createTableQuery = createTableQuery.TrimEnd(',', ' ') + ");";

            using (MySqlCommand cmd = new MySqlCommand(createTableQuery, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }

        private void InsertDataFromXml(MySqlConnection conn, string tableName, XElement xml)
        {
            foreach (var row in xml.Elements("row"))
            {
                string insertQuery = $"INSERT INTO `{tableName}` (";

                foreach (var col in row.Elements())
                {
                    insertQuery += $"`{col.Name.LocalName}`, ";
                }

                // 마지막 쉼표 제거 및 쿼리 종료
                insertQuery = insertQuery.TrimEnd(',', ' ') + ") VALUES (";

                foreach (var col in row.Elements())
                {
                    insertQuery += $"'{col.Value}', ";
                }

                // 마지막 쉼표 제거 및 쿼리 종료
                insertQuery = insertQuery.TrimEnd(',', ' ') + ");";

                using (MySqlCommand cmd = new MySqlCommand(insertQuery, conn))
                {
                    cmd.ExecuteNonQuery();
                }
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
