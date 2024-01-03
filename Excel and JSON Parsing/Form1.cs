//using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Excel;

namespace Excel_Parsing_and_DataGridView
{
    public partial class Form1 : Form
    {
        DataSet data;
        public Form1()
        {
            InitializeComponent();
        }

        private void openFileBtn_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog1.FileName;
                Load_Excel(fileName);
            }
        }

        private void saveJSONbtn_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                Save_JSON();
        }

        public void Load_Excel(string path)
        {
            if (path.EndsWith(".xlsx"))
            {
                try
                {
                    using (var fileStream = File.Open(path, FileMode.Open, FileAccess.Read))
                    {
                        IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
                        reader.IsFirstRowAsColumnNames = true;
                        data = reader.AsDataSet();
                        dataGridView1.DataSource = data.Tables[0];
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                }
            }

            else if (path.EndsWith(".json"))
            {
                var jsonFile = File.ReadAllText(path);
                dataGridView1.DataSource = JsonConvert.DeserializeObject<DataTable>(jsonFile);
            }
           
        }

        public void Save_JSON()
        {
            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                Dictionary<string, object> rowData = new Dictionary<string, object>();

                foreach (DataGridViewColumn column in dataGridView1.Columns)
                {
                    rowData[column.Name] = row.Cells[column.Index].Value;
                }

                data.Add(rowData);
            }

            var jsonFile = JsonConvert.SerializeObject(data, Formatting.Indented);
            String fileName = saveFileDialog1.FileName;
            File.WriteAllText(fileName, jsonFile);
        } 
    }
}




/*
 * 
 * Using Interop :
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook excelWB = excelApp.Workbooks.Open(path);
                    Excel.Worksheet excelWS = excelWB.Sheets[1];

                    DataTable data = new DataTable();

                    Excel.Range excelRange = excelWS.UsedRange;
                    for (int row = 1; row <= excelRange.Rows.Count; row++)
                    {
                        DataRow dataRow = data.NewRow();
                        for (int col = 1; col <= excelRange.Columns.Count; col++)
                        {
                            if (row == 1)       
                            {
                                string headerData = (excelWS.Cells[row, col] as Excel.Range).Value;
                                data.Columns.Add(headerData);
                            }
                            else
                            {
                                dataRow[col - 1] = Convert.ToString((excelWS.Cells[row, col] as Excel.Range).Value);
                            }
                        }
                        if (row > 1)
                        {
                            data.Rows.Add(dataRow);
                        }

                        foreach (DataRow row_ in data.Rows)
                        {
                            Employee employee = new Employee
                            {
                                Id = Convert.ToInt32(row_["Id"]),
                                Name = row_["Name"].ToString(),
                                Age = Convert.ToInt32(row_["Age"]),
                                Designation = row_["Designation"].ToString(),
                            };
                            employees.Add(employee);
                        }
                    }
                    dataGridView1.DataSource = data;
                    */
