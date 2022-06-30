using ExcelDataReader;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace CBSProject
{
    public partial class Form1 : Form
    {
        private string connectionstring = @"Data Source=192.168.60.37;Initial Catalog=CBSDatabase;User ID=mbe;Password=Mbe@1234";

        public Form1()
        {
            InitializeComponent();
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT First_Name, Last_Name,Agent_ID FROM AGENT", conn);

                adapter.Fill(datatableDB);
            }
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                SqlDataAdapter adapter2 = new SqlDataAdapter("SELECT MBERefNo, MID_ID FROM SiteTerminal", conn);

                adapter2.Fill(datatableDBS);
            }

        }

        private DataTable dataTableMerge = new DataTable();
        private DataTable datatableDB = new DataTable();
        private DataTable datatableDBS = new DataTable();

        private DataTableCollection tableCollection;

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = dlg.FileName;
                    using (var stream = File.Open(dlg.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = ExcelDataReader => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = false
                                }
                            });
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                                cboSheet.Items.Add(table.TableName);
                        }
                    }
                }
            }
        }

        private DataTable dataTableExcel = new DataTable();

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataTableExcel = tableCollection[cboSheet.SelectedItem.ToString()];

            dataTableMerge.TableName = "Agents";
            dataTableMerge.Columns.Add("First name");

            string fName;

            int count = 0;
            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                fName = datatableDB.Rows[i][0].ToString() + " " + datatableDB.Rows[i][1].ToString();
                dataTableMerge.Rows.Add();
                dataTableMerge.Rows[count][0] = fName;
                count++;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dataTableSimplifying = new DataTable();
            dataTableSimplifying.TableName = "Agents";
            dataTableSimplifying.Columns.Add("First name");
            dataTableSimplifying.Columns.Add("Last name");
            dataTableSimplifying.Columns.Add("Agent_ID");
            dataTableSimplifying.Columns.Add("MBERefID");
            dataTableSimplifying.Columns.Add("MID_ID");
            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                dataTableSimplifying.Rows.Add();
                dataTableSimplifying.Rows[i][0] = (datatableDB.Rows[i][0].ToString());
                dataTableSimplifying.Rows[i][1] = (datatableDB.Rows[i][1].ToString());
                dataTableSimplifying.Rows[i][2] = (datatableDB.Rows[i][2].ToString());
                dataTableSimplifying.Rows[i][3] = (datatableDBS.Rows[i][0].ToString());
                dataTableSimplifying.Rows[i][4] = (datatableDBS.Rows[i][1].ToString());

            }

            dataGridView1.DataSource = dataTableSimplifying;

            for (int i = 0; i < dataTableSimplifying.Rows.Count; i++)
            {
                if (dataTableMerge.Rows[i][0].ToString() != dataTableExcel.Rows[i][0].ToString()
                   || datatableDBS.Rows[i][0].ToString() != dataTableExcel.Rows[i][1].ToString())
                {
                    dataGridView1.Rows[i].Cells[0].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[i].Cells[1].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[i].Cells[2].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[i].Cells[3].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[i].Cells[4].Style.ForeColor = Color.Red;
                }
            }
        }
    }
}