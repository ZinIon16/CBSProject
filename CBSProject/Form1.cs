using ExcelDataReader;
using System;
using System.Data;
using System.Data.SqlClient;
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
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT First_Name, Last_Name FROM AGENT", conn);

                adapter.Fill(dt);
                //dataGridView1.DataSource = dt;
            }
        }

        private DataTable dataTable = new DataTable();
        private DataTable dt = new DataTable();

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

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt2 = new DataTable();
            dt2 = tableCollection[cboSheet.SelectedItem.ToString()];

            dataTable.TableName = "Agents";
            dataTable.Columns.Add("First name");
            dataTable.Columns.Add("Last Name");

            //DataRow row12 = dataSet.Tables["Banks"].NewRow();

            string fName;
            string firstName;
            string lastName;
            int count = 0;
            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                fName = dt2.Rows[i][0].ToString();
                string[] fNameArr = fName.Split(' ');
                firstName = fNameArr[0];
                lastName = fNameArr[1];
                if (firstName == dt.Rows[i][0].ToString() && lastName == dt.Rows[i][1].ToString())
                {
                    dataTable.Rows.Add();
                    dataTable.Rows[count][0] = firstName;
                    dataTable.Rows[count][1] = lastName;
                    count++;
                }

            }
        }

        private SqlDataAdapter adapter2;
        private DataTable DataTable2 = new DataTable();

        private void button1_Click(object sender, EventArgs e)
        {
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    adapter2 = new SqlDataAdapter("SELECT Agent_ID,First_Name,Last_Name FROM AGENT WHERE First_Name = '" + dataTable.Rows[i][0].ToString() + "'", conn);
                    adapter2.Fill(DataTable2);
                }

                dataGridView1.DataSource = DataTable2;
            }
        }
    }
}