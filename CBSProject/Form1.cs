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

        private DataTable dt2 = new DataTable();
        private string lastName;

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            dt2 = tableCollection[cboSheet.SelectedItem.ToString()];

            dataTable.TableName = "Agents";
            dataTable.Columns.Add("First name");
            //dataTable.Columns.Add("Last Name");

            //DataRow row12 = dataSet.Tables["Banks"].NewRow();

            string fName;
            string firstName;

            int count = 0;
            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                fName = dt.Rows[i][0].ToString() + " " + dt.Rows[i][1].ToString();
                //string[] fNameArr = fName.Split(' ');
                //firstName = fNameArr[0];
                //if(fNameArr.Length == 3)
                //{
                //    lastName = fNameArr[1] + fNameArr[2];
                //}
                //else
                //{
                //    lastName = fNameArr[1];
                //}

                dataTable.Rows.Add();
                dataTable.Rows[count][0] = fName;
                //dataTable.Rows[count][1] = lastName;
                count++;
            }
        }

        private SqlDataAdapter adapter2;
        private DataTable DataTable2 = new DataTable();

        private void button1_Click(object sender, EventArgs e)
        {
           

            DataTable dt3 = new DataTable();
            dt3.TableName = "Agents";
            dt3.Columns.Add("First name");
            dt3.Columns.Add("Last name");
            dt3.Columns.Add("Agent_ID");
            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                
                    dt3.Rows.Add();
                    dt3.Rows[i][0] = (dt.Rows[i][0].ToString());
                dt3.Rows[i][1] = (dt.Rows[i][1].ToString());
                dt3.Rows[i][2] = (dt.Rows[i][2].ToString());


            }
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    string qry = "SELECT Agent_ID, First_Name, Last_Name FROM AGENT WHERE First_Name = '" + dt.Rows[i][0].ToString() + "'" + "and Last_Name = '" + dt.Rows[i][1].ToString() + "'";

                    adapter2 = new SqlDataAdapter("SELECT Agent_ID,First_Name,Last_Name FROM AGENT WHERE First_Name = '" 
                        + dt.Rows[i][0].ToString() + "'" + "and Last_Name = '" + dt.Rows[i][1].ToString() + "'", conn);
                    adapter2.Fill(DataTable2);
                }
            }
            dataGridView1.DataSource = dt3;

            for (int i = 0; i < dt3.Rows.Count; i++)
            {
                if (dataTable.Rows[i][0].ToString() != dt2.Rows[i][0].ToString() )
                {
                    dataGridView1.Rows[i ].Cells[0].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[i ].Cells[1].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[i ].Cells[2].Style.ForeColor = Color.Red;
                    //dataGridView1.Rows[i].Cells[3].Style.ForeColor = Color.Red;
                }
            }

        
        }
    }
}