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
                SqlDataAdapter adapter2 = new SqlDataAdapter("SELECT MBERefNo, SiteTerminalID FROM SiteTerminal", conn);
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
            dataTableMerge.Columns.Add("AgentName");
            dataTableMerge.Columns.Add("AgentID");
      
            string fName;

            int count = 0;
            for (int i = 0; i < datatableDB.Rows.Count; i++)
            {
                fName = datatableDB.Rows[i][0].ToString() + " " + datatableDB.Rows[i][1].ToString();
                dataTableMerge.Rows.Add();
                dataTableMerge.Rows[count][0] = fName;
                dataTableMerge.Rows[count][1] = datatableDB.Rows[i][2].ToString();
                count++;
            }
            count = 0;
       

        }
        DataTable DtFinal = new DataTable();
        DataTable Dt1 = new DataTable();
        DataTable Dt2 = new DataTable();
        bool AddedFlag = false;
        bool UpdatedFlag = false;
        private void button1_Click(object sender, EventArgs e)
        {
            

            Dt1.Columns.Add("AgentName");
            Dt1.Columns.Add("AgentID");

            Dt2.Columns.Add("MBERefNo");
            Dt2.Columns.Add("SiteTerminalID");

          
            int counter = 0;
            Boolean flag=false;
            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                flag = false;
          
                    if (dataTableExcel.Rows[i][0].ToString() == dataTableMerge.Rows[i][0].ToString())
                    {
                        flag = true;
                        Dt1.Rows.Add();
                        Dt1.Rows[counter][0] = dataTableMerge.Rows[i][0].ToString();
                        Dt1.Rows[counter][1] = dataTableMerge.Rows[i][1].ToString();
                        counter++;
                    }
           
                if (flag == false)
                {
                    Dt1.Rows.Add();
                    Dt1.Rows[counter][0] = dataTableExcel.Rows[i][0].ToString();
                    Dt1.Rows[counter][1] = "";
                    counter++;
                }
            }
            counter= 0;
            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                flag = false;
            
                    if (dataTableExcel.Rows[i][1].ToString() == datatableDBS.Rows[i][0].ToString())
                    {
                        flag = true;
                        Dt2.Rows.Add();
                        Dt2.Rows[counter][0] = datatableDBS.Rows[i][0].ToString();
                        Dt2.Rows[counter][1] = datatableDBS.Rows[i][1].ToString();
                        counter++;
                    }

                if (flag == false)
                {
                    Dt2.Rows.Add();
                    Dt2.Rows[counter][0] = dataTableExcel.Rows[i][1].ToString();
                    Dt2.Rows[counter][1] = "";
                    counter++;
                }
            }
      
    
            DtFinal.TableName = "Agents";
            DtFinal.Columns.Add("AgentName");
            DtFinal.Columns.Add("AgentID");
            DtFinal.Columns.Add("MBERefNo");
            DtFinal.Columns.Add("SiteTerminalID");
            DtFinal.Columns.Add("CommissionAmount");
            DtFinal.Columns.Add("PaymentMode");

            counter = 0;
            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                DtFinal.Rows.Add(Dt1.Rows[i][0]);
                DtFinal.Rows[i][1] = Dt1.Rows[i][1];
                DtFinal.Rows[i][2] = Dt2.Rows[i][0];
                DtFinal.Rows[i][3] = Dt2.Rows[i][1];
                DtFinal.Rows[i][4] = dataTableExcel.Rows[i][2];
                DtFinal.Rows[i][5] = dataTableExcel.Rows[i][3];
            }
            dataGridView1.DataSource = DtFinal;



            for (int i = 0; i < DtFinal.Rows.Count; i++)
            {
                for (int j = 0; j < DtFinal.Columns.Count; j++)
                {
                    if (DtFinal.Rows[i][j].ToString() == "")
                    {
                        dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Red;
                    }
                }
            }

        }

        private void btnDB_Click(object sender, EventArgs e)
        {
            DateTime dateTime = dateTimePicker.Value;
            string date;
            string month;
            if (dateTime.Month.ToString().Length < 2)
            {
                month = "0" + dateTime.Month.ToString();
                date = dateTime.Year.ToString() + month;
            }
            else
            {
                month = dateTime.Month.ToString();
                date = dateTime.Year.ToString() + month;
            }
            int counter = 0;
            int counterdt1 = 0;
            int counterdt2 = 0;

            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                counterdt1++;
                counterdt2++;
                SqlConnection conn = new SqlConnection(connectionstring);
                using (var cmd = conn.CreateCommand())
                {
                    string result;
                    string ID;
                    string Agent_ID;
                    string Terminal_ID;
                    decimal ComAmount;
                    string PaymentMode;


                    if (Dt1.Rows[counter][1].ToString() == "")
                    {
                        counterdt1++;
                    }
                    if (Dt2.Rows[counter][1].ToString() == "")
                    {
                        counterdt2++;
                    }
                    Agent_ID = Dt1.Rows[counter][1].ToString();
                    Terminal_ID = Dt2.Rows[counter][1].ToString();
                    ComAmount = Convert.ToDecimal(dataTableExcel.Rows[i][3]);
                    PaymentMode = dataTableExcel.Rows[i][2].ToString();
                    conn.Open();
                    cmd.CommandText = "SpCheckAndUpdateRecord '" + Agent_ID + "','" + Terminal_ID + "'," + ComAmount + ",'" + PaymentMode + "','" + date + "'";


                    using (var reader2 = cmd.ExecuteReader())
                    {
                        if (reader2.HasRows)
                        {
                            while (reader2.Read())
                            {

                                result = (reader2["Result"].ToString());
                                if (result == "1")
                                {
                                    UpdatedFlag = true;
                                }

                                else
                                {
                                    AddedFlag = true;
                                }

                            }
                        }
                        conn.Close();
                    }

                }
            }
            if (UpdatedFlag == true)
            {
                MessageBox.Show("Record updated successfully");
            }
            else if (AddedFlag == true)
            {
                MessageBox.Show("Record added successfully");
            }
        }
    }
}