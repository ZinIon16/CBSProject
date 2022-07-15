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
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                SqlDataAdapter adapter2 = new SqlDataAdapter("SELECT AgentID, SiteTerminalID,CommissionAmount,PaymentMode, PaymentYM FROM TerminalCommissionCalculation", conn);
                adapter2.Fill(datatableTCC);
            }

        }

        private DataTable dataTableMerge = new DataTable();
        private DataTable datatableDB = new DataTable();
        private DataTable datatableDBS = new DataTable();
        private DataTable datatableTCC = new DataTable();
        private DataTable datatableTCCSimple = new DataTable();

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
            datatableTCCSimple.TableName = "TCC";
            datatableTCCSimple.Columns.Add("AgentID");
            datatableTCCSimple.Columns.Add("SiteTerminalID");
            datatableTCCSimple.Columns.Add("CommissionAmount");
            datatableTCCSimple.Columns.Add("PaymentMode");
            datatableTCCSimple.Columns.Add("PaymentYM");
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
            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                datatableTCCSimple.Rows.Add();
                datatableTCCSimple.Rows[i][0] = datatableTCC.Rows[i][0].ToString();
                datatableTCCSimple.Rows[i][1] = datatableTCC.Rows[i][1].ToString();
                datatableTCCSimple.Rows[i][2] = datatableTCC.Rows[i][2].ToString();
                datatableTCCSimple.Rows[i][3] = datatableTCC.Rows[i][3].ToString();
                datatableTCCSimple.Rows[i][4] = datatableTCC.Rows[i][4].ToString();
                count++;
            }

        }
        DataTable DtFinal = new DataTable();
        DataTable Dt1 = new DataTable();
        DataTable Dt2 = new DataTable();
        private void button1_Click(object sender, EventArgs e)
        {

            DataTable dataTableSimplifying = new DataTable();
            dataTableSimplifying.TableName = "Agents";
            dataTableSimplifying.Columns.Add("First name");
            dataTableSimplifying.Columns.Add("Last name");
            dataTableSimplifying.Columns.Add("Agent_ID");
            dataTableSimplifying.Columns.Add("MBERefNo");
            dataTableSimplifying.Columns.Add("SiteTerminalID");
            dataTableSimplifying.Columns.Add("CommissionAmount");
            dataTableSimplifying.Columns.Add("PaymentMode");
            dataTableSimplifying.Columns.Add("PaymentYM");

            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                dataTableSimplifying.Rows.Add();
                dataTableSimplifying.Rows[i][0] = (datatableDB.Rows[i][0].ToString());
                dataTableSimplifying.Rows[i][1] = (datatableDB.Rows[i][1].ToString());
                dataTableSimplifying.Rows[i][2] = (datatableDB.Rows[i][2].ToString());
                dataTableSimplifying.Rows[i][3] = (datatableDBS.Rows[i][0].ToString());
                dataTableSimplifying.Rows[i][4] = (datatableDBS.Rows[i][1].ToString());
                dataTableSimplifying.Rows[i][5] = (datatableTCCSimple.Rows[i][2].ToString());
                dataTableSimplifying.Rows[i][6] = (datatableTCCSimple.Rows[i][3].ToString());
                dataTableSimplifying.Rows[i][7] = (datatableTCCSimple.Rows[i][4].ToString());

            }
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

            Dt1.Columns.Add("AgentName");
            Dt1.Columns.Add("AgentID");

            Dt2.Columns.Add("MBERefNo");
            Dt2.Columns.Add("SiteTerminalID");

            DtFinal.TableName = "Agents";
            DtFinal.Columns.Add("AgentName");
            DtFinal.Columns.Add("AgentID");
            DtFinal.Columns.Add("MBERefNo");
            DtFinal.Columns.Add("SiteTerminalID");
            DtFinal.Columns.Add("CommissionAmount");
            DtFinal.Columns.Add("PaymentMode");
            DtFinal.Columns.Add("PaymentYM");
            int counter = 0;
            Boolean flag=false;
            for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            {
                flag = false;
                for (int j = 0; j < dataTableMerge.Rows.Count; j++)
                {
                    if (dataTableExcel.Rows[i][0].ToString() == dataTableMerge.Rows[j][0].ToString())
                    {
                        flag = true;
                        Dt1.Rows.Add();
                        Dt1.Rows[counter][0] = dataTableMerge.Rows[j][0].ToString();
                        Dt1.Rows[counter][1] = dataTableMerge.Rows[j][1].ToString();
                        counter++;
                    }
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
                for (int j = 0; j < datatableDBS.Rows.Count; j++)
                {
                    if (dataTableExcel.Rows[i][1].ToString() == datatableDBS.Rows[j][0].ToString())
                    {
                        flag = true;
                        Dt2.Rows.Add();
                        Dt2.Rows[counter][0] = datatableDBS.Rows[j][0].ToString();
                        Dt2.Rows[counter][1] = datatableDBS.Rows[j][1].ToString();
                        counter++;
                    }
                }
                if (flag == false)
                {
                    Dt2.Rows.Add();
                    Dt2.Rows[counter][0] = dataTableExcel.Rows[i][1].ToString();
                    Dt2.Rows[counter][1] = "";
                    counter++;
                }
            }
            counter = 0;

            for (int i = 0; i < dataTableSimplifying.Rows.Count; i++)
            {
                for (int j = 0; j < dataTableSimplifying.Rows.Count; j++)
                {
                    using (SqlConnection conn = new SqlConnection(connectionstring))
            {

                string Agent_ID;
                string Terminal_ID;
                Agent_ID = Dt1.Rows[i][1].ToString();
                Terminal_ID = Dt2.Rows[j][1].ToString();
                        //Agent_ID = "299";
                        //Terminal_ID = "44932";
                conn.Open();
                string Query = ("SELECT * FROM TerminalCommissionCalculation WHERE " + "AgentID ='"
                    + Agent_ID + "' and SiteTerminalID='" + Terminal_ID + "' and PaymentYM ='" + date + "'");
                //SqlCommand cmd=new SqlCommand(Query,conn);
                //SqlDataReader reader = cmd.ExecuteReader();

                SqlDataAdapter adapter2 = new SqlDataAdapter("SELECT * FROM TerminalCommissionCalculation WHERE " + "AgentID ='"
                    + Agent_ID + "' and SiteTerminalID='" + Terminal_ID + "' and PaymentYM ='" + date + "'", conn);
                adapter2.Fill(DtFinal);
                    }
                }
            }

            dataGridView1.DataSource = DtFinal;

            //    for (int i = 0; i < dataTableExcel.Rows.Count; i++)
            //    {
            //        for (int j = 0; j < dataTableExcel.Rows.Count; j++)
            //        {
            //            if (dataTableExcel.Rows[i][0].ToString() != dataTableMerge.Rows[j][0].ToString())
            //            {
            //                dataGridView1.Rows[i].Cells[0].Style.ForeColor = Color.Red;
            //            }
            //        }
            //    }

            //    DateTime dateTime = dateTimePicker.Value;
            //    string date;
            //    string month;
            //    if(dateTime.Month.ToString().Length < 2)
            //    {
            //        month = "0"+ dateTime.Month.ToString();
            //        date = dateTime.Year.ToString() + month;
            //    }
            //    else
            //    {
            //        month = dateTime.Month.ToString();
            //        date = dateTime.Year.ToString() + month;
            //    }


            //    for (int i = 0; i < dataTableSimplifying.Rows.Count; i++)
            //    {

            //        if (dataTableMerge.Rows[i][0].ToString() != dataTableExcel.Rows[i][0].ToString())
            //        {
            //            dataGridView1.Rows[i].Cells[0].Style.ForeColor = Color.Red;
            //        }
            //        if (datatableDBS.Rows[i][0].ToString() != dataTableExcel.Rows[i][1].ToString())
            //        {
            //            dataGridView1.Rows[i].Cells[3].Style.ForeColor = Color.Red;
            //        }

            //        if (Convert.ToDecimal(datatableTCCSimple.Rows[i][2]) != Convert.ToDecimal(dataTableExcel.Rows[i][3]))
            //        {
            //            dataGridView1.Rows[i].Cells[5].Style.ForeColor = Color.Red;
            //        }
            //        if (datatableTCCSimple.Rows[i][3].ToString() != dataTableExcel.Rows[i][2].ToString())
            //        {
            //            dataGridView1.Rows[i].Cells[6].Style.ForeColor = Color.Red;
            //        }
            //    }

            //    //for (int i = 0; i < dataTableSimplifying.Rows.Count; i++)
            //    //{
            //    //    for (int j = 0; j < dataTableSimplifying.Rows.Count; j++)
            //    //    {
            //            using (SqlConnection conn = new SqlConnection(connectionstring))
            //            {

            //                string Agent_ID;
            //                //Agent_ID = datatableDB.Rows[i][2].ToString();
            //                string Terminal_ID;
            //                //Terminal_ID = datatableDBS.Rows[j][1].ToString();
            //                Agent_ID = "299";
            //                Terminal_ID = "44932";
            //                conn.Open();
            //                string Query = ("SELECT * FROM TerminalCommissionCalculation WHERE " + "AgentID ='"
            //                    + Agent_ID + "' and SiteTerminalID='" + Terminal_ID + "' and PaymentYM ='" + date + "'");
            //                //SqlCommand cmd=new SqlCommand(Query,conn);
            //                //SqlDataReader reader = cmd.ExecuteReader();

            //                SqlDataAdapter adapter2 = new SqlDataAdapter("SELECT * FROM TerminalCommissionCalculation WHERE " + "AgentID ='"
            //                    + Agent_ID + "' and SiteTerminalID='" + Terminal_ID + "' and PaymentYM ='" + date + "'", conn);
            //                adapter2.Fill(DtFinal);
            //            }
            //    //    }
            //    //}
            //    dataGridView1.DataSource = DtFinal;
        }

    }
}