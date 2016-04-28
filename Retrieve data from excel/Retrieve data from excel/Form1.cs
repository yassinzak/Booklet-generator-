using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace Retrieve_data_from_excel
{
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }


        public void getsheets()
        {
            comboBox2.Items.Clear();
            string stringconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textselect.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn = new OleDbConnection(stringconn);

            try {
                conn.Open();
                DataTable Sheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);


                //for (int i = 0; i < Sheets.Rows.Count; i++)
                //{
                string worksheets = Sheets.Rows[0]["TABLE_NAME"].ToString();
                string sqlQuery = String.Format("SELECT * FROM [{0}]", worksheets);
                OleDbDataAdapter da = new OleDbDataAdapter(sqlQuery, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                foreach (DataRow dr in Sheets.Rows)
                {
                    string sht = dr[2].ToString().Replace("'", "");
                    sht = sht.Substring(0, sht.Length - 1);
                    comboBox2.Items.Add(sht);
                }
            } catch(Exception ex)
            {
                MessageBox.Show("please exit the excel file to be opened by the program ");
            }
               
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfd = new OpenFileDialog();
            if (opfd.ShowDialog() == DialogResult.OK)
                textselect.Text = opfd.FileName;
            getsheets();
        }


        public List<string> GetTableColumnNames(string tableName)
        {
            string stringconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textselect.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn = new OleDbConnection(stringconn);
                conn.Open();
                var schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new Object[] { null, null, tableName });
                if (schemaTable == null)
                    return null;
                var columnOrdinalForName = schemaTable.Columns["COLUMN_NAME"].Ordinal;
                return (from DataRow r in schemaTable.Rows select r.ItemArray[columnOrdinalForName].ToString()).ToList();
        }

        //public void printmessage(){
        //    var conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textselect.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
        //    using (var con = new OleDbConnection(conStr)){
        //        con.Open();
        //        using (var cmd = new OleDbCommand("Select * from [" + textchoice.Text + "$]",con))
        //        using (var reader = cmd.ExecuteReader(CommandBehavior.SchemaOnly)){
        //            var table = reader.GetSchemaTable();
        //            var nameCol= table.Columns["ColumnName"];
        //            foreach (DataRow row in table.Rows)
        //            {
        //                MessageBox.Show(row[nameCol].ToString()) ;
        //            }
        //        }
        //    }
        //}
     

        private void button2_Click(object sender, EventArgs e)
        {
            if (textselect.Text != "")
            {
               try
                {
                    string stringconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textselect.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
                    OleDbConnection conn = new OleDbConnection(stringconn);
                    OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + comboBox2.Text + "$]", conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    comboBox1.Items.Clear();
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        string columnName = this.dataGridView1.Columns[i].Name;
                        comboBox1.Items.Add(columnName);
                    }
                   // getsheets();
               }
               catch (OleDbException ex)
                { MessageBox.Show("ER"); }
           //printmessage();
            }
            else
                MessageBox.Show("ER");
   
        }

        private void button3_Click(object sender, EventArgs e)
        {

            bool time = false;
            DateTime date = dateTimePicker1.Value;
            double duration = 0; ;
            if (textDuration.Text != "") 
                try
                {
                    duration = Convert.ToDouble(textDuration.Text);
                    time = true;
                }
                catch (Exception ex) { MessageBox.Show("the value entered in the field is wrong please try again"); return; }
              StreamWriter w = File.AppendText("C:\\new\\Text.doc");
              w.Write("\t Time\t" + "|");
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                string columnName = this.dataGridView1.Columns[i].Name;
                
                    w.Write("\t" + columnName + "\t" + "|");
                
            }
           
                w.WriteLine("\n");
                w.WriteLine("-------------------------------------------------------------------");
            
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)         
            {
                if (time)
                {
                    w.Write("\t" + date.ToString("HH:mm") + " - " + date.AddMinutes(duration).ToString("HH:mm") + "\t" + "|");
                    date = date.AddMinutes(duration);
                }
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    
                        w.Write("\t" + dataGridView1.Rows[i].Cells[j].Value.ToString() + "\t" + "|");
                    
                } 
      
              
                            w.WriteLine("");
                            w.WriteLine("-------------------------------------------------------------------");

                    
               
            }             
            w.Close();
            MessageBox.Show("Data Exported"); 
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "" )
                MessageBox.Show("ER");
            else
            {
                try
                {

                    string stringconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textselect.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
                    OleDbConnection conn = new OleDbConnection(stringconn);
                    OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + comboBox2.Text + "$] where " + comboBox1.Text + " = " + textBox1.Text, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    comboBox1.Items.Clear();
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        string columnName = this.dataGridView1.Columns[i].Name;
                        comboBox1.Items.Add(columnName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Result Null");
                }
                //printmessage();
            }
            
        }
    }

 }

