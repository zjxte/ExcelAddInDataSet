using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using System.Data.OleDb;

namespace ExcelAddInDataSet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string myDB = "ProductNotes-Search.xlsx";
                System.Data.OleDb.OleDbConnection myConnection;
                DataSet dtset;
                System.Data.OleDb.OleDbDataAdapter myCommand;
                myConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + myDB + ";Extended Properties=dBASE IV;");
                myCommand = new System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [Sheet1$}", myConnection);
                myCommand.TableMappings.Add("Table", "TestTable");
                dtset = new System.Data.DataSet();
                myCommand.Fill(dtset);
                dataGridView1.DataSource = dtset.Tables[0];
                myConnection.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
