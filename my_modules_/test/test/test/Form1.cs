using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\KVPL\KVPLS.mdb";
            string select = "SELECT SPZeu.* FROM SPZEU";
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, con);
            archiveDataSet.spZEUDataTable zeuTable = new archiveDataSet.spZEUDataTable();
            adapter.Fill(zeuTable);
            archiveDataSet ds = new archiveDataSet();

            dataGridView1.DataSource = zeuTable;
            System.Data.DataRow[] rows = zeuTable.Select();

            /*
            DataTable dr_art_line_2 = ds.Tables["QuantityInIssueUnit"];

            foreach (DataRow row in dr_art_line_2.Rows)
            {
                QuantityInIssueUnit_value = Convert.ToInt32(row["columnname"]);
            } */
            //dataGridView1.DataMember = "spZEU";
            
        }
    }
}
