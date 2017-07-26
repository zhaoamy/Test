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

namespace test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Console.WriteLine("Hello World");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string conn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' 'C:\\Users\\Amy Zhao\\Documents\\Test\\testingExcel.xlsx';Extended Properties='Excel 12.0;HDR=Yes';";
            OleDbConnection objConn = new OleDbConnection(conn);
            objConn.Open();
            OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [Sheet1$]", objConn);
            OleDbDataAdapter objAdapter = new OleDbDataAdapter();
            objAdapter.SelectCommand = objCmdSelect;
            DataSet objDataset = new DataSet();
            objAdapter.Fill(objDataset);
            objConn.Close();
            DataTable objTable = objDataset.Tables[0];

            for (int i = 0; i < objTable.Columns.Count; i++)
            {
                for (int j = 0; j < objTable.Colum)
            }
        }
    }
}
