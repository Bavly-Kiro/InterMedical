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

namespace InterMedical
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        DataTable mainDataTable = new DataTable();




        private void Form3_Load(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();


            try
            {
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Debug\Database21.accdb");


                con.Open();

                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select * from products";
                cmd.ExecuteNonQuery();


                mainDataTable.Clear();

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(mainDataTable);


                con.Close();

                double idd = 0;
                string nam = "";

                for (int i = 0; i < mainDataTable.Rows.Count; i++)
                {



                    idd = mainDataTable.Rows[i].Field<int>("ID");
                    nam = mainDataTable.Rows[i].Field<string>("nam");


                    dataGridView1.Rows.Insert(i, 1);

                    dataGridView1.Rows[i].Cells[0].Value = idd;
                    dataGridView1.Rows[i].Cells[1].Value = nam;





                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex.Message);
                Console.WriteLine(ex.Message);
            }




        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
