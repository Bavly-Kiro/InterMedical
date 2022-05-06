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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        DataTable mainDataTable = new DataTable();

        public string ReturnValue1 { get; set; }
        public string ReturnValue2 { get; set; }
        public string ReturnValue3 { get; set; }

        public double canceled { get; set; }


        private void Form2_Load(object sender, EventArgs e)
        {

            this.ReturnValue2 = "0";

            this.ReturnValue3 = "0";

            this.canceled = 0;


            try
            {
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Debug\Database21.accdb");



                con.Open();

            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select * from groupp";
            cmd.ExecuteNonQuery();

                mainDataTable.Clear();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(mainDataTable);


            con.Close();


                materialComboBox1.DisplayMember = "nam";
                materialComboBox1.DataSource = mainDataTable;




            } catch (Exception ex)
            {
                MessageBox.Show("Error " + ex.Message);
                Console.WriteLine(ex.Message);
            }

        }

        private void materialButton2_Click(object sender, EventArgs e)
        {

            canceled = 1;
            this.Close();

        }

        private void materialButton1_Click(object sender, EventArgs e)
        {

            canceled = 0;

            if (materialCheckbox1.Checked)
            {
                this.ReturnValue2 = "1";


            }
            else {

                this.ReturnValue2 = "0";

            }

            if (materialCheckbox2.Checked)
            {
                this.ReturnValue3 = "1";

            }
            else
            {

                this.ReturnValue3 = "0";

            }

            this.ReturnValue1 = mainDataTable.Rows[materialComboBox1.SelectedIndex].Field<int>("id") + "";


            this.DialogResult = DialogResult.OK;
            this.Close();

        }



    }
}
