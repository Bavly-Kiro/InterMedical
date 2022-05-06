using MaterialSkin.Controls;
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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace InterMedical
{
    public partial class Form1 : MaterialForm
    {

        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Debug\Database21.accdb");

        System.Data.DataTable mainDataTable;
        System.Data.DataTable sellProductmainDataTable;
        System.Data.DataTable materialButton4mainDataTable;
        System.Data.DataTable materialButton5mainDataTable;
        System.Data.DataTable loadTab6GridViewmainDataTable;
        System.Data.DataTable materialButton6mainDataTable;
        System.Data.DataTable checkProductsAvailapilitymainDataTable;
        System.Data.DataTable dataTableOfComboBox;


        double totalPrice;
        string commandText = "";
        string bCode = "";
        public string codeNum = "";
        string sellCommandText = "";


        //for sell
        string spid = "";
        string snum = "";
        string sprice = "";

        DateTime dateTime;
        string sDate = "";

        string sGroupName = "0";
        string sCash = "0";
        string isReset = "0";
        double available = 0;





        readonly MaterialSkin.MaterialSkinManager materialSkinManager;
        public Form1()
        {
            InitializeComponent();
            materialSkinManager = MaterialSkin.MaterialSkinManager.Instance;
            materialSkinManager.EnforceBackcolorOnAllComponents = true;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkin.MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new MaterialSkin.ColorScheme(
                MaterialSkin.Primary.Indigo500,
                MaterialSkin.Primary.Indigo700,
                MaterialSkin.Primary.Indigo100,
                MaterialSkin.Accent.Pink200,
                MaterialSkin.TextShade.WHITE);
        }


        //once form is loaded
        private void Form1_Load(object sender, EventArgs e)
        {

            mainDataTable = new System.Data.DataTable();
            checkProductsAvailapilitymainDataTable = new System.Data.DataTable();
            sellProductmainDataTable = new System.Data.DataTable();
            materialButton4mainDataTable = new System.Data.DataTable();
            materialButton5mainDataTable = new System.Data.DataTable();
            loadTab6GridViewmainDataTable = new System.Data.DataTable();
            materialButton6mainDataTable = new System.Data.DataTable();
            dataTableOfComboBox = new System.Data.DataTable();


            totalPrice = 0;

            bCodeList.Clear();
            //mainDataGridView.AllowUserToAddRows = true;


            // allDataPricesGridView.Columns[4].DefaultCellStyle.Format = "MM/dd/yyyy";


        }

        List<string> bCodeList = new List<string>();



        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {



                if (bCodeTextBox.Text != "")
                {


                    bCode = bCodeTextBox.Text;

                    commandText = "select * from products where bcode ='" + bCode + "'";
                    addProductToList();


                }
                else
                {

                    MessageBox.Show("قم بادخال الباركود او قرأته بالجهاز");

                }

            }


        }

        void addProductToList()
        {


            try
            {


                con.Open();

                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = commandText;
                cmd.ExecuteNonQuery();

                mainDataTable.Clear();

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(mainDataTable);

                string nam = "";
                string id = "";
                string price = "";


                if (mainDataTable.Rows.Count > 0)
                {

                    if (bCodeList.Count != 0)
                    {
                        for (int i = 0; i < bCodeList.Count; i++)
                        {

                            if (bCodeList[i] == mainDataTable.Rows[0].Field<string>("bcode"))
                            {
                                MessageBox.Show("المنتج موجود بالفعل, لاضافة المزيد قم بتغير العدد");
                                con.Close();
                                return;

                            }

                        }

                        bCodeList.Add(mainDataTable.Rows[0].Field<string>("bcode"));


                    }
                    else
                    {

                        bCodeList.Add(mainDataTable.Rows[0].Field<string>("bcode"));

                    }


                    nam = mainDataTable.Rows[0].Field<string>("nam");
                    id = Convert.ToString(mainDataTable.Rows[0].Field<int>("id"));
                    price = mainDataTable.Rows[0].Field<string>("price");


                    mainDataGridView.Rows[0].Cells[0].Value = id;
                    mainDataGridView.Rows[0].Cells[1].Value = nam;
                    mainDataGridView.Rows[0].Cells[2].Value = price;
                    mainDataGridView.Rows[0].Cells[3].Value = "1";
                    //mainDataGridView.DataSource = mainDataTable;


                    mainDataGridView.Rows.Insert(0, 1);

                    calculateTotalPrice();


                }
                else
                    MessageBox.Show("لا يوجد منتج بهذا الرقم");


                mainDataTable.Clear();
                bCodeTextBox.Text = "";
                codeTextBox.Text = "";
                con.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex.Message);
                Console.WriteLine(ex.Message);
                bCodeTextBox.Text = "";
            }



        }


        void calculateTotalPrice()
        {

            totalPrice = 0;

            for (int i = 1; i < mainDataGridView.Rows.Count; i++)
            {

                if (mainDataGridView.Rows[i].Cells[2].Value != null || mainDataGridView.Rows[0].Cells[2].Value != null)
                {

                    if (totalPrice == 0)
                    {

                        totalPrice = double.Parse(mainDataGridView.Rows[1].Cells[2].Value.ToString()) * double.Parse(mainDataGridView.Rows[1].Cells[3].Value.ToString());

                    }
                    else
                    {

                        totalPrice = totalPrice + (double.Parse(mainDataGridView.Rows[i].Cells[2].Value.ToString()) * double.Parse(mainDataGridView.Rows[i].Cells[3].Value.ToString()));

                    }
                }

            }

            totalLabel.Text = totalPrice + "";

        }


        private void mainDataGridView_KeyDown(object sender, KeyEventArgs e)
        {


            if (e.KeyCode == Keys.Enter)
            {

                calculateTotalPrice();

            }


        }

        private void codeTextBox_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {

                if (codeTextBox.Text == "")
                {

                    MessageBox.Show("برجاء كتابة الكود");

                }
                else
                {

                    codeNum = codeTextBox.Text;

                    for (int i = 0; i < mainDataGridView.Rows.Count; i++)
                    {

                        if (mainDataGridView.Rows[i].Cells[0].Value != null && mainDataGridView.Rows[i].Cells[0].Value.ToString() == codeNum)
                        {

                            MessageBox.Show("المنتج موجود بالفعل, لاضافة المزيد قم بتغير العدد");
                            return;
                        }

                    }

                    commandText = "select * from products where ID =" + codeNum;
                    addProductToList();

                }


            }


        }

        private void materialButton3_Click(object sender, EventArgs e)
        {

            reset();

        }



        void reset()
        {

            mainDataGridView.Rows.Clear();
            bCodeTextBox.Text = "";
            mainDataTable.Clear();
            totalLabel.Text = "0";
            totalPrice = 0;
            commandText = "";
            bCode = "";
            codeNum = "";
            codeTextBox.Text = "";
            bCodeList.Clear();
            sGroupName = "0";
            sCash = "0";
            isReset = "0";
            available = 0;
            con.Close();

        }


        private void materialButton2_Click(object sender, EventArgs e)
        {

            //sellCommandText = "insert into sell (pid, num, price, datee) values('" + spid + "', '"+ snum + "', '" + sprice + "', '" + sDate + "')";

            if (mainDataGridView.Rows.Count > 1)
            {


                if (mainDataGridView.Rows[1].Cells[0].Value != null)
                {

                    checkProductsAvailapility();

                    if (available == 1)
                    {

                        sellProduct();

                    }



                }
                else
                {

                    MessageBox.Show("برجاء ادخال المنتجاات !!!!");

                }

            }
            else
            {

                MessageBox.Show("برجاء ادخال المنتجاات !!!!");

            }



        }


        void checkProductsAvailapility()
        {

            for (int i = 0; i < mainDataGridView.Rows.Count; i++)
            {

                if (mainDataGridView.Rows[i].Cells[0].Value != null)
                {

                    string code = mainDataGridView.Rows[i].Cells[0].Value.ToString();
                    string num = mainDataGridView.Rows[i].Cells[3].Value.ToString();


                    try
                    {


                        //to check if having that number of product

                        con.Open();

                        OleDbCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "select * from products where id =" + code;
                        cmd.ExecuteNonQuery();


                        checkProductsAvailapilitymainDataTable.Clear();

                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                        da.Fill(checkProductsAvailapilitymainDataTable);


                        con.Close();


                        if (double.Parse(checkProductsAvailapilitymainDataTable.Rows[0].Field<string>("num")) < double.Parse(num))
                        {

                            MessageBox.Show("ليس هناك عدد كافي من " + checkProductsAvailapilitymainDataTable.Rows[0].Field<string>("nam") + " في المخزن");

                            available = 0;
                            return;
                        }



                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error checkProducts " + ex.Message);
                        Console.WriteLine(ex.Message);
                    }

                }

            }

            available = 1;

        }



        void sellProduct()
        {

            double fatora = 0;
            try
            {

                OleDbCommand cmd = con.CreateCommand();


                //get last value in the last table or row 
                con.Open();

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM sell WHERE id = (SELECT max(id) FROM sell)";
                cmd.ExecuteNonQuery();

                sellProductmainDataTable.Clear();

                OleDbDataAdapter daa = new OleDbDataAdapter(cmd);
                daa.Fill(sellProductmainDataTable);


                con.Close();

                fatora = double.Parse(sellProductmainDataTable.Rows[0].Field<string>("fatora")) + 1;


                for (int i = 0; i < mainDataGridView.Rows.Count; i++)
                {


                    if (mainDataGridView.Rows[i].Cells[2].Value != null)
                    {

                        //data ti add in sell table
                        spid = mainDataGridView.Rows[i].Cells[0].Value.ToString();
                        snum = mainDataGridView.Rows[i].Cells[3].Value.ToString();
                        sprice = mainDataGridView.Rows[i].Cells[2].Value.ToString();

                        dateTime = DateTime.UtcNow.Date;
                        sDate = dateTime.ToString("dd/MM/yyyy");



                        //to get old price and other data 
                        con.Open();

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "select * from prices where pid ='" + spid + "' order by datee";
                        cmd.ExecuteNonQuery();

                        sellProductmainDataTable.Clear();

                        daa.Fill(sellProductmainDataTable);

                        string oldPrice = "";

                        con.Close();




                        List<int> pIDList = new List<int>();

                        List<string> pNumList = new List<string>();


                        for (int z = 0; z < sellProductmainDataTable.Rows.Count; z++)
                        {
                            pIDList.Add(sellProductmainDataTable.Rows[z].Field<int>("id"));
                            pNumList.Add(sellProductmainDataTable.Rows[z].Field<string>("num"));

                        }

                        double number = double.Parse(snum);

                        for (int q = 0; q < sellProductmainDataTable.Rows.Count; q++)
                        {
                            number = number - double.Parse(pNumList[q]);

                            if (number > 0)
                            {

                                pNumList[q] = 0 + "";

                            }
                            else
                            {

                                pNumList[q] = number * -1 + "";

                                break;
                            }



                        }



                        for (int q = 0; q < sellProductmainDataTable.Rows.Count; q++)
                        {

                            if (double.Parse(pNumList[q]) > 0)
                            {

                                //update row

                                con.Open();


                                cmd.CommandText = "update prices set num='" + pNumList[q] + "' where ID =" + pIDList[q];
                                cmd.ExecuteNonQuery();


                                con.Close();

                            }
                            else
                            {

                                //delete row

                                con.Open();

                                cmd.CommandText = "delete from prices where id =" + pIDList[q];
                                cmd.ExecuteNonQuery();


                                con.Close();

                            }

                            oldPrice = sellProductmainDataTable.Rows[q].Field<string>("price") + "";

                        }


                        con.Open();

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "select * from products where ID =" + spid;
                        cmd.ExecuteNonQuery();

                        sellProductmainDataTable.Clear();

                        daa.Fill(sellProductmainDataTable);


                        con.Close();





                        con.Open();


                        double finalNum = double.Parse(sellProductmainDataTable.Rows[0].Field<string>("num")) - double.Parse(snum);

                        cmd.CommandText = "update products set num='" + finalNum + "' where ID =" + spid;
                        cmd.ExecuteNonQuery();


                        con.Close();



                        //add data to sell table
                        con.Open();

                        cmd.CommandText = "insert into sell (pid, num, price, datee, oldPrice, groupp, cashh,fatora) " +
                            "values('" + spid + "'," +
                            " '" + snum + "'," +
                            " '" + sprice + "'," +
                            " '" + sDate + "'," +
                            " '" + oldPrice + "', '" + sGroupName + "', '" + sCash + "', '" + fatora + "')";

                        cmd.ExecuteNonQuery();

                        con.Close();






                    }





                }


                cmd.CommandType = CommandType.Text;

                //kml hnaa
                if (double.Parse(sCash) == 1)
                {

                    con.Open();

                    calculateTotalPrice(); 


                    cmd.CommandText = "update groupp set price= price + '" + totalPrice + "' where ID =" + sGroupName;
                    cmd.ExecuteNonQuery();


                    con.Close();

                }




                //print fatora
                if (double.Parse(isReset) == 1)
                {
                    textBox26.Text = fatora + "";
                    printt();

                }

                reset();
                MessageBox.Show("تم بنجاح");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error sProduct" + ex.Message);
                Console.WriteLine(ex.Message);
            }








        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            if (mainDataGridView.Rows.Count > 1)
            {


                if (mainDataGridView.Rows[1].Cells[0].Value != null)
                {

                    checkProductsAvailapility();

                    if (available == 1)
                    {

                        using (var form = new Form2())
                        {
                            var result = form.ShowDialog();
                            if (result == DialogResult.OK)
                            {
                                sGroupName = form.ReturnValue1;            //values preserved after close
                                sCash = form.ReturnValue2;
                                isReset = form.ReturnValue3;

                                double cancel = form.canceled;
                                if (cancel == 0)
                                {

                                    sellProduct();

                                }

                            }
                        }



                    }
                    else
                    {
                        return;

                    }


                }
                else
                {

                    MessageBox.Show("برجاء ادخال المنتجاات !!!!");

                }

            }
            else
            {

                MessageBox.Show("برجاء ادخال المنتجاات !!!!");

            }






        }

        private void materialButton4_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrWhiteSpace(textBox3.Text) || string.IsNullOrWhiteSpace(textBox4.Text) || string.IsNullOrWhiteSpace(textBox27.Text))
            {

                MessageBox.Show("برجاء ادخال جميع البيانات كاملة");

            }
            else
            {


                try
                {
                    string bcodeCheck = textBox4.Text;

                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from products where bcode ='" + bcodeCheck + "'";
                    cmd.ExecuteNonQuery();


                    materialButton4mainDataTable.Clear();

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(materialButton4mainDataTable);


                    con.Close();


                    if (materialButton4mainDataTable.Rows.Count == 0)
                    {


                        con.Open();

                        cmd.CommandText = "insert into products (nam, num, price, bcode) " +
                            "values('" + textBox1.Text + "'," +
                            " '" + textBox3.Text + "'," +
                            " '" + textBox2.Text + "'," +
                            " '" + bcodeCheck + "')";

                        cmd.ExecuteNonQuery();

                        con.Close();


                        //get p id hna
                        con.Open();

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "select * from products where bcode ='" + bcodeCheck + "'";
                        cmd.ExecuteNonQuery();


                        materialButton4mainDataTable.Clear();

                        da.Fill(materialButton4mainDataTable);


                        con.Close();



                        con.Open();

                        dateTime = DateTime.UtcNow.Date;
                        sDate = dateTime.ToString("dd/MM/yyyy");


                        cmd.CommandText = "insert into prices (pid, num, price, datee) " +
                            "values('" + materialButton4mainDataTable.Rows[0].Field<int>("id") + "'," +
                            " '" + textBox3.Text + "'," +
                            " '" + textBox27.Text + "'," +
                            " '" + sDate + "')";

                        cmd.ExecuteNonQuery();

                        con.Close();

                        //clear
                        MessageBox.Show("تم بنجاح");

                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox3.Text = "";
                        textBox4.Text = "";
                        textBox27.Text = "";


                    }
                    else
                    {
                        materialButton4mainDataTable.Clear();
                        MessageBox.Show("الباركود مسجل بالفعل!");

                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB4 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }

            }



        }

        private void materialButton5_Click(object sender, EventArgs e)
        {


            try
            {

                OleDbCommand cmd = con.CreateCommand();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);


                if ((string.IsNullOrWhiteSpace(textBox8.Text) && string.IsNullOrWhiteSpace(textBox5.Text)) || (!string.IsNullOrWhiteSpace(textBox8.Text) && !string.IsNullOrWhiteSpace(textBox5.Text)))
                {

                    MessageBox.Show("برجاء ادخال كود المنتج او قرأته باستخدام قارئ الباركود");

                }
                else
                {
                    if (string.IsNullOrWhiteSpace(textBox6.Text) && string.IsNullOrWhiteSpace(textBox7.Text))
                    {

                        MessageBox.Show("برجاء ادخال جميع البيانات كاملة");

                    }
                    else
                    {

                        if (!string.IsNullOrWhiteSpace(textBox8.Text))
                        {

                            con.Open();

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "select * from products where bcode ='" + textBox8.Text + "'";
                            cmd.ExecuteNonQuery();


                            materialButton5mainDataTable.Clear();

                            da.Fill(materialButton5mainDataTable);


                            con.Close();

                            if (materialButton5mainDataTable.Rows.Count == 0)
                            {

                                materialButton5mainDataTable.Clear();
                                MessageBox.Show("لا يوجد منتج بهذا الباركود");

                            }
                            else
                            {

                                con.Open();

                                dateTime = DateTime.UtcNow.Date;
                                sDate = dateTime.ToString("dd/MM/yyyy");


                                cmd.CommandText = "insert into prices (pid, num, price, datee) " +
                                    "values('" + materialButton5mainDataTable.Rows[0].Field<int>("id") + "'," +
                                    " '" + textBox6.Text + "'," +
                                    " '" + textBox7.Text + "'," +
                                    " '" + sDate + "')";

                                cmd.ExecuteNonQuery();

                                con.Close();

                                con.Open();

                                double ppnum = double.Parse(materialButton5mainDataTable.Rows[0].Field<string>("num")) + double.Parse(textBox6.Text);

                                cmd.CommandText = "update products set num = + '" + ppnum + "' where ID =" + materialButton5mainDataTable.Rows[0].Field<int>("id");
                                cmd.ExecuteNonQuery();


                                con.Close();



                                //clear
                                MessageBox.Show("تم بنجاح");

                                textBox5.Text = "";
                                textBox6.Text = "";
                                textBox7.Text = "";
                                textBox8.Text = "";





                            }




                        }
                        else if (!string.IsNullOrWhiteSpace(textBox5.Text))
                        {


                            con.Open();

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "select * from products where ID =" + textBox5.Text;
                            cmd.ExecuteNonQuery();


                            materialButton5mainDataTable.Clear();

                            da.Fill(materialButton5mainDataTable);


                            con.Close();

                            if (materialButton5mainDataTable.Rows.Count == 0)
                            {

                                materialButton5mainDataTable.Clear();
                                MessageBox.Show("لا يوجد منتج بهذا الكود");

                            }
                            else
                            {

                                con.Open();

                                dateTime = DateTime.UtcNow.Date;
                                sDate = dateTime.ToString("dd/MM/yyyy");


                                cmd.CommandText = "insert into prices (pid, num, price, datee) " +
                                    "values('" + materialButton5mainDataTable.Rows[0].Field<int>("id") + "'," +
                                    " '" + textBox6.Text + "'," +
                                    " '" + textBox7.Text + "'," +
                                    " '" + sDate + "')";

                                cmd.ExecuteNonQuery();

                                con.Close();

                                con.Open();

                                double ppnum = double.Parse(materialButton5mainDataTable.Rows[0].Field<string>("num")) + double.Parse(textBox6.Text);

                                cmd.CommandText = "update products set num = + '" + ppnum + "' where ID =" + materialButton5mainDataTable.Rows[0].Field<int>("id");
                                cmd.ExecuteNonQuery();


                                con.Close();



                                //clear
                                MessageBox.Show("تم بنجاح");

                                textBox5.Text = "";
                                textBox6.Text = "";
                                textBox7.Text = "";
                                textBox8.Text = "";



                            }




                        }



                    }


                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error materialB5 " + ex.Message);
                Console.WriteLine(ex.Message);
            }






        }


        private void tab1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage6"])//your specific tabname
            {
                // your stuff

                loadTab6GridViewData();


            }else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage9"])//your specific tabname
            {
                // your stuff

                loadTab7GridViewData();


            }


        }


        private void tab5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (materialTabControl1.SelectedTab == materialTabControl1.TabPages["tabPage3"])//your specific tabname
            {
                // your stuff




            }


        }



        void loadTab6GridViewData()
        {

            allDataGridView.Rows.Clear();


            try
            {

                con.Open();

                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select * from products";
                cmd.ExecuteNonQuery();


                loadTab6GridViewmainDataTable.Clear();

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(loadTab6GridViewmainDataTable);


                con.Close();

                string nam = "";
                string id = "";
                string price = "";
                string num = "";

                for (int i = 0; i < loadTab6GridViewmainDataTable.Rows.Count; i++)
                {

                    nam = loadTab6GridViewmainDataTable.Rows[i].Field<string>("nam");
                    id = Convert.ToString(loadTab6GridViewmainDataTable.Rows[i].Field<int>("id"));
                    price = loadTab6GridViewmainDataTable.Rows[i].Field<string>("price");
                    num = loadTab6GridViewmainDataTable.Rows[i].Field<string>("num");

                    allDataGridView.Rows.Insert(i, 1);

                    allDataGridView.Rows[i].Cells[0].Value = id;
                    allDataGridView.Rows[i].Cells[1].Value = nam;
                    allDataGridView.Rows[i].Cells[2].Value = price;
                    allDataGridView.Rows[i].Cells[3].Value = num;



                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loadT6 " + ex.Message);
                Console.WriteLine(ex.Message);
            }


        }

        double selectedCode = 0;

        private void materialButton6_Click(object sender, EventArgs e)
        {

            if (!string.IsNullOrWhiteSpace(textBox9.Text))
            {

                //abd2 al search
                try
                {
                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from products where bcode ='" + textBox9.Text + "'";
                    cmd.ExecuteNonQuery();


                    materialButton6mainDataTable.Clear();

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(materialButton6mainDataTable);


                    con.Close();

                    if (materialButton6mainDataTable.Rows.Count > 0)
                    {
                        selectedCode = materialButton6mainDataTable.Rows[0].Field<int>("ID");
                        textBox13.Text = materialButton6mainDataTable.Rows[0].Field<string>("nam");
                        textBox10.Text = materialButton6mainDataTable.Rows[0].Field<string>("bcode");
                        textBox12.Text = materialButton6mainDataTable.Rows[0].Field<string>("price");

                    }
                    else
                    {
                        MessageBox.Show("لا يوجد منتج بهذا الكود");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB6 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }



            }
            else
            {

                MessageBox.Show("برجاء ادخال او قرأة الباركود");

            }


        }


        

        DataRow[] rows;

        private void allDataGridView_SelectionChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in allDataGridView.SelectedRows)
            {
                if (row.Cells[0].Value != null) { 

                selectedCode = double.Parse(row.Cells[0].Value.ToString());
                    textBox13.Text = row.Cells[1].Value.ToString();
                    textBox12.Text = row.Cells[2].Value.ToString();

                    rows = loadTab6GridViewmainDataTable.Select("ID = " + row.Cells[0].Value.ToString());

                    textBox10.Text = rows[0].Field<string>("bcode");
                }
                    



            }
        }

        private void materialButton7_Click(object sender, EventArgs e)
        {


            if (string.IsNullOrWhiteSpace(textBox13.Text) || string.IsNullOrWhiteSpace(textBox10.Text) || string.IsNullOrWhiteSpace(textBox12.Text))
            {

                MessageBox.Show("برجاء اختيار منتج من الجدول او البحث عنه لتعديله");

            }
            else
            {

                try
                {
                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "update products set nam='" + textBox13.Text + "', bcode='" + textBox10.Text + "', price ='" + textBox12.Text + "' where ID =" + selectedCode;
                    cmd.ExecuteNonQuery();


                    con.Close();

                    loadTab6GridViewData();

                    MessageBox.Show("تم بنجاح");


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB7 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }





            }




        }




        private void materialButton8_Click(object sender, EventArgs e)
        {





            if (string.IsNullOrWhiteSpace(textBox13.Text) || string.IsNullOrWhiteSpace(textBox10.Text) || string.IsNullOrWhiteSpace(textBox12.Text))
            {

                MessageBox.Show("برجاء اختيار منتج من الجدول او البحث عنه لمسحه");

            }
            else
            {

                string message = "سيتم مسح جميع بيانات هذا المنتج واعداده واسعاره السابقة ؟؟";
                string title = "مسح المنتج";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result = MessageBox.Show(message, title, buttons);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        con.Open();

                        OleDbCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "delete from products where bcode ='" + textBox10.Text + "'";
                        cmd.ExecuteNonQuery();


                        con.Close();


                        con.Open();

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "delete from prices where pid ='" + selectedCode + "'";
                        cmd.ExecuteNonQuery();


                        con.Close();


                        loadTab6GridViewData();

                        textBox13.Text = "";
                        textBox10.Text = "";
                        textBox12.Text = "";


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error materialB8 " + ex.Message);
                        Console.WriteLine(ex.Message);
                    }


                }
                else
                {
                    // Do something  
                }




            }





        }


        double selectedCode2 = 0;
        DataRow[] rows2;
        System.Data.DataTable tab7NameMainDataTable = new System.Data.DataTable();
        double gridNum = 0;
        double gridNumId = 0;


        private void allDataPricesGridView_SelectionChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in allDataPricesGridView.SelectedRows)
            {

                if (row.Cells[0].Value != null) { 
               
                    selectedCode2 = double.Parse(row.Cells[0].Value.ToString());

                    textBox15.Text = row.Cells[1].Value.ToString();
                    textBox11.Text = row.Cells[2].Value.ToString();
                    textBox14.Text = row.Cells[3].Value.ToString();

                    rows2 = mainDataTabletab7Grid.Select("ID = " + row.Cells[0].Value.ToString());

                    gridNum = double.Parse(row.Cells[3].Value.ToString());

                    gridNumId = double.Parse(row.Cells[0].Value.ToString());

                }




            }
        }


        System.Data.DataTable mainDataTabletab7Grid = new System.Data.DataTable();


        void loadTab7GridViewData()
        {

            allDataPricesGridView.Rows.Clear();



            try
            {

                con.Open();

                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select * from prices";
                cmd.ExecuteNonQuery();


                mainDataTabletab7Grid.Clear();

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(mainDataTabletab7Grid);


                con.Close();

                string nam = "";
                string id = "";
                string price = "";
                string num = "";
                string datee = "";

                for (int i = 0; i < mainDataTabletab7Grid.Rows.Count; i++)
                {


                    con.Open();

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from products where ID =" + mainDataTabletab7Grid.Rows[i].Field<string>("pid");
                    cmd.ExecuteNonQuery();


                    tab7NameMainDataTable.Clear();

                    da.Fill(tab7NameMainDataTable);

                    con.Close();


                    id = mainDataTabletab7Grid.Rows[i].Field<int>("ID") + "";

                    nam = tab7NameMainDataTable.Rows[0].Field<string>("nam");
                    price = mainDataTabletab7Grid.Rows[i].Field<int>("price") + "";
                    num = mainDataTabletab7Grid.Rows[i].Field<string>("num");
                    datee = mainDataTabletab7Grid.Rows[i].Field<DateTime>("datee").ToString("dd/MM/yyyy");

                    allDataPricesGridView.Rows.Insert(i, 1);

                    allDataPricesGridView.Rows[i].Cells[0].Value = id;
                    allDataPricesGridView.Rows[i].Cells[1].Value = nam;
                    allDataPricesGridView.Rows[i].Cells[2].Value = price;
                    allDataPricesGridView.Rows[i].Cells[3].Value = num;
                    allDataPricesGridView.Rows[i].Cells[4].Value = datee;



                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loadT7G " + ex.Message);
                Console.WriteLine(ex.Message);
            }


        }


        System.Data.DataTable mainDataTableproductsrows = new System.Data.DataTable();



        private void materialButton10_Click(object sender, EventArgs e)
        {





            if (string.IsNullOrWhiteSpace(textBox15.Text) || string.IsNullOrWhiteSpace(textBox11.Text) || string.IsNullOrWhiteSpace(textBox14.Text))
            {

                MessageBox.Show("برجاء اختيار منتج من الجدول لتعديله");

            }
            else
            {

                try
                {


                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "update prices set num='" + textBox14.Text + "', price='" + textBox11.Text + "' where ID =" + selectedCode2;
                    cmd.ExecuteNonQuery();


                    con.Close();






                    con.Open();

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from products where id =" + rows2[0].Field<string>("pid"); //main num
                    cmd.ExecuteNonQuery();


                    mainDataTableproductsrows.Clear();

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(mainDataTableproductsrows);


                    con.Close();


                    //          num in grid      Entered num  
                    double newCount = System.Math.Abs(gridNum - double.Parse(textBox14.Text));
                    double mainNum = double.Parse(mainDataTableproductsrows.Rows[0].Field<string>("num"));

                    MessageBox.Show("تم بنجاح" + newCount);

                    if (double.Parse(textBox14.Text) > gridNum)
                    {

                        //increase main num

                        con.Open();

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "update products set num='" + (mainNum + newCount) + "' where ID =" + rows2[0].Field<string>("pid");
                        cmd.ExecuteNonQuery();


                        con.Close();


                    }
                    else if (double.Parse(textBox14.Text) < gridNum)
                    {

                        //decrease main num

                        con.Open();

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "update products set num='" + (mainNum - newCount) + "' where ID =" + rows2[0].Field<string>("pid");
                        cmd.ExecuteNonQuery();


                        con.Close();

                    }


                    loadTab7GridViewData();

                    MessageBox.Show("تم بنجاح");

                    textBox15.Text = "";
                    textBox11.Text = "";
                    textBox14.Text = "";


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB10 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }





            }







        }




        private void materialButton9_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textBox15.Text) || string.IsNullOrWhiteSpace(textBox11.Text) || string.IsNullOrWhiteSpace(textBox14.Text))
            {

                MessageBox.Show("برجاء اختيار منتج من الجدول لمسحه");

            }
            else
            {

                try
                {
                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "delete from prices where ID =" + gridNumId;
                    cmd.ExecuteNonQuery();


                    con.Close();

                    con.Open();

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from products where id =" + rows2[0].Field<string>("pid"); //main num
                    cmd.ExecuteNonQuery();


                    mainDataTableproductsrows.Clear();

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(mainDataTableproductsrows);


                    con.Close();




                    con.Open();

                    double mainNum = double.Parse(mainDataTableproductsrows.Rows[0].Field<string>("num"));


                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "update products set num='" + (mainNum - gridNum) + "' where ID =" + rows2[0].Field<string>("pid");
                    cmd.ExecuteNonQuery();


                    con.Close();


                    loadTab7GridViewData();

                    textBox15.Text = "";
                    textBox11.Text = "";
                    textBox14.Text = "";


                    MessageBox.Show("تم بنجاح");


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB9 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }





            }





        }



        private void materialButton11_Click(object sender, EventArgs e)
        {



            if (string.IsNullOrWhiteSpace(textBox19.Text) || string.IsNullOrWhiteSpace(textBox16.Text))
            {

                MessageBox.Show("برجاء ادخال جميع البيانات كاملة");

            }
            else
            {


                try
                {

                    con.Open();


                    OleDbCommand cmd = con.CreateCommand();

                    cmd.CommandText = "insert into groupp (nam, phnum) " +
                            "values('" + textBox19.Text + "'," +
                            " '" + textBox16.Text + "')";

                    cmd.ExecuteNonQuery();

                    con.Close();

                    //clear
                    MessageBox.Show("تم بنجاح");

                    textBox19.Text = "";
                    textBox16.Text = "";



                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB11 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }

            }


        }



        private void tab2Group_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage11"])//your specific tabname
            {
                // your stuff

                loadTab2GroupGridViewData();


            } else if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage12"])//your specific tabname
            {
                // your stuff

                



                // loadTab2GroupGridViewData();


            }


        }


        System.Data.DataTable mainDataTabletab2GroupGrid = new System.Data.DataTable();


        void loadTab2GroupGridViewData()
        {

            groupDataGridView.Rows.Clear();



            try
            {

                con.Open();

                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select * from groupp";
                cmd.ExecuteNonQuery();


                mainDataTabletab2GroupGrid.Clear();

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(mainDataTabletab2GroupGrid);


                con.Close();

                double idd = 0;
                string nam = "";
                string price = "";
                string num = "";

                for (int i = 0; i < mainDataTabletab2GroupGrid.Rows.Count; i++)
                {



                    idd = mainDataTabletab2GroupGrid.Rows[i].Field<int>("ID");
                    nam = mainDataTabletab2GroupGrid.Rows[i].Field<string>("nam");
                    price = mainDataTabletab2GroupGrid.Rows[i].Field<int>("price") + "";
                    num = mainDataTabletab2GroupGrid.Rows[i].Field<string>("phnum");

                    groupDataGridView.Rows.Insert(i, 1);

                    groupDataGridView.Rows[i].Cells[0].Value = idd;
                    groupDataGridView.Rows[i].Cells[1].Value = nam;
                    groupDataGridView.Rows[i].Cells[2].Value = num;
                    groupDataGridView.Rows[i].Cells[3].Value = price;




                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loadT2Grp " + ex.Message);
                Console.WriteLine(ex.Message);
            }


        }




        private void groupDataGridView_SelectionChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in groupDataGridView.SelectedRows)
            {


                textBox22.Text = row.Cells[0].Value.ToString();
                textBox20.Text = row.Cells[1].Value.ToString();
                textBox17.Text = row.Cells[2].Value.ToString();
                textBox18.Text = row.Cells[3].Value.ToString();



            }
        }


        System.Data.DataTable searchNameDataTable = new System.Data.DataTable();


        private void materialButton14_Click(object sender, EventArgs e)
        {



            if (!string.IsNullOrWhiteSpace(textBox21.Text))
            {

                //abd2 al search
                try
                {
                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from groupp where nam ='" + textBox21.Text + "'";
                    cmd.ExecuteNonQuery();


                    searchNameDataTable.Clear();

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(searchNameDataTable);


                    con.Close();

                    if (searchNameDataTable.Rows.Count > 0)
                    {

                        textBox22.Text = searchNameDataTable.Rows[0].Field<int>("id") + "";
                        textBox20.Text = searchNameDataTable.Rows[0].Field<string>("nam");
                        textBox17.Text = searchNameDataTable.Rows[0].Field<string>("phnum");
                        textBox18.Text = searchNameDataTable.Rows[0].Field<int>("price") + "";

                    }
                    else
                    {
                        MessageBox.Show("لا توجد مجموعة بهذا الاسم");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB14 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }



            }
            else
            {

                MessageBox.Show("برجاء ادخال اسم المجموعة");

            }



        }

        private void materialButton13_Click(object sender, EventArgs e)
        {




            if (string.IsNullOrWhiteSpace(textBox22.Text) || string.IsNullOrWhiteSpace(textBox20.Text) || string.IsNullOrWhiteSpace(textBox17.Text) || string.IsNullOrWhiteSpace(textBox18.Text))
            {

                MessageBox.Show("برجاء اختيار مجموعة من الجدول او البحث عنه لتعديله");

            }
            else
            {

                try
                {
                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "update groupp set nam='" + textBox20.Text + "', phnum='" + textBox17.Text + "', price =" + textBox18.Text + " where ID =" + textBox22.Text;
                    cmd.ExecuteNonQuery();


                    con.Close();

                    loadTab2GroupGridViewData();

                    MessageBox.Show("تم بنجاح");


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB13 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }





            }



        }

        private void materialButton12_Click(object sender, EventArgs e)
        {



            if (string.IsNullOrWhiteSpace(textBox22.Text) || string.IsNullOrWhiteSpace(textBox20.Text) || string.IsNullOrWhiteSpace(textBox17.Text) || string.IsNullOrWhiteSpace(textBox18.Text))
            {

                MessageBox.Show("برجاء اختيار مجموعة من الجدول او البحث عنه لمسحه");

            }
            else
            {

                try
                {
                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "delete from groupp where ID =" + textBox22.Text;
                    cmd.ExecuteNonQuery();


                    con.Close();


                    loadTab2GroupGridViewData();

                    MessageBox.Show("تم بنجاح");

                    textBox22.Text = "";
                    textBox20.Text = "";
                    textBox17.Text = "";
                    textBox18.Text = "";


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error materialB12 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }


            }



        }



        private void button1_Click(object sender, EventArgs e)
        {

            var myForm = new Form3();
            myForm.Show();


        }



        System.Data.DataTable mainDataTableDataReport = new System.Data.DataTable();
        System.Data.DataTable mainDataTableDataReportname = new System.Data.DataTable();


        private void button2_Click(object sender, EventArgs e)
        {


            reportsDataGridView.Rows.Clear();



            try
            {

                con.Open();

                dateTime = DateTime.UtcNow.Date;
                sDate = dateTime.ToString("dd/MM/yyyy");




                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                // cmd.CommandText = "select * from sell WHERE DATEDIFF('D', datee, NOW) > 90";
                cmd.CommandText = "select * from sell WHERE datee = " + DateTime.Parse(sDate).ToOADate();

                cmd.ExecuteNonQuery();


                mainDataTableDataReport.Clear();

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(mainDataTableDataReport);


                con.Close();

                int id = 0;
                string nam = "";
                string num = "";
                string oldPrice = "";
                string newPrice = "";
                string datee = "";


                for (int i = 0; i < mainDataTableDataReport.Rows.Count; i++)
                {



                    con.Open();

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from products where id =" + mainDataTableDataReport.Rows[i].Field<string>("pid");
                    cmd.ExecuteNonQuery();

                    mainDataTableDataReportname.Clear();

                    da.Fill(mainDataTableDataReportname);


                    con.Close();



                    id = Int16.Parse(mainDataTableDataReport.Rows[i].Field<string>("fatora"));

                    nam = mainDataTableDataReportname.Rows[0].Field<string>("nam");

                    num = mainDataTableDataReport.Rows[i].Field<string>("num");
                    oldPrice = mainDataTableDataReport.Rows[i].Field<string>("oldPrice");

                    newPrice = mainDataTableDataReport.Rows[i].Field<string>("price");
                    datee = mainDataTableDataReport.Rows[i].Field<DateTime>("datee").ToString("dd/MM/yyyy");

                    reportsDataGridView.Rows.Insert(i, 1);

                    reportsDataGridView.Rows[i].Cells[0].Value = id;
                    reportsDataGridView.Rows[i].Cells[1].Value = nam;
                    reportsDataGridView.Rows[i].Cells[2].Value = num;
                    reportsDataGridView.Rows[i].Cells[3].Value = oldPrice;
                    reportsDataGridView.Rows[i].Cells[4].Value = newPrice;
                    reportsDataGridView.Rows[i].Cells[5].Value = datee;





                }

                double calcNum = 0;
                for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                {
                    if (reportsDataGridView.Rows[z].Cells[2].Value != null)
                    {
                        calcNum = calcNum + double.Parse(reportsDataGridView.Rows[z].Cells[2].Value.ToString());
                    }
                }

                textBox23.Text = calcNum + "";


                double priceNNum = 0;
                for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                {
                    if (reportsDataGridView.Rows[z].Cells[4].Value != null)
                    {
                        priceNNum = priceNNum + (double.Parse(reportsDataGridView.Rows[z].Cells[4].Value.ToString()) * double.Parse(reportsDataGridView.Rows[z].Cells[2].Value.ToString()));
                    }
                }

                textBox24.Text = priceNNum + "";

                double priceONum = 0;
                for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                {
                    if (reportsDataGridView.Rows[z].Cells[3].Value != null)
                    {
                        priceONum = priceONum + double.Parse(reportsDataGridView.Rows[z].Cells[3].Value.ToString());
                    }
                }

                textBox25.Text = priceNNum - priceONum + "";


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error B2 " + ex.Message);
                Console.WriteLine(ex.Message);
            }



        }

        private void button3_Click(object sender, EventArgs e)
        {


            {


                reportsDataGridView.Rows.Clear();



                try
                {

                    con.Open();

                    dateTime = DateTime.Now.AddMonths(-1); ;
                    sDate = dateTime.ToString("dd/MM/yyyy");




                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    // cmd.CommandText = "select * from sell WHERE DATEDIFF('D', datee, NOW) > 90";
                    cmd.CommandText = "select * from sell WHERE datee > " + DateTime.Parse(sDate).ToOADate();

                    cmd.ExecuteNonQuery();


                    mainDataTableDataReport.Clear();

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(mainDataTableDataReport);


                    con.Close();

                    int id = 0;
                    string nam = "";
                    string num = "";
                    string oldPrice = "";
                    string newPrice = "";
                    string datee = "";


                    for (int i = 0; i < mainDataTableDataReport.Rows.Count; i++)
                    {



                        con.Open();

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "select * from products where id =" + mainDataTableDataReport.Rows[i].Field<string>("pid");
                        cmd.ExecuteNonQuery();

                        mainDataTableDataReportname.Clear();

                        da.Fill(mainDataTableDataReportname);


                        con.Close();



                        id = Int16.Parse(mainDataTableDataReport.Rows[i].Field<string>("fatora"));

                        nam = mainDataTableDataReportname.Rows[0].Field<string>("nam");

                        num = mainDataTableDataReport.Rows[i].Field<string>("num");
                        oldPrice = mainDataTableDataReport.Rows[i].Field<string>("oldPrice");

                        newPrice = mainDataTableDataReport.Rows[i].Field<string>("price");
                        datee = mainDataTableDataReport.Rows[i].Field<DateTime>("datee").ToString("dd/MM/yyyy");

                        reportsDataGridView.Rows.Insert(i, 1);

                        reportsDataGridView.Rows[i].Cells[0].Value = id;
                        reportsDataGridView.Rows[i].Cells[1].Value = nam;
                        reportsDataGridView.Rows[i].Cells[2].Value = num;
                        reportsDataGridView.Rows[i].Cells[3].Value = oldPrice;
                        reportsDataGridView.Rows[i].Cells[4].Value = newPrice;
                        reportsDataGridView.Rows[i].Cells[5].Value = datee;





                    }

                    double calcNum = 0;
                    for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                    {
                        if (reportsDataGridView.Rows[z].Cells[2].Value != null)
                        {
                            calcNum = calcNum + double.Parse(reportsDataGridView.Rows[z].Cells[2].Value.ToString());
                        }
                    }

                    textBox23.Text = calcNum + "";


                    double priceNNum = 0;
                    for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                    {
                        if (reportsDataGridView.Rows[z].Cells[4].Value != null)
                        {
                            priceNNum = priceNNum + (double.Parse(reportsDataGridView.Rows[z].Cells[4].Value.ToString()) * double.Parse(reportsDataGridView.Rows[z].Cells[2].Value.ToString()));
                        }
                    }

                    textBox24.Text = priceNNum + "";

                    double priceONum = 0;
                    for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                    {
                        if (reportsDataGridView.Rows[z].Cells[3].Value != null)
                        {
                            priceONum = priceONum + double.Parse(reportsDataGridView.Rows[z].Cells[3].Value.ToString());
                        }
                    }

                    textBox25.Text = priceNNum - priceONum + "";


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error B3 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }



            }



        }

        private void button4_Click(object sender, EventArgs e)
        {



            reportsDataGridView.Rows.Clear();



            try
            {

                con.Open();

                dateTime = DateTime.Now.AddMonths(-12); ;
                sDate = dateTime.ToString("dd/MM/yyyy");




                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                // cmd.CommandText = "select * from sell WHERE DATEDIFF('D', datee, NOW) > 90";
                cmd.CommandText = "select * from sell WHERE datee > " + DateTime.Parse(sDate).ToOADate();

                cmd.ExecuteNonQuery();


                mainDataTableDataReport.Clear();

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(mainDataTableDataReport);


                con.Close();

                int id = 0;
                string nam = "";
                string num = "";
                string oldPrice = "";
                string newPrice = "";
                string datee = "";


                for (int i = 0; i < mainDataTableDataReport.Rows.Count; i++)
                {



                    con.Open();

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from products where id =" + mainDataTableDataReport.Rows[i].Field<string>("pid");
                    cmd.ExecuteNonQuery();

                    mainDataTableDataReportname.Clear();

                    da.Fill(mainDataTableDataReportname);


                    con.Close();



                    id = Int16.Parse(mainDataTableDataReport.Rows[i].Field<string>("fatora"));

                    nam = mainDataTableDataReportname.Rows[0].Field<string>("nam");

                    num = mainDataTableDataReport.Rows[i].Field<string>("num");
                    oldPrice = mainDataTableDataReport.Rows[i].Field<string>("oldPrice");

                    newPrice = mainDataTableDataReport.Rows[i].Field<string>("price");
                    datee = mainDataTableDataReport.Rows[i].Field<DateTime>("datee").ToString("dd/MM/yyyy");

                    reportsDataGridView.Rows.Insert(i, 1);

                    reportsDataGridView.Rows[i].Cells[0].Value = id;
                    reportsDataGridView.Rows[i].Cells[1].Value = nam;
                    reportsDataGridView.Rows[i].Cells[2].Value = num;
                    reportsDataGridView.Rows[i].Cells[3].Value = oldPrice;
                    reportsDataGridView.Rows[i].Cells[4].Value = newPrice;
                    reportsDataGridView.Rows[i].Cells[5].Value = datee;





                }

                double calcNum = 0;
                for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                {
                    if (reportsDataGridView.Rows[z].Cells[2].Value != null)
                    {
                        calcNum = calcNum + double.Parse(reportsDataGridView.Rows[z].Cells[2].Value.ToString());
                    }
                }

                textBox23.Text = calcNum + "";


                double priceNNum = 0;
                for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                {
                    if (reportsDataGridView.Rows[z].Cells[4].Value != null)
                    {
                        priceNNum = priceNNum + (double.Parse(reportsDataGridView.Rows[z].Cells[4].Value.ToString()) * double.Parse(reportsDataGridView.Rows[z].Cells[2].Value.ToString()));
                    }
                }

                textBox24.Text = priceNNum + "";

                double priceONum = 0;
                for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                {
                    if (reportsDataGridView.Rows[z].Cells[3].Value != null)
                    {
                        priceONum = priceONum + double.Parse(reportsDataGridView.Rows[z].Cells[3].Value.ToString());
                    }
                }

                textBox25.Text = priceNNum - priceONum + "";


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error B4 " + ex.Message);
                Console.WriteLine(ex.Message);
            }



        }




        double fatoraN = 0;

        private void reportsDataGridView_SelectionChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in reportsDataGridView.SelectedRows)
            {

                if(row.Cells[0].Value != null)
                {
                    fatoraN = double.Parse(row.Cells[0].Value.ToString());
                    textBox26.Text = row.Cells[0].Value.ToString();

                } 



            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

            fillReportTablewithFatora();



        }


        void fillReportTablewithFatora(){



            if (!string.IsNullOrWhiteSpace(textBox26.Text))
            {

                reportsDataGridView.Rows.Clear();


                //abd2 al search
                try
                {
                    con.Open();

 
                    OleDbCommand cmd = con.CreateCommand();
                      cmd.CommandType = CommandType.Text;
                    // cmd.CommandText = "select * from sell WHERE DATEDIFF('D', datee, NOW) > 90";
                    cmd.CommandText = "select * from sell WHERE fatora ='" + textBox26.Text + "'";

                    cmd.ExecuteNonQuery();


                    mainDataTableDataReport.Clear();

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        da.Fill(mainDataTableDataReport);


                    con.Close();

                    int id = 0;
        string nam = "";
        string num = "";
        string oldPrice = "";
        string newPrice = "";
        string datee = "";

                    if (mainDataTableDataReport.Rows.Count > 0) { 
                        
                        for (int i = 0; i<mainDataTableDataReport.Rows.Count; i++)
                                 {


                          //  if (mainDataTableDataReport.Rows[0]) { 
                        con.Open();

                            cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "select * from products where id =" + mainDataTableDataReport.Rows[i].Field<string>("pid");
                        cmd.ExecuteNonQuery();

                        mainDataTableDataReportname.Clear();

                            da.Fill(mainDataTableDataReportname);


                        con.Close();



                            id = Int16.Parse(mainDataTableDataReport.Rows[i].Field<string>("fatora"));
                            nam = mainDataTableDataReportname.Rows[0].Field<string>("nam");
                            num = mainDataTableDataReport.Rows[i].Field<string>("num");
                        oldPrice = mainDataTableDataReport.Rows[i].Field<string>("oldPrice");

                            newPrice = mainDataTableDataReport.Rows[i].Field<string>("price");
                        datee = mainDataTableDataReport.Rows[i].Field<DateTime>("datee").ToString("dd/MM/yyyy");

        reportsDataGridView.Rows.Insert(i, 1);

                            reportsDataGridView.Rows[i].Cells[0].Value = id;
                        reportsDataGridView.Rows[i].Cells[1].Value = nam;
                        reportsDataGridView.Rows[i].Cells[2].Value = num;
                        reportsDataGridView.Rows[i].Cells[3].Value = oldPrice;
                        reportsDataGridView.Rows[i].Cells[4].Value = newPrice;
                        reportsDataGridView.Rows[i].Cells[5].Value = datee;




                        }
    //  }
}
                    else
{

    MessageBox.Show("لا يوجد فاتورة بهذا الرقم ");

}

                }
                catch (Exception ex)
{
    MessageBox.Show("Error B6 " + ex.Message);
    Console.WriteLine(ex.Message);
}



            }
            else
{

    MessageBox.Show("برجاء كتابة رقم الفاتورة او اختيارها من الجدول");

}
            
            }


        System.Data.DataTable dataTableforName = new System.Data.DataTable();

        private void button5_Click(object sender, EventArgs e)
        {
            reportsDataGridView.Rows.Clear();

            printt();

        }

        void printt() {


            try
            {

                if (!string.IsNullOrWhiteSpace(textBox26.Text))
                {
                    fillReportTablewithFatora();



                    String path = @"E:\Debug\";
                    _Application excel = new _Excel.Application();
                    Workbook wb;
                    Worksheet ws;

                    wb = excel.Workbooks.Open(path + "template.xlsx");
                    ws = wb.Worksheets[1];

                    // [ t7t, ->]
                    // ws.Cells[1,2] = 555;

                    //get data of the sheeeeetttt


                    ws.Cells[4, 6] = mainDataTableDataReport.Rows[0].Field<string>("fatora");
                    ws.Cells[4, 1] = mainDataTableDataReport.Rows[0].Field<DateTime>("datee").ToString("dd/MM/yyyy");

                    //to get group name
                    if (double.Parse(mainDataTableDataReport.Rows[0].Field<string>("fatora")) != 0)
                    {

                        con.Open();

                        OleDbCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "select * from groupp where id =" + mainDataTableDataReport.Rows[0].Field<string>("groupp");
                        cmd.ExecuteNonQuery();


                        dataTableforName.Clear();

                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                        da.Fill(dataTableforName);


                        con.Close();



                        ws.Cells[7, 1] = dataTableforName.Rows[0].Field<string>("nam");


                    }
                    else
                    {

                        ws.Cells[7, 1] = "افراد";

                    }

                    //to get payed cash or not
                    if (double.Parse(mainDataTableDataReport.Rows[0].Field<string>("cashh")) != 0)
                    {

                        ws.Cells[7, 6] = "آجل";
                    }
                    else
                    {

                        ws.Cells[7, 6] = "كاش";
                    }

                    //5ant al total
                    //ws.Cells[10, 2] = textBox24.Text;

                    double priceNNum = 0;
                    for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                    {
                        if (reportsDataGridView.Rows[z].Cells[4].Value != null)
                        {
                            priceNNum = priceNNum + (double.Parse(reportsDataGridView.Rows[z].Cells[4].Value.ToString()) * double.Parse(reportsDataGridView.Rows[z].Cells[2].Value.ToString()));
                        }
                    }

                    ws.Cells[10, 2] = priceNNum;

                    //to fill the table
                    for (int i = 0; i < mainDataTableDataReport.Rows.Count; i++)
                    {

                        ws.Cells[13 + i, 8] = i + 1;
                        ws.Cells[13 + i, 6] = reportsDataGridView.Rows[i].Cells[1].Value.ToString();
                        ws.Cells[13 + i, 5] = mainDataTableDataReport.Rows[i].Field<string>("num");
                        ws.Cells[13 + i, 3] = mainDataTableDataReport.Rows[i].Field<string>("price");
                        ws.Cells[13 + i, 1] = double.Parse(mainDataTableDataReport.Rows[i].Field<string>("price")) * double.Parse(mainDataTableDataReport.Rows[i].Field<string>("num"));


                    }
                    // wb.Save();

                    wb.SaveAs(path + @"fwatir\" + textBox26.Text + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);
                    wb.Close();
                    excel.Quit();


                    ///////////


                    string filePath = path + @"fwatir\" + textBox26.Text + ".xlsx";
                    
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();



                    // Open the Workbook:
                    wb = excelApp.Workbooks.Open(
                      path + @"fwatir\" + textBox26.Text + ".xlsx",
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing);


                    ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

                    ws.PrintOut(
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();

                    //Marshal.FinalReleaseComObject(ws);

                    wb.Close(false, Type.Missing, Type.Missing);
                    Marshal.FinalReleaseComObject(wb);

                    excelApp.Quit();
                    Marshal.FinalReleaseComObject(excelApp);





                    MessageBox.Show("تم بنجاح");


                }
                else
                {

                    MessageBox.Show("برجاء كتابة رقم الفاتورة او اختيارها من الجدول");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error B6 " + ex.Message);
                Console.WriteLine(ex.Message);
            }



        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

        }


        double clicked = 0;
        private void button7_Click(object sender, EventArgs e)
        {
            if (clicked == 0) {

                try
                {


                    con.Open();

                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select * from groupp";
                    cmd.ExecuteNonQuery();

                    dataTableOfComboBox.Clear();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dataTableOfComboBox);

                    con.Close();


                    comboBox1.DisplayMember = "nam";
                    comboBox1.DataSource = dataTableOfComboBox;




                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error " + ex.Message);
                    Console.WriteLine(ex.Message);
                }

                clicked = 1;

            }else
            {

                reportsDataGridView.Rows.Clear();


                try
                {
                    con.Open();


                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    // cmd.CommandText = "select * from sell WHERE DATEDIFF('D', datee, NOW) > 90";
                    cmd.CommandText = "select * from sell WHERE groupp ='" + dataTableOfComboBox.Rows[comboBox1.SelectedIndex].Field<int>("id") + "'";

                    cmd.ExecuteNonQuery();


                    mainDataTableDataReport.Clear();

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(mainDataTableDataReport);


                    con.Close();

                    int id = 0;
                    string nam = "";
                    string num = "";
                    string oldPrice = "";
                    string newPrice = "";
                    string datee = "";
                    string csh = "";

                    if (mainDataTableDataReport.Rows.Count > 0)
                    {

                        for (int i = 0; i < mainDataTableDataReport.Rows.Count; i++)
                        {


                            //  if (mainDataTableDataReport.Rows[0]) { 
                            con.Open();

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "select * from products where id =" + mainDataTableDataReport.Rows[i].Field<string>("pid");
                            cmd.ExecuteNonQuery();

                            mainDataTableDataReportname.Clear();

                            da.Fill(mainDataTableDataReportname);


                            con.Close();



                            id = Int16.Parse(mainDataTableDataReport.Rows[i].Field<string>("fatora"));
                            nam = mainDataTableDataReportname.Rows[0].Field<string>("nam");
                            num = mainDataTableDataReport.Rows[i].Field<string>("num");
                            oldPrice = mainDataTableDataReport.Rows[i].Field<string>("oldPrice");

                            newPrice = mainDataTableDataReport.Rows[i].Field<string>("price");
                            datee = mainDataTableDataReport.Rows[i].Field<DateTime>("datee").ToString("dd/MM/yyyy");
                            csh = mainDataTableDataReport.Rows[i].Field<string>("cashh");

                            reportsDataGridView.Rows.Insert(i, 1);

                            reportsDataGridView.Rows[i].Cells[0].Value = id;
                            reportsDataGridView.Rows[i].Cells[1].Value = nam;
                            reportsDataGridView.Rows[i].Cells[2].Value = num;
                            reportsDataGridView.Rows[i].Cells[3].Value = oldPrice;
                            reportsDataGridView.Rows[i].Cells[4].Value = newPrice;
                            reportsDataGridView.Rows[i].Cells[5].Value = datee;

                            if (double.Parse(csh) == 0) {

                                reportsDataGridView.Rows[i].Cells[6].Value = "لا";

                            }
                            else {

                                reportsDataGridView.Rows[i].Cells[6].Value = "نعم";

                            }




                        }
                        //  }
                    }
                    else
                    {

                        MessageBox.Show("لا يوجد فاتورة بهذا الرقم ");

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error B6 " + ex.Message);
                    Console.WriteLine(ex.Message);
                }



            }




        }

        private void button8_Click(object sender, EventArgs e)
        {

            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;

            try
            {

                if (!string.IsNullOrWhiteSpace(textBox26.Text)){

                    //here

                    reportsDataGridView.Rows.Clear();

                    fillReportTablewithFatora();


                    for (int i = 0; i < mainDataTableDataReport.Rows.Count; i++)
                    {

                        //name      reportsDataGridView.Rows[i].Cells[1].Value.ToString();
                        // al3dd       mainDataTableDataReport.Rows[i].Field<string>("num");



                        con.Open();


                        OleDbDataAdapter daa = new OleDbDataAdapter(cmd);


                        cmd.CommandText = "select * from products where nam ='" + reportsDataGridView.Rows[i].Cells[1].Value.ToString() + "'";
                        cmd.ExecuteNonQuery();

                        sellProductmainDataTable.Clear();

                        daa.Fill(sellProductmainDataTable);


                        con.Close();



                        con.Open();


                        double finalNum = double.Parse(sellProductmainDataTable.Rows[0].Field<string>("num")) + double.Parse(mainDataTableDataReport.Rows[i].Field<string>("num"));

                        cmd.CommandText = "update products set num='" + finalNum + "' where nam ='" + reportsDataGridView.Rows[i].Cells[1].Value.ToString() + "'";
                        cmd.ExecuteNonQuery();


                        con.Close();



                    }


                    //to get group name
                    if (double.Parse(mainDataTableDataReport.Rows[0].Field<string>("fatora")) != 0)
                    {

                        cmd.CommandType = CommandType.Text;
                        OleDbDataAdapter daa = new OleDbDataAdapter(cmd);

                        //group

                        //to get payed cash or not
                        if (double.Parse(mainDataTableDataReport.Rows[0].Field<string>("cashh")) != 0)
                        {
                            //didn't pay

                            con.Open();

                            double priceNNum = 0;
                            for (int z = 0; z < reportsDataGridView.Rows.Count; z++)
                            {
                                if (reportsDataGridView.Rows[z].Cells[4].Value != null)
                                {
                                    priceNNum = priceNNum + (double.Parse(reportsDataGridView.Rows[z].Cells[4].Value.ToString()) * double.Parse(reportsDataGridView.Rows[z].Cells[2].Value.ToString()));
                                }
                            }



                            cmd.CommandText = "update groupp set price= price + '" + priceNNum + "' where id =" + mainDataTableDataReport.Rows[0].Field<string>("groupp");
                            cmd.ExecuteNonQuery();


                            con.Close();


                        }
                        else
                        {
                            //payed cash


                        }

                    }
                    else
                    {
                        //afrad


                    }



                    con.Open();

                    cmd.CommandText = "delete from sell where fatora ='" + textBox26.Text + "'";
                    cmd.ExecuteNonQuery();


                    con.Close();


                    reportsDataGridView.Rows.Clear();


                    MessageBox.Show("تم بنجاح");



                }
                else
                {

                    MessageBox.Show("برجاء كتابة رقم الفاتورة او اختيارها من الجدول");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error B6 " + ex.Message);
                Console.WriteLine(ex.Message);
            }


        }
    }



}








