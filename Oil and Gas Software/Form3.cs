using MetroFramework.Forms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Oil_and_Gas_Software
{
    public partial class Form3 : MetroForm
    {
        SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        SqlDataAdapter adapt2;
        SqlDataAdapter adapt3;
        DataTable dt1 = new DataTable();
        DataTable dt11 = new DataTable();

        DataTable dt2 = new DataTable();
        DataTable dt22 = new DataTable();
        DataTable dt3 = new DataTable();


        //ID variable used in Updating and Deleting Record  
        int ID = 0;
        public Form3()
        {

            InitializeComponent();
            int ID = 0;


        }
        //Display Data in DataGridView  1
        private void DisplayData()
        {
            dt1.Rows.Clear(); 
            con.Open();
            adapt = new SqlDataAdapter("select distinct catid 'Category ID', catname 'Category_Name'  from CATEGORY", con);
            adapt.Fill(dt1);
            dataGridView1.DataSource = dt1;
            con.Close();
          
        }
        //Display Data in DataGridView  2       
        private void DisplayData2()
        {
            dt2.Rows.Clear();

            con.Open();
            adapt2 = new SqlDataAdapter("select Subid 'ID' ,Subname 'Subcategory',catname 'Category'  from SUBCATEGORY,category where category.catid = subcategory.catid   order by subid", con);
            adapt2.Fill(dt2);
            dataGridView2.DataSource = dt2;
            con.Close();
        }
        //Display Data in DataGridView 3
 
        private void DisplayData3()
        {
            dt3.Rows.Clear();
            con.Open();
            adapt3 = new SqlDataAdapter("select matid 'ID' ,matname 'Material',subname 'Subcategory'  from SUBCATEGORY,materials where materials.subid = subcategory.subid  order by matid ", con);
            adapt3.Fill(dt3);
            dataGridView3.DataSource = dt3;
            con.Close();
        }
        //Clear Data  
        private void ClearData()
        {
            txt_Name.Text = "";
            txt_Name1.Text = "";
            txt_Name2.Text = "";
            txt_Name3.Text = "";
            ID = 0;
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
            dataGridView2.Visible = false;
            dataGridView3.Visible = true;
            label3.Visible = true;
            label3.Text = "* Please choose Subacategory for all Materials operations";
            label2.Enabled = true;
            label1.Text = "Material Name";
            label2.Text = "Subcategory";
            comboBox1.Enabled = false;
            comboBox1.Visible = false;
            comboBox2.Visible = true;
            comboBox2.Enabled = true;
            metroButton1.Enabled = false;
            metroButton3.Enabled = false;
            metroButton2.Enabled = true;
            txt_Name3.Visible = true;
            txt_Name2.Visible = false;
            txt_Name1.Visible = false;
            txt_Name.Enabled = false;
            refreshdatlSubategory();
            DisplayData3();
        }
        public void refreshdataCategory()
        {
            dt11.Rows.Clear();
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select catid ,catname   from category  order by catname ", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(dt11);
            dr = dt11.NewRow();
            if (dt11 != null)
            {
                dr.ItemArray = new object[] { 0, "--Select Category--" };
                dt11.Rows.InsertAt(dr, 0);
                comboBox1.ValueMember = "catid";
                comboBox1.DisplayMember = "catname";
                /*clear white space in datatable*/
                dt11.AsEnumerable().ToList().ForEach(row =>
                {
                    var cellList = row.ItemArray.ToList();
                    row.ItemArray = cellList.Select(x => x.ToString().Trim()).ToArray();
                });
                /*clear white space in datatable*/
                comboBox1.DataSource = dt11;
                comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
                comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;

                con.Close();
            }
            else
            {
                // MessageBox.Show("Please choose a folder to import 'Materials'  ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void refreshdatlSubategory()
        {
            dt22.Rows.Clear();
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select Subid,Subname from SUBCATEGORY order by Subname", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(dt22);
            dr = dt22.NewRow();
            if (dt22 != null)
            {
                dr.ItemArray = new object[] { 0, "--Select Subcategory--" };
                dt22.Rows.InsertAt(dr, 0);
                comboBox2.ValueMember = "Subid";
                comboBox2.DisplayMember = "Subname";
                /*clear white space in datatable*/
                dt22.AsEnumerable().ToList().ForEach(row =>
                {
                    var cellList = row.ItemArray.ToList();
                    row.ItemArray = cellList.Select(x => x.ToString().Trim()).ToArray();
                });
                /*clear white space in datatable*/
                comboBox2.DataSource = dt22;
                comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
                comboBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;


                con.Close();
            }
            else
            {
                MessageBox.Show("Please choose a folder to import 'Materials'  ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            this.Hide();
            frm1.Show();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Form2 frm2 = new Form2();
            this.Hide();
            frm2.Show();
        }

        private void metroCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
       
        }
        //Insert Record  
        private void metroButton1_Click(object sender, EventArgs e)
        {
            ///category 
            if (metroRadioButton3.Checked == true)
            {

                if (txt_Name.Text != "")
                {
                    cmd = new SqlCommand("insert into CATEGORY(CatName) values(@name) ", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@name", txt_Name.Text);
                    if (DialogResult.Yes == MessageBox.Show("Do You Want Insert this record ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        // do what u want
                        cmd.ExecuteNonQuery();
                        con.Close();
                        MessageBox.Show("Record Inserted Successfully");
                        DisplayData();
                        ClearData();
                    }
                    else
                    {
                        con.Close();
                    }

                }
                else
                {
                    MessageBox.Show("Please Provide Details!");
                }

            }

            ///subategory 
            if (metroRadioButton1.Checked == true)
            {
                if (txt_Name.Text != "" && comboBox1.SelectedIndex != -1)
                {
                    cmd = new SqlCommand("insert into subcategory(subname,catid) values(@name,@catid)", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@name", txt_Name.Text);
                    cmd.Parameters.AddWithValue("@catid", comboBox1.SelectedValue);
                    if (DialogResult.Yes == MessageBox.Show("Do You Want Insert this record ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        // do what u want
                        cmd.ExecuteNonQuery();
                        con.Close();
                        MessageBox.Show("Record Inserted Successfully");
                        DisplayData2();
                        ClearData();
                    }
                    else
                    {
                        con.Close();
                    }

                }
                else
                {
                    MessageBox.Show("Please Provide Details!");
                }


            }

        }
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // check null cells
            foreach (DataGridViewRow rw in this.dataGridView1.Rows)
            {
                for (int i = 0; i < rw.Cells.Count; i++)
                {
                    if (rw.Cells[i].Value == null || rw.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[i].Value.ToString()))
                    {
                        //   MessageBox.Show("ogg");       
                    }
                    else
                    {

                        ID = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                        txt_Name1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    }
                }
            }





        }
        //Update Record  
        private void metroButton2_Click(object sender, EventArgs e)
        {
            ///category 
            if (metroRadioButton3.Checked == true)
            {

                if (txt_Name1.Text != "")
                {
                    cmd = new SqlCommand("update CATEGORY set catname=@name where catid=@id", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    cmd.Parameters.AddWithValue("@name", txt_Name.Text);
                    if (DialogResult.Yes == MessageBox.Show("Do You Want Update ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        // do what u want
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Updated Successfully");
                        con.Close();
                        DisplayData();
                        ClearData();
                    }
                    else
                    {
                        con.Close();
                    }

                }
                else
                {
                    MessageBox.Show("Please Select Record to Update");
                }

            }


            ///subcategory 
            if (metroRadioButton1.Checked == true)
            {

                if (txt_Name2.Text != "")
                {
                    cmd = new SqlCommand("update subcategory set subname=@name , catid=@catid where subid=@id", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    cmd.Parameters.AddWithValue("@name", txt_Name.Text);
                    cmd.Parameters.AddWithValue("@catid", comboBox1.SelectedValue);

                    if (DialogResult.Yes == MessageBox.Show("Do You Want Update ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        // do what u want
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Updated Successfully");
                        con.Close();
                        DisplayData2();
                        ClearData();
                    }
                    else
                    {
                        con.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Please Select Record to Update");
                }

            }


            ///materials 
            if (metroRadioButton2.Checked == true)
            {
                if (txt_Name3.Text != "")
                {
                    cmd = new SqlCommand("update materials set subid=@subid  where matid=@id", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    cmd.Parameters.AddWithValue("@subid", comboBox2.SelectedValue);

                    if (DialogResult.Yes == MessageBox.Show("Do You Want Update ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        // do what u want
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Updated Successfully");
                        con.Close();
                        DisplayData3();
                        ClearData();
                    }
                    else
                    {
                        con.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Please Select Record to Update");

                }
            }    
        }
            //Delete Record  
       private void metroButton3_Click(object sender, EventArgs e)
            {
            ///category 
            if (metroRadioButton3.Checked == true)
            {

                if (ID != 0)
                {
                    cmd = new SqlCommand("delete CATEGORY where catid=@id", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    if (DialogResult.Yes == MessageBox.Show("Do You Want Delete ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        // do what u want
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Deleted Successfully");
                        con.Close();
                        DisplayData();
                        ClearData();
                    }
                    else
                    {
                        con.Close();

                    }

                }
                else
                {
                    MessageBox.Show("Please Select Record to Delete");
                }

            }

            ///subcategory
            if (metroRadioButton1.Checked == true)
                {
                if (ID != 0)
                {
                    cmd = new SqlCommand("delete subcategory where subid=@id", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    if (DialogResult.Yes == MessageBox.Show("Do You Want Delete ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        // do what u want
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Deleted Successfully");
                        con.Close();
                        DisplayData2();
                        ClearData();
                    }
                    else
                    {
                        con.Close();

                    }

                }
                else
                {
                    MessageBox.Show("Please Select Record to Delete");
                }
            }

            ///materials
            if (metroRadioButton1.Checked == true)
            {
                if (ID != 0)
                {
                    cmd = new SqlCommand("delete materials where matid=@id", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    if (DialogResult.Yes == MessageBox.Show("Do You Want Delete ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        // do what u want
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Deleted Successfully");
                        con.Close();
                        DisplayData3();
                        ClearData();
                    }
                    else
                    {
                        con.Close();

                    }

                }
                else
                {
                    MessageBox.Show("Please Select Record to Delete");
                }
            }

        }

        private void dataGridView2_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           // check null cells
            foreach (DataGridViewRow rw in this.dataGridView2.Rows)
            {
                for (int i = 0; i < rw.Cells.Count; i++)
                {
                    if (rw.Cells[i].Value == null || rw.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[i].Value.ToString()))
                    {
                        //   MessageBox.Show("ogg");       
                    }
                    else
                    {

                            ID = Convert.ToInt32(dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString());
                            txt_Name2.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                           // comboBox1.DisplayMember = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();


        }
    }
}
        }

        private void metroCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void dataGridView3_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // check null cells
            foreach (DataGridViewRow rw in this.dataGridView3.Rows)
            {
                for (int i = 0; i < rw.Cells.Count; i++)
                {
                    if (rw.Cells[i].Value == null || rw.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[i].Value.ToString()))
                    {
                        //   MessageBox.Show("ogg");       
                    }
                    else
                    {

                        ID = Convert.ToInt32(dataGridView3.Rows[e.RowIndex].Cells[0].Value.ToString());
                        txt_Name3.Text = dataGridView3.Rows[e.RowIndex].Cells[1].Value.ToString();
                      //  comboBox1.DisplayMember = dataGridView3.Rows[e.RowIndex].Cells[2].Value.ToString();


                    }
                }
            }
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DataRow dr;


            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select  CATEGORY.catid,CATEGORY.catname from SUBCATEGORY,CATEGORY  " +
                "  where  CATEGORY.CatID  = SUBCATEGORY.Catid  and SUBCATEGORY.SubID= @C1 order by  CATEGORY.catid";



            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
            cmd.Parameters["@C1"].Value = comboBox2.SelectedValue;

          
            //Creating Sql Data Adapter
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter Da = new SqlDataAdapter(cmd);
            Da.Fill(dt);
            dr = dt.NewRow();

            if (dt != null && dt.Rows.Count >= 0)
            {
                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);


                comboBox3.ValueMember = "catid";
                comboBox3.DisplayMember = "catname";


                comboBox3.DataSource = dt;
            }
            conn.Close();

        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txt_Name_TextChanged(object sender, EventArgs e)
        {
            DataView dv = dt1.DefaultView;

            dv.RowFilter = " Category_Name LIKE '" + txt_Name1.Text + "%'";
            dataGridView1.DataSource = dv;

        }

        private void metroRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (metroRadioButton2.Checked == true)
            {
                dataGridView1.Visible = false;
                dataGridView2.Visible = false;
                dataGridView3.Visible = true;
                label3.Visible = true;
                label3.Text = "* Please choose Subacategory for all Materials operations";
                label2.Enabled = true;
                label1.Text = "Material Name";
                label2.Text = "Subcategory";
                comboBox1.Enabled = false;
                comboBox1.Visible = false;
                comboBox2.Visible = true;
                comboBox2.Enabled = true;
                metroButton1.Enabled = false;
                metroButton3.Enabled = false;
                metroButton2.Enabled = true;
                txt_Name3.Visible = true;
                txt_Name2.Visible = false;
                txt_Name1.Visible = false;
                txt_Name.Enabled = false;
                refreshdatlSubategory();
                DisplayData3();

            }
            else
            {
                metroButton1.Enabled = true;
                metroButton3.Enabled = true;
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                dataGridView3.Visible = false;
                comboBox1.Enabled = false;
                comboBox1.Visible = true;
                comboBox2.Visible = false;
                comboBox2.Enabled = false;
                label3.Visible = false;
                label2.Text = "Category";
                label1.Text = "Category Name";
                

                DisplayData();
                label2.Enabled = false;

            }
        }

        private void metroRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (metroRadioButton1.Checked == true)
            {
                dataGridView1.Visible = false;
                dataGridView2.Visible = true;
                label3.Visible = true;
                label2.Enabled = true;
                label1.Text = "Subcategory Name";
                label3.Text = "* Please choose Category for all Subcategory operations";

                comboBox1.Enabled = true;
                comboBox1.Visible = true;
                comboBox2.Visible = false;
                metroButton3.Enabled = false;
                metroButton2.Enabled = true;
                metroButton1.Enabled = true;
                txt_Name3.Visible = false;
                txt_Name2.Visible = true;
                txt_Name1.Visible = false;

                refreshdataCategory();
                DisplayData2();

            }
            else
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                dataGridView3.Visible = false;
                comboBox2.Visible = false;
                label3.Visible = false;
                label1.Text = "Category Name";
                DisplayData();
                label2.Enabled = false;
                comboBox1.Enabled = false;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            comboBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;

        }

        private void metroRadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (metroRadioButton3.Checked == true)
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                dataGridView3.Visible = false;
                comboBox2.Visible = false;
                label3.Visible = false;
                label1.Text = "Category Name";
                DisplayData();
                label2.Enabled = false;
                comboBox1.Enabled = false;

                metroButton3.Enabled = true;
                metroButton2.Enabled = true;
                metroButton1.Enabled = true;

                txt_Name3.Visible = false;
                txt_Name2.Visible = false;
                txt_Name1.Visible = true;
                txt_Name.Enabled = true;

            }
        }

        private void txt_Name2_TextChanged(object sender, EventArgs e)
        {
            DataView dv2 = dt2.DefaultView;

            dv2.RowFilter = " Subcategory LIKE '" + txt_Name2.Text + "%'";
            dataGridView2.DataSource = dv2;
        }

        private void txt_Name3_TextChanged(object sender, EventArgs e)
        {
            DataView dv3 = dt3.DefaultView;

            dv3.RowFilter = " Material LIKE '" + txt_Name3.Text + "%'";
            dataGridView3.DataSource = dv3;
        }

        private void comboBox1_MouseEnter(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        private void comboBox2_MouseEnter(object sender, EventArgs e)
        {
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            comboBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_Enter(object sender, EventArgs e)
        {
            DataRow dr;


            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select  CATEGORY.catid,CATEGORY.catname from SUBCATEGORY,CATEGORY  " +
                "  where  CATEGORY.CatID  = SUBCATEGORY.Catid  and SUBCATEGORY.SubID= @C1 order by  CATEGORY.catid";



            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
            cmd.Parameters["@C1"].Value = comboBox2.SelectedValue;


            //Creating Sql Data Adapter
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter Da = new SqlDataAdapter(cmd);
            Da.Fill(dt);
            dr = dt.NewRow();

            if (dt != null && dt.Rows.Count >= 0)
            {
                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);


                comboBox3.ValueMember = "catid";
                comboBox3.DisplayMember = "catname";


                comboBox3.DataSource = dt;
            }
            conn.Close();
        }

        private void comboBox2_MouseClick(object sender, MouseEventArgs e)
        {
          
        }

        private void comboBox2_DragEnter(object sender, DragEventArgs e)
        {
            DataRow dr;


            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select  CATEGORY.catid,CATEGORY.catname from SUBCATEGORY,CATEGORY  " +
                "  where  CATEGORY.CatID  = SUBCATEGORY.Catid  and SUBCATEGORY.SubID= @C1 order by  CATEGORY.catid";



            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
            cmd.Parameters["@C1"].Value = comboBox2.SelectedValue;


            //Creating Sql Data Adapter
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter Da = new SqlDataAdapter(cmd);
            Da.Fill(dt);
            dr = dt.NewRow();

            if (dt != null && dt.Rows.Count >= 0)
            {
                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);


                comboBox3.ValueMember = "catid";
                comboBox3.DisplayMember = "catname";


                comboBox3.DataSource = dt;
            }
            conn.Close();
        }

        private void comboBox2_DisplayMemberChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_TabIndexChanged(object sender, EventArgs e)
        {
            DataRow dr;


            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select  CATEGORY.catid,CATEGORY.catname from SUBCATEGORY,CATEGORY  " +
                "  where  CATEGORY.CatID  = SUBCATEGORY.Catid  and SUBCATEGORY.SubID= @C1 order by  CATEGORY.catid";



            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
            cmd.Parameters["@C1"].Value = comboBox2.SelectedValue;


            //Creating Sql Data Adapter
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter Da = new SqlDataAdapter(cmd);
            Da.Fill(dt);
            dr = dt.NewRow();

            if (dt != null && dt.Rows.Count >= 0)
            {
                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);


                comboBox3.ValueMember = "catid";
                comboBox3.DisplayMember = "catname";


                comboBox3.DataSource = dt;
            }
            conn.Close();
        }
    }

    }

