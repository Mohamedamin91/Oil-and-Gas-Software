using Aspose.Cells;
using MetroFramework.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Oil_and_Gas_Software
{
    public partial class Form2 : MetroForm
    {
        DataTable dt = new DataTable();
        SqlDataReader reader;
        DateTime Datefrom;
        DateTime DateTo;
        
        public Form2()
        {
            InitializeComponent();
        }
        public void BindTotal()
        {
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
             SqlCommand cmd1 = new SqlCommand(" select  COUNT(distinct matname)'material' from materials ", con);
              SqlCommand cmd2 = new SqlCommand(" select  count (distinct Rigname)  'RigNo' from Rigs  ", con);
            SqlCommand cmd3 = new SqlCommand(" select  count (distinct wellname) 'Well No' from wells ", con);
            //SqlCommand cmd4 = new SqlCommand(" select  COUNT(distinct Rigname)'Well No (type) Contain data' from rigs where Contain = 1 ", con);
            //SqlCommand cmd5 = new SqlCommand(" select  COUNT(distinct Rigname)'Well No (type) Non Contain data' from rigs where Contain = 0 ", con);
            SqlCommand cmd6 = new SqlCommand(" select  COUNT(distinct CatID)'Category' from Category", con);
            SqlCommand cmd7 = new SqlCommand(" select  COUNT( distinct Subid)'SubCategory' from SUBCATEGORY", con);
            //SqlCommand cmd8 = new SqlCommand(" select  COUNT(distinct KeywordName)'Material Without' from materials where CatID = 0 ", con);
            SqlCommand cmd9 = new SqlCommand(" select COUNT ( distinct reportid) from reports ", con);


            cmd1.ExecuteNonQuery();
            cmd2.ExecuteNonQuery();
            cmd3.ExecuteNonQuery();
            ////  cmd4.ExecuteNonQuery();
            ////cmd5.ExecuteNonQuery();
            cmd6.ExecuteNonQuery();
            cmd7.ExecuteNonQuery();
            //cmd8.ExecuteNonQuery();
            cmd9.ExecuteNonQuery();
            reader = cmd1.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    label13.Text = reader[0].ToString();
                }
            }
            con.Close();
            con.Open();
            reader = cmd2.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    label15.Text = reader[0].ToString();
                }
            }
            con.Close();
            con.Open();
            reader = cmd3.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    label17.Text = reader[0].ToString();
                }
            }
            con.Close();
            //con.Open();
            //reader = cmd4.ExecuteReader();
            //if (reader.HasRows)
            //{
            //    while (reader.Read())
            //    {
            //        label7.Text = reader[0].ToString();
            //    }
            //}
            //con.Close();
            //con.Open();
            //reader = cmd5.ExecuteReader();
            //if (reader.HasRows)
            //{
            //    while (reader.Read())
            //    {
            //        label10.Text = reader[0].ToString();
            //    }
            //}
            //con.Close();
            //con.Open();
            con.Close();
            con.Open();
            reader = cmd6.ExecuteReader();
            if (reader.HasRows)
            {
               while (reader.Read())
                {
                    label9.Text = reader[0].ToString();
                }
            }
            con.Close();
            con.Open();
            con.Close();
            con.Open();
            reader = cmd7.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    label11.Text = reader[0].ToString();
                }
            }
            //con.Close();
            //con.Open();
            //con.Close();
            //con.Open();
            //reader = cmd8.ExecuteReader();
            //if (reader.HasRows)
            //{
            //    while (reader.Read())
            //    {
            //        label16.Text = reader[0].ToString();
            //    }
            //}
            con.Close();
            con.Open();
            reader = cmd9.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    label8.Text = reader[0].ToString();
                }
            }
            con.Close();

        }
        public void refreshdataMaterialWithout()
        {
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT KeywordName FROM [OILREPORT].[dbo].KEYWORDS where CatID= 0", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);

            reader = cmd.ExecuteReader();
            //    listBox1.Visible = false;
            ////metroLabel7.Visible = false;
            while (reader.Read())
            {
                //   listBox1.Items.Add(reader["KeywordName"]);
            }
            con.Close();
        }
        public void refreshdataRIGS()
        {
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select distinct RigID, Rigname from Rigs where Rigname !=''    ", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {
                dr.ItemArray = new object[] { 0, "--Select Rigs--" };
                dt.Rows.InsertAt(dr, 0);
                RigComboBox.ValueMember = "RigID";
                RigComboBox.DisplayMember = "Rigname";
                RigComboBox.DataSource = dt;

                con.Close();
            }
            else
            {
                MessageBox.Show("Please choose a folder to import 'Materials'  ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void refreshdataMaterialSubategory()
        {
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select Subid,Subname from Subcatogory order by Subname", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {
                dr.ItemArray = new object[] { 0, "--Select Subcatogory--" };
                dt.Rows.InsertAt(dr, 0);
                SubCatComboBox.ValueMember = "Subid";
                SubCatComboBox.DisplayMember = "Subname";
                /*clear white space in datatable*/
                dt.AsEnumerable().ToList().ForEach(row =>
                {
                    var cellList = row.ItemArray.ToList();
                    row.ItemArray = cellList.Select(x => x.ToString().Trim()).ToArray();
                });
                /*clear white space in datatable*/
                SubCatComboBox.DataSource = dt;

                con.Close();
            }
            else
            {
                MessageBox.Show("Please choose a folder to import 'Materials'  ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void refreshdataMaterialCategory()
        {
            
            if (dateTimePicker1.Value != null || dateTimePicker2.Value != null)
            {
                DataRow dr;
                SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
                con.Open();
                SqlCommand cmd = new SqlCommand("select distinct CATEGORY.CatID,CATEGORY.CatName from  REPORTS,CATEGORY ,SUBCATEGORY,MATERIALS " +
                    "where CATEGORY.CatID = SUBCATEGORY.Catid and SUBCATEGORY.Subid = MATERIALS.SubID and REPORTS.Date >= @C1 and REPORTS.Date <= @C2 ", con);
                cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Date));
                cmd.Parameters["@C1"].Value = dateTimePicker1.Value;

                cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                cmd.Parameters["@C2"].Value = dateTimePicker2.Value;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                dr = dt.NewRow();
                if (dt != null)
                {
                    dr.ItemArray = new object[] { 0, "--Select Category--" };
                    dt.Rows.InsertAt(dr, 0);
                    this.CatComboBox.ValueMember = "CatID";
                    this.CatComboBox.DisplayMember = "CatName";
                    this.CatComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    this.CatComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
                    this.CatComboBox.DropDownStyle = ComboBoxStyle.DropDown;

                    /*clear white space in datatable*/
                    dt.AsEnumerable().ToList().ForEach(row =>
                    {
                        var cellList = row.ItemArray.ToList();
                        row.ItemArray = cellList.Select(x => x.ToString().Trim()).ToArray();
                    });
                    /*clear white space in datatable*/
                    CatComboBox.DataSource = dt;
                    con.Close();
                }
                else
                {
                //    MessageBox.Show("Please choose a folder to import 'Materials'  ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Please choose a Date   ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);


            }

        }
        public void refreshdataMaterial()
        {
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT1;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select KeywordID,KeywordName from KEYWORDS order by KeywordName", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {
                dr.ItemArray = new object[] { 0, "--Select Material--" };
                dt.Rows.InsertAt(dr, 0);
                MatComboBox.ValueMember = "KeywordID";
                MatComboBox.DisplayMember = "KeywordName";
                MatComboBox.DataSource = dt;
                con.Close();
            }
            else
            {
                MessageBox.Show("Please choose a folder to import 'Materials'  ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void refreshdataWell()
        {
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();

            SqlCommand cmd = new SqlCommand("select WELLID,Wellname from wells where Wellname !='' order by Wellname ", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {

                dr.ItemArray = new object[] { 0, "Select Well" };
                dt.Rows.InsertAt(dr, 0);
                WellComboBox.ValueMember = "RigID";
                WellComboBox.DisplayMember = "Wellname";
                /*clear white space in datatable*/
                dt.AsEnumerable().ToList().ForEach(row =>
                {
                    var cellList = row.ItemArray.ToList();
                    row.ItemArray = cellList.Select(x => x.ToString().Trim()).ToArray();
                });
                /*clear white space in datatable*/
                WellComboBox.DataSource = dt;
                con.Close();
            }













            else
            {
                MessageBox.Show("Please choose a folder to import 'Rig Names'  ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void BindRigsNonContainData()
        {
            DataTable dt2 = new DataTable();
            using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@"))
            {
                using (SqlCommand cmd = new SqlCommand("select distinct  [Rigname],[Depth] ,[last24],[DaysSince]from FILES,Rigs where FILES.RigID = Rigs.RigID and FILES.Contain='0'", con))
                {

                    using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                    {
                        using (dt2)
                        {
                            ada.Fill(dt2);
                            //  dataGridView2.Visible = true;
                            //   dataGridView2.DataSource = dt2;
                            //this.dataGridView2.Columns[0].Visible = false;
                            //this.dataGridView2.Columns[5].Visible = false;
                         //   RowsNuumlblNEW.Visible = true;
                            //  RowsNuumlblNEW.Text = dataGridView2.Rows.Count.ToString("N0");

                        }

                    }
                }
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
        //    refreshdataMaterialCategory();
        }

        private void CatComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {

            MatComboBox.DataSource = null;
            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
          

            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct SUBCATEGORY.Subid,Subname from SUBCATEGORY,CATEGORY,reports,materials where CATEGORY.CatID = SUBCATEGORY.Catid and " +
                " SUBCATEGORY.subid = materials.subid " +
                " and materials.Catid= @C1  and " +
                " reports.date >= @C2  and reports.date <= @C3  order by Subname";

            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
            cmd.Parameters["@C1"].Value = CatComboBox.SelectedValue;

            cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
            cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

            cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
            cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
            //Creating Sql Data Adapter
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter Da = new SqlDataAdapter(cmd);
            Da.Fill(dt);

            if (dt != null && dt.Rows.Count > 0)
            {
                SubCatComboBox.ValueMember = "Subid";
                SubCatComboBox.DisplayMember = "Subname";

                SubCatComboBox.DataSource = dt;
            }
            conn.Close();

        }

        private void SubCatComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {



            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct MATID,MATName from MATERIALS,SUBCATEGORY,CATEGORY,reports " +
                "where CATEGORY.CatID  = SUBCATEGORY.Catid and SUBCATEGORY.subid = materials.subid " +
                " and MATERIALS.SubID= @C1  and  reports.date >= @C2  and reports.date <= @C3   order by MATName";

        

            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
            cmd.Parameters["@C1"].Value = SubCatComboBox.SelectedValue;

            cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
            cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

            cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
            cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
            //Creating Sql Data Adapter
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter Da = new SqlDataAdapter(cmd);
            Da.Fill(dt);

            if (dt != null && dt.Rows.Count > 0)
            {
                MatComboBox.ValueMember = "MATID";
                MatComboBox.DisplayMember = "MATName";

                MatComboBox.DataSource = dt;
            }
            conn.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
           

            if (CatComboBox.SelectedValue != null && SubCatComboBox.SelectedValue != null && MatComboBox.SelectedValue != null)
            {


                DataTable dt2 = new DataTable();
                dt2.Rows.Clear();
                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                {

                    using (SqlCommand cmd = new SqlCommand("" +
                        "select  rigs.Rigname 'Rig', WELLS.Wellname 'Well No', CATEGORY.CatName 'Category', SUBCATEGORY.Subname 'Subcategory', MATERIALS.MATName " +
                        "'Materials', MUD_TRATMENT.QTY'QTY', MUD_TRATMENT.PackingQTY'PQTY', MUD_TRATMENT.UnitName 'Unit', REPORTS.DEPTH, LAST24, DAYSSINCE 'Days since' ,reports.Date 'Date' " +
                "from RIGS,WELLS,REPORTS,MUD_TRATMENT,MATERIALS ,CATEGORY,SUBCATEGORY " +
                "where  REPORTS.RIGID = rigs.RIGID and reports.WELLID = WELLS.WELLID  and " +
                " MUD_TRATMENT .MATID = MATERIALS.MATID and MUD_TRATMENT .REPORTID = REPORTS.REPORTID and CATEGORY.CatID" +
                " = SUBCATEGORY.Catid and SUBCATEGORY.Catid = MATERIALS .SubID  and  category.catid =@C1 and SUBCATEGORY.subid=@C2 and MUD_TRATMENT.matid=@C3 and reports.date >=@C4 and reports.date<=@C5" +
                " order by WELLS.Wellname   ", con))

                    {
                        cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
                        cmd.Parameters["@C1"].Value = CatComboBox.SelectedValue;
                        
                        cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Int));
                        cmd.Parameters["@C2"].Value = SubCatComboBox.SelectedValue;
                       
                        cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Int));
                        cmd.Parameters["@C3"].Value = MatComboBox.SelectedValue;

                        cmd.Parameters.Add(new SqlParameter("@C4", SqlDbType.Date));
                        cmd.Parameters["@C4"].Value = dateTimePicker1.Value;

                        cmd.Parameters.Add(new SqlParameter("@C5", SqlDbType.Date));
                        cmd.Parameters["@C5"].Value = dateTimePicker2.Value;



                        con.Open();
                        using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                        {
                            using (dt2)
                            {

                                ada.Fill(dt2);
                                dataGridView1.DataSource = dt2;

                                this.dataGridView1.Columns[0].Visible = true;
                                this.dataGridView1.Columns[1].Visible = true;
                                this.dataGridView1.Columns[2].Visible = true;
                                this.dataGridView1.Columns[3].Visible = true;
                                this.dataGridView1.Columns[10].Width = 350;
                                //dataGridView1.Columns[0].DisplayIndex = 1;
                                //dataGridView1.Columns[1].DisplayIndex = 0;
                                dataGridView1.Visible = true;
                                //RowsNuumlblNEW.Visible = true;
                                //subtotallbl.Visible = true;
                                //SubTONEW.Visible = true;
                                //PackingQTYlblNEW.Visible = true; PQTYNAME.Visible = false;
                                //UnitNamelblNEW.Visible = true; UNITNAMEN.Visible = true;
                                //RowsNuumlblNEW.Text = dataGridView1.Rows.Count.ToString("N0");
                                int a = 0;
                                string StringPQTY = "";
                                string StringUnitName = "";
                                foreach (DataGridViewRow r in dataGridView1.Rows)
                                {
                                    {
                                        //a += Convert.ToInt32(r.Cells[4].Value);
                                        //StringPQTY = r.Cells[3].Value.ToString();
                                        //StringUnitName = r.Cells[5].Value.ToString();
                                    }

                                }
                                //PackingQTYlblNEW.Text = StringPQTY.ToString();
                                //UnitNamelblNEW.Text = StringUnitName.ToString();
                                //SubTONEW.Text = a.ToString("N0");



                            }
                        }
                    }
                }





            }

        }

        private void Total_CheckedChanged(object sender, EventArgs e)
        {
            if (Total.Checked == true)
            {
                BindTotal();
                groupBox1.Visible = true;
            }
            else
            {
                groupBox1.Visible = false;
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            refreshdataMaterialCategory();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            refreshdataMaterialCategory();
        }
    }
}
