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
        SqlDataReader reader;
      
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
                dr.ItemArray = new object[] { 0, "All" };
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
                SqlCommand cmd = new SqlCommand("select distinct CATEGORY.CatID,CATEGORY.CatName from  REPORTS,CATEGORY ,SUBCATEGORY,MATERIALS,MUD_TRATMENT " +
                    "where reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid and CATEGORY.CatID = SUBCATEGORY.Catid and SUBCATEGORY.Subid = " +
                    " MATERIALS.SubID and REPORTS.Date >= @C1 and REPORTS.Date <= @C2 ", con);
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

                    dr.ItemArray = new object[] { 0, "All" };
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
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select matid,matname from materials order by matname", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {
                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);
                MatComboBox.ValueMember = "matid";
                MatComboBox.DisplayMember = "matname";
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
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();

            SqlCommand cmd = new SqlCommand("select WELLID,Wellname from wells where Wellname !='' order by Wellname ", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {

                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);
                WellComboBox.ValueMember = "Wellid";
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

        }

        private void CatComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {


            DataRow dr;
            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct SUBCATEGORY.Subid,Subname from SUBCATEGORY,CATEGORY,reports,materials,MUD_TRATMENT where reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid  and CATEGORY.CatID = SUBCATEGORY.Catid and " +
                " SUBCATEGORY.subid = materials.subid " +
                " and CATEGORY.Catid= @C1  and " +
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
            dr = dt.NewRow();


            if (dt != null && dt.Rows.Count >= 0)
            {
                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);

                SubCatComboBox.ValueMember = "Subid";
                SubCatComboBox.DisplayMember = "Subname";


                SubCatComboBox.DataSource = dt;
                refreshdataMaterial();
                refreshdataRIGS();
                refreshdataWell();



            }

            conn.Close();



        }

        private void SubCatComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DataRow dr;


            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct MATERIALS.MATID,MATName from MATERIALS,SUBCATEGORY,CATEGORY,reports,MUD_TRATMENT " +
                "where CATEGORY.CatID  = SUBCATEGORY.Catid and SUBCATEGORY.subid = materials.subid and reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid " +
                " and SUBCATEGORY.SubID= @C1  and  reports.date >= @C2  and reports.date <= @C3   order by MATName";



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
            dr = dt.NewRow();

            if (dt != null && dt.Rows.Count >= 0)
            {
                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);


                MatComboBox.ValueMember = "MATID";
                MatComboBox.DisplayMember = "MATName";
               

                MatComboBox.DataSource = dt;
            }
            conn.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {

            string SQuery = "select  rigs.Rigname  'Rig',WELLS.Wellname 'Well No'," +
                " CATEGORY.CatName 'Category',SUBCATEGORY.Subname 'Subcategory', MATERIALS.MATName 'Materials', MUD_TRATMENT.QTY'QTY'," +
                " MUD_TRATMENT.PackingQTY'PQTY', QTY* PackingQTY 'Total/Unit',MUD_TRATMENT.UnitName 'Unit',REPORTS.DEPTH ,LAST24,DAYSSINCE 'Days since' , reports.Date 'Date' " +
                " from " +
                "RIGS, WELLS, REPORTS, MUD_TRATMENT, MATERIALS, CATEGORY, SUBCATEGORY where " +
                "REPORTS.RIGID = rigs.RIGID and " +
                " reports.WELLID = WELLS.WELLID  and " +
                "MUD_TRATMENT.MATID = MATERIALS.MATID and" +
                " MUD_TRATMENT.REPORTID = REPORTS.REPORTID and" +
                " CATEGORY.CatID = SUBCATEGORY.Catid and" +
                " SUBCATEGORY.subid = MATERIALS.SubID and " +
                "reports.date >= @C2  and  reports.date <= @C3 ";

            /*Query builder **/

            if ((int)CatComboBox.SelectedValue != 0)
            {
                SQuery = SQuery + " and CATEGORY.catid = " + CatComboBox.SelectedValue;
                DataTable dt2 = new DataTable();
                // dt2.Rows.Clear();
                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                {

                    using (SqlCommand cmd = new SqlCommand(SQuery, con))
                    {
                        cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                        cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

                        cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                        cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
                        con.Open();

                        using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                        {
                            using (dt2)
                            {
                                ada.Fill(dt2);
                                dataGridView1.DataSource = dt2;
                            }
                            con.Close();

                        }

                        if ((int)SubCatComboBox.SelectedValue != 0)
                        {
                              dt2.Rows.Clear();

                            SQuery = SQuery + " and SUBCATEGORY.subid = " + SubCatComboBox.SelectedValue;

                            using (SqlCommand cmd1 = new SqlCommand(SQuery,con))
                            {
                                cmd1.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                                cmd1.Parameters["@C2"].Value = dateTimePicker1.Value;

                                cmd1.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                                cmd1.Parameters["@C3"].Value = dateTimePicker2.Value;
                                con.Open();

                                using (SqlDataAdapter ada = new SqlDataAdapter(cmd1))
                                {
                                    using (dt2)
                                    {

                                        ada.Fill(dt2);
                                        dataGridView1.DataSource = dt2;
                                    }

                                    con.Close();

                                }
                                if ((int)MatComboBox.SelectedValue != 0)
                                {
                                      dt2.Rows.Clear();

                                    SQuery = SQuery + " and materials.matid =  " + MatComboBox.SelectedValue;
                                    using (SqlCommand cmd2 = new SqlCommand(SQuery,con))
                                    {
                                        cmd2.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                                        cmd2.Parameters["@C2"].Value = dateTimePicker1.Value;

                                        cmd2.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                                        cmd2.Parameters["@C3"].Value = dateTimePicker2.Value;
                                        con.Open();
                                        
                                        using (SqlDataAdapter ada = new SqlDataAdapter(cmd2))
                                        {
                                            using (dt2)
                                            {

                                                ada.Fill(dt2);
                                                dataGridView1.DataSource = dt2;
                                            }

                                            con.Close();


                                        }

                                    }

                                }

                            }
                        }
                    }


                }
            }

           else  if ((int)RigComboBox.SelectedValue != 0)
            {
                DataTable dt2 = new DataTable();
                dt2.Rows.Clear();
                SQuery = SQuery + " and rigs.rigid = " + RigComboBox.SelectedValue;
                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                {

                    using (SqlCommand cmd = new SqlCommand(SQuery, con))
                    {
                        cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                        cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

                        cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                        cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
                        con.Open();
                        MessageBox.Show(SQuery);

                        using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                        {
                            using (dt2)
                            {
                                ada.Fill(dt2);
                                dataGridView1.DataSource = dt2;
                            }
                            con.Close();
                        }

                    }


                }


            }

            else if ((int)WellComboBox.SelectedValue != 0)
            {

                DataTable dt2 = new DataTable();
                dt2.Rows.Clear();

                SQuery = SQuery + " and wells.wellid = " + WellComboBox.SelectedValue;
                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                {

                    using (SqlCommand cmd = new SqlCommand(SQuery, con))
                    {
                        cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                        cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

                        cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                        cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
                        con.Open();
                        MessageBox.Show(SQuery);

                        using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                        {
                            using (dt2)
                            {
                                ada.Fill(dt2);
                                dataGridView1.DataSource = dt2;
                            }
                            con.Close();
                        }

                    }


                }



            }

            else
            {
                DataTable dt3 = new DataTable();
                dt3.Rows.Clear();
                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                {
                    using (SqlCommand cmd = new SqlCommand(SQuery, con))
                    {

                        cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                        cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

                        cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                        cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
                        con.Open();
                        MessageBox.Show(SQuery);

                        using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                        {

                            using (dt3)
                            {
                                ada.Fill(dt3);
                                dataGridView1.DataSource = dt3;
                            }
                       //     dt3.Rows.Clear();
                            con.Close();
                        }
                    }

                }
            }
            MessageBox.Show(SQuery);


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

        private void MatComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DataRow dr;
            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");

            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct rigs.rigid,rigname from MATERIALS,SUBCATEGORY,CATEGORY,reports,MUD_TRATMENT,rigs " +
                "where reports.rigid = rigs.rigid and CATEGORY.CatID  = SUBCATEGORY.Catid and SUBCATEGORY.subid = materials.subid and reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid " +
                " and MATERIALS.matid= @C1  and  reports.date >= @C2  and reports.date <= @C3   order by rigname";



            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
            cmd.Parameters["@C1"].Value = MatComboBox.SelectedValue;

            cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
            cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

            cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
            cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
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

                RigComboBox.ValueMember = "rigid";
                RigComboBox.DisplayMember = "rigname";

                RigComboBox.DataSource = dt;
            }
            conn.Close();
        }

        private void RigComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DataRow dr;
            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");

            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct wells.wellid,wellname from wells,MATERIALS,SUBCATEGORY,CATEGORY,reports,MUD_TRATMENT,rigs " +
                "where wells.wellid= reports.wellid and reports.rigid = rigs.rigid and CATEGORY.CatID  = SUBCATEGORY.Catid and SUBCATEGORY.subid = materials.subid and reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid " +
                " and rigs.rigid= @C1  and  reports.date >= @C2  and reports.date <= @C3   order by wellname";



            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
            cmd.Parameters["@C1"].Value = RigComboBox.SelectedValue;

            cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
            cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

            cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
            cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
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

                WellComboBox.ValueMember = "wellid";
                WellComboBox.DisplayMember = "wellname";

                WellComboBox.DataSource = dt;
            }
            conn.Close();
        }

        private void MatComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void MatComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //if (radioButton1.Checked == true)
            //{
            //    groupBox3.Enabled = true;
            //}
            //else 
            //{
            //    groupBox3.Enabled = false;


            //}
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            //if (radioButton2.Checked == true)
            //{
            //    groupBox3.Enabled = true;
            //}
            //else
            //{
            //    groupBox3.Enabled = false;


            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            reset();
         
        }
        public void reset()
        {
            if (CatComboBox.SelectedIndex != 0)
            {
                CatComboBox.SelectedIndex = 0;
            }
            else
            { }

            if (SubCatComboBox.SelectedIndex != 0)
            {
                SubCatComboBox.SelectedIndex = 0;
            }
            else { }

            if (MatComboBox.SelectedIndex != 0)
            {
                MatComboBox.SelectedIndex = 0;
            }
            else { }
            if (RigComboBox.SelectedIndex != 0)
            {
                RigComboBox.SelectedIndex = 0;
            }
            else { }
            if (WellComboBox.SelectedIndex != 0)
            {
                WellComboBox.SelectedIndex = 0;
            }
            else { }

            dataGridView1.DataSource = null;
            RowsNuumlblNEW.Text = string.Empty;
            subtot.Text = string.Empty;
            SubTONEW.Text = string.Empty;

        }
    }
}


