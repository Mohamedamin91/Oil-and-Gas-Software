using MetroFramework.Forms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;



namespace Oil_and_Gas_Software
{
    public partial class Form2 : MetroForm
    {
        SqlDataReader reader;
        DataSet ds2 = new DataSet();
        DataSet ds42 = new DataSet();
        DataSet ds23 = new DataSet();
        public Form2()
        {
            InitializeComponent();
        }
        public void BindTotal()
        {
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
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
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@");
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
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select distinct RigID, Rigname from Rigs where Rigname !='' order by rigname   ", con);
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
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select Subid,Subname from SUBCATEGORY order by Subname", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {
                dr.ItemArray = new object[] { 0, "All" };
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

            if (dateTimePicker1.Value != null && dateTimePicker2.Value != null || (int)CatComboBox.SelectedValue != 0)
            {
                    DataRow dr;

                    SqlConnection con = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
                    con.Open();
                    SqlCommand cmd = new SqlCommand("select distinct CATEGORY.CatID,CATEGORY.CatName from  REPORTS,CATEGORY ,SUBCATEGORY,MATERIALS,MUD_TRATMENT , rigs ,wells " +
                        " where rigs.rigid =reports.rigid and wells.wellid=reports.wellid and reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid and CATEGORY.CatID = SUBCATEGORY.Catid and SUBCATEGORY.Subid = " +
                        " MATERIALS.SubID and REPORTS.Date >= @C1 and REPORTS.Date <= @C2  order by CATEGORY.CatName  ", con);
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
                    this.CatComboBox.CreateControl();
                    this.CatComboBox.SelectedValue = 0;





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
                MessageBox.Show("Please choose a Date and Category   ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);


            }

        }
        public void refreshdataMaterial()
        {
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
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
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
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
            using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@"))
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
                 // no smaller than design time size
                 this.MinimumSize = new System.Drawing.Size(this.Width, this.Height);
           
           
                refreshdataMaterialSubategory();
                refreshdataMaterial();
                refreshdataRIGS();
                refreshdataWell();

            
            // no larger than screen size
            //    this.MaximumSize = new System.Drawing.Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, (int)System.Windows.SystemParameters.PrimaryScreenHeight);
           
            this.AutoSize = true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;

        }

        private void CatComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {


            DataRow dr;
            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct  SUBCATEGORY.Subid,Subname from SUBCATEGORY,CATEGORY,reports,materials,MUD_TRATMENT , rigs,wells where " +
                "  wells.wellid= reports.wellid and rigs.rigid=reports.rigid and  reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid  and CATEGORY.CatID = SUBCATEGORY.Catid and " +
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


            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");


            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct MATERIALS.MATID,MATName from MATERIALS,SUBCATEGORY,CATEGORY,reports,MUD_TRATMENT ,rigs,wells " +
                " where  wells.wellid= reports.wellid and rigs.rigid=reports.rigid and CATEGORY.CatID  = SUBCATEGORY.Catid and SUBCATEGORY.subid = materials.subid and " +
                " reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid " +
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
            if (CatComboBox.SelectedIndex == -1)

            {
                MessageBox.Show("Please choose a Date and Category   ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
              //  crystalReportViewer1.Visible = false;
                dataGridView1.Visible = true;
           

                string SQuery = "select  rigs.Rigname  'Rig',WELLS.Wellname 'Well No'," +
                    " CATEGORY.CatName 'Category',SUBCATEGORY.Subname 'Subcategory', MATERIALS.MATName 'Materials'," +
                    " MUD_TRATMENT.QTY'QTY', " +
                    " PackingQTYNewValue 'PQTY',UnitNewValue 'Unit', REPORTS.DEPTH ,LAST24,DAYSSINCE 'Days since' , reports.Date 'Date' " +
                    " from " +
                    "RIGS, WELLS, REPORTS, MUD_TRATMENT, MATERIALS, CATEGORY, SUBCATEGORY where " +
                    "REPORTS.RIGID = rigs.RIGID and " +
                    " reports.WELLID = WELLS.WELLID  and " +
                    "MUD_TRATMENT.MATID = MATERIALS.MATID and" +
                    " MUD_TRATMENT.REPORTID = REPORTS.REPORTID and" +
                    " CATEGORY.CatID = SUBCATEGORY.Catid and" +
                    " SUBCATEGORY.subid = MATERIALS.SubID and " +
                    "reports.date >= @C2  and  reports.date <= @C3   ";

                string SQuery2 = " SELECT MATERIALS.MATName 'Material', SUM(MUD_TRATMENT.QTY) as Total   , PackingQTYNewValue 'PQTY', UnitNewValue 'Unit'   " +
                    " FROM RIGS, WELLS, REPORTS, MUD_TRATMENT, MATERIALS, CATEGORY, SUBCATEGORY " +
                    " where REPORTS.RIGID = rigs.RIGID and " +
                    " reports.WELLID = WELLS.WELLID  and " +
                    " MUD_TRATMENT.MATID = MATERIALS.MATID and " +
                    " MUD_TRATMENT.REPORTID = REPORTS.REPORTID and " +
                    " CATEGORY.CatID = SUBCATEGORY.Catid and " +
                    " SUBCATEGORY.subid = MATERIALS.SubID and " +
                    " REPORTS.Date >= @C2 and REPORTS.Date <= @C3  ";



                string GroupQuery = " GROUP BY MATERIALS.MATName ,PackingQTYNewValue,UnitNewValue ";
                string GroupQuery2 = " GROUP BY MATERIALS.MATName,PackingQTYNewValue,UnitNewValue ";


                // " GROUP BY  MATERIALS.MATName,MUD_TRATMENT.PackingQTY,MUD_TRATMENT.UnitName; ";





                /*Query builder **/


                if ((int)CatComboBox.SelectedValue != 0)
                {
                    SQuery = SQuery + " and CATEGORY.catid = " + CatComboBox.SelectedValue;

                    SQuery2 = SQuery2 + " and CATEGORY.catid = " + CatComboBox.SelectedValue + GroupQuery;
                    DataTable dt = new DataTable();
                    DataTable dt2 = new DataTable();

                    // dt main();
                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
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
                                using (dt)
                                {
                                    ada.Fill(dt);
                                    dataGridView2.DataSource = dt;
                                    con.Close();

                                }
                            }


                            if ((int)SubCatComboBox.SelectedValue != 0)
                            {
                                dt.Rows.Clear();

                                SQuery = SQuery + " and SUBCATEGORY.subid = " + SubCatComboBox.SelectedValue;

                                using (SqlCommand cmd1 = new SqlCommand(SQuery, con))
                                {
                                    cmd1.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                                    cmd1.Parameters["@C2"].Value = dateTimePicker1.Value;

                                    cmd1.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                                    cmd1.Parameters["@C3"].Value = dateTimePicker2.Value;
                                    con.Open();

                                    using (SqlDataAdapter ada = new SqlDataAdapter(cmd1))
                                    {
                                        using (dt)
                                        {

                                            ada.Fill(dt);
                                            dataGridView2.DataSource = dt;
                                        }
                                        con.Close();
                                    }
                                    if ((int)MatComboBox.SelectedValue != 0)
                                    {
                                        dt.Rows.Clear();

                                        SQuery = SQuery + " and materials.matid =  " + MatComboBox.SelectedValue;
                                        using (SqlCommand cmd2 = new SqlCommand(SQuery, con))
                                        {
                                            cmd2.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                                            cmd2.Parameters["@C2"].Value = dateTimePicker1.Value;

                                            cmd2.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                                            cmd2.Parameters["@C3"].Value = dateTimePicker2.Value;
                                            con.Open();

                                            using (SqlDataAdapter ada = new SqlDataAdapter(cmd2))
                                            {
                                                using (dt)
                                                {

                                                    ada.Fill(dt);
                                                    dataGridView2.DataSource = dt;

                                                }
                                                con.Close();
                                            }

                                        }

                                    }

                                }
                            }
                        }

                    }
                    // dt main();

                    // dt2 sec();
                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                    {
                        using (SqlCommand cmd = new SqlCommand(SQuery2, con))
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
                                    //ds2.Tables.Add(dt2);
                                    //ds2.WriteXmlSchema("Summary.xml");


                                    dataGridView2.DataSource = dt2;
                                    this.dataGridView2.Columns[3].Width = 50;
                                    this.dataGridView2.Columns[1].Width = 70;




                                    con.Close();

                                }
                            }


                            if ((int)SubCatComboBox.SelectedValue != 0)
                            {
                                dt2.Rows.Clear();

                                SQuery2 = SQuery2.Replace(GroupQuery, " ");
                                GroupQuery = GroupQuery2;
                                SQuery2 = SQuery2 + " and SUBCATEGORY.subid = " + SubCatComboBox.SelectedValue + GroupQuery;

                                using (SqlCommand cmd1 = new SqlCommand(SQuery2, con))
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
                                            //ds2.Tables.Add(dt2);
                                            //ds2.WriteXmlSchema("Summary.xml");


                                            dataGridView2.DataSource = dt2;
                                            this.dataGridView2.Columns[3].Width = 50;
                                            this.dataGridView2.Columns[1].Width = 70;

                                        }
                                        con.Close();
                                    }
                                    if ((int)MatComboBox.SelectedValue != 0)
                                    {
                                        dt2.Rows.Clear();

                                        SQuery2 = SQuery2.Replace(GroupQuery, " ");
                                        GroupQuery = GroupQuery2;
                                        SQuery2 = SQuery2 + " and materials.matid = " + MatComboBox.SelectedValue + GroupQuery;
                                        using (SqlCommand cmd2 = new SqlCommand(SQuery2, con))
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
                                                    //ds2.Tables.Add(dt2);
                                                    //ds2.WriteXmlSchema("Summary.xml");


                                                    dataGridView2.DataSource = dt2;
                                                    this.dataGridView2.Columns[3].Width = 50;
                                                    this.dataGridView2.Columns[1].Width = 70;

                                                }
                                                con.Close();
                                            }

                                        }

                                    }

                                }
                            }
                        }

                    }
                    // dt2 sec();


                }

                if ((int)RigComboBox.SelectedValue != 0)
                {
                    DataTable dt4 = new DataTable();
                    dataGridView1.DataSource = null;
                    dataGridView2.DataSource = null;
                    dt4.Rows.Clear();
                    SQuery = SQuery + " and rigs.rigid = " + RigComboBox.SelectedValue;

                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
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
                                using (dt4)
                                {
                                    ada.Fill(dt4);
                                    dataGridView2.DataSource = dt4;

                                }
                                con.Close();
                            }

                        }


                    }


                    // dt2 sec();
                    DataTable dt42 = new DataTable();
                    SQuery2 = SQuery2.Replace(GroupQuery, " ");
                    GroupQuery = GroupQuery2;
                    SQuery2 = SQuery2 + " and rigs.rigid = " + RigComboBox.SelectedValue + GroupQuery;
                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                    {
                        using (SqlCommand cmd1 = new SqlCommand(SQuery2, con))
                        {
                            cmd1.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                            cmd1.Parameters["@C2"].Value = dateTimePicker1.Value;

                            cmd1.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                            cmd1.Parameters["@C3"].Value = dateTimePicker2.Value;
                            con.Open();


                            using (SqlDataAdapter ada = new SqlDataAdapter(cmd1))
                            {
                                using (dt42)
                                {
                                    ada.Fill(dt42);
                                    //ds42.Tables.Add(dt42);
                                    //ds42.WriteXmlSchema("Summary.xml");


                                    dataGridView2.DataSource = dt42;
                                    this.dataGridView2.Columns[3].Width = 50;
                                    this.dataGridView2.Columns[1].Width = 70;





                                    con.Close();

                                }
                            }


                        }

                    }
                    // dt2 sec();


                }

                if ((int)WellComboBox.SelectedValue != 0)
                {

                    DataTable dt2 = new DataTable();
                    DataTable dt23 = new DataTable();
                    dataGridView2.DataSource = null;
                  dt2.Rows.Clear();


                    SQuery = SQuery + " and wells.wellid = " + WellComboBox.SelectedValue;
                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                    {

                        using (SqlCommand cmd = new SqlCommand(SQuery, con))
                        {
                            cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                            cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

                            cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                            cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
                            con.Open();
                            ///   MessageBox.Show(SQuery);

                            using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                            {
                                using (dt2)
                                {
                                    ada.Fill(dt2);
                                    dataGridView2.DataSource = dt2;
                                }
                                con.Close();
                            }

                        }


                    }


                    // dt2 sec();
                    SQuery2 = SQuery2.Replace(GroupQuery, " ");
                    GroupQuery = GroupQuery2;
                    SQuery2 = SQuery2 + " and wells.wellid = " + WellComboBox.SelectedValue + GroupQuery;

                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                    {
                        using (SqlCommand cmd1 = new SqlCommand(SQuery2, con))
                        {
                            cmd1.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                            cmd1.Parameters["@C2"].Value = dateTimePicker1.Value;

                            cmd1.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                            cmd1.Parameters["@C3"].Value = dateTimePicker2.Value;
                            con.Open();


                            using (SqlDataAdapter ada = new SqlDataAdapter(cmd1))
                            {
                                using (dt23)
                                {
                                    ada.Fill(dt23);
                                    //ds23.Tables.Add(dt23);
                                    //ds23.WriteXmlSchema("Summary.xml");


                                    dataGridView2.DataSource = dt23;
                                    this.dataGridView2.Columns[3].Width = 50;
                                    this.dataGridView2.Columns[1].Width = 70;





                                    con.Close();

                                }
                            }


                        }

                    }
                    // dt2 sec();


                }






                DataTable dt3 = new DataTable();
                dataGridView1.DataSource = null;
                dt3.Rows.Clear();
                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                {
                    using (SqlCommand cmd = new SqlCommand(SQuery, con))
                    {

                        cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                        cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

                        cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                        cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
                        con.Open();
                        //  MessageBox.Show(SQuery);

                        using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                        {

                            using (dt3)
                            {
                                ada.Fill(dt3);
                                dataGridView1.DataSource = dt3;
                                this.dataGridView1.Columns[7].Visible = true;

                            }
                            //     dt3.Rows.Clear();
                            con.Close();
                        }
                    }

                }

                //this.dataGridView1.Columns["P-QTY"].Visible = false;
                //this.dataGridView1.Columns["unitt"].Visible = false; 
                //this.dataGridView2.Columns["P-QTY"].Visible = false;
                //this.dataGridView2.Columns["unitt"].Visible = false;
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

        private void MatComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DataRow dr;
            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");

            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct rigs.rigid,rigname from MATERIALS,SUBCATEGORY,CATEGORY,reports,MUD_TRATMENT,rigs ,wells " +
                " where wells.wellid= reports.wellid and reports.rigid = rigs.rigid and CATEGORY.CatID  = SUBCATEGORY.Catid " +
                " and SUBCATEGORY.subid = materials.subid and reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid " +
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
            SqlConnection conn = new SqlConnection(@"Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");

            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select distinct wells.wellid,wellname from wells,MATERIALS,SUBCATEGORY,CATEGORY,reports,MUD_TRATMENT,rigs " +
                "where wells.wellid= reports.wellid and reports.rigid = rigs.rigid and CATEGORY.CatID  = SUBCATEGORY.Catid and SUBCATEGORY.subid = materials.subid and " +
                "reports.reportid = MUD_TRATMENT.reportid and materials.matid=MUD_TRATMENT.matid " +
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
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            reset();
         
        }
        public void reset()
        {
            if (CatComboBox.SelectedIndex == -1)

            {
                MessageBox.Show("Please choose a Date and Category   ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
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
                dataGridView2.DataSource = null;
                RowsNuumlblNEW.Text = string.Empty;
                subtot.Text = string.Empty;
                SubTONEW.Text = string.Empty;

            }
         

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (ds2 != null)
            {
                //crystalReportViewer1.Visible = true;
                //dataGridView1.Visible = false;
                ////transefer data to crystalreportviewer
                //SummaryCrystalReport cr = new SummaryCrystalReport();
                //cr.SetDataSource(ds2);
                //crystalReportViewer1.ReportSource = cr;

            }
            if (ds42 != null)
            {
                //crystalReportViewer1.Visible = true;
                //dataGridView1.Visible = false;
                ////transefer data to crystalreportviewer
                //SummaryCrystalReport cr = new SummaryCrystalReport();
                //cr.SetDataSource(ds42);
                //crystalReportViewer1.ReportSource = cr;

            }
            if (ds23 != null)
            {
                //crystalReportViewer1.Visible = true;
                //dataGridView1.Visible = false;
                ////transefer data to crystalreportviewer
                //SummaryCrystalReport cr = new SummaryCrystalReport();
                //cr.SetDataSource(ds23);
                //crystalReportViewer1.ReportSource = cr;
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

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            this.Hide();
            frm3.Show();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Form4 frm4 = new Form4();
            this.Hide();
            frm4.Show();
        }
    }
}


