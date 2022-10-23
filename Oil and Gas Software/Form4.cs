using iTextSharp.text;
using iTextSharp.text.pdf;
using MetroFramework.Forms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Oil_and_Gas_Software
{
    public partial class Form4 : MetroForm
    {
        SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
        SqlCommand cmd;
        int ID = 0;

        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            pictureBox1.Enabled = false;
            refreshdataMaterialSubategory();
            refreshdataMaterial();
            refreshdataRIGS();
            refreshdataWell();

        }
        private void ClearData()
        {
            txt_NamePQTY.Text = "";
            txt_NameUnitName.Text = "";

            ID = 0;
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


            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
              //dataGridView1.Columns[2].Visible = false;
            if (CatComboBox.SelectedIndex == -1)

            {
                MessageBox.Show("Please choose a Date and Category   ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                //  crystalReportViewer1.Visible = false;
                string SQuery = "select  rigs.Rigname  'Rig',WELLS.Wellname 'Well No'," +
                    " CATEGORY.CatName 'Category',SUBCATEGORY.Subname 'Subcategory', MATERIALS.MATName 'Materials'," +
                    " MUD_TRATMENT.QTY'QTY'," +
                    " MUD_TRATMENT.PackingQTY'PQTY',MUD_TRATMENT.UnitName 'Unit',REPORTS.DEPTH ,LAST24,DAYSSINCE 'Days since' , reports.Date 'Date' " +
                    " from " +
                    "RIGS, WELLS, REPORTS, MUD_TRATMENT, MATERIALS, CATEGORY, SUBCATEGORY where " +
                    "REPORTS.RIGID = rigs.RIGID and " +
                    " reports.WELLID = WELLS.WELLID  and " +
                    "MUD_TRATMENT.MATID = MATERIALS.MATID and" +
                    " MUD_TRATMENT.REPORTID = REPORTS.REPORTID and" +
                    " CATEGORY.CatID = SUBCATEGORY.Catid and" +
                    " SUBCATEGORY.subid = MATERIALS.SubID and " +
                    "reports.date >= @C2  and  reports.date <= @C3   ";

                string SQuery2 = " SELECT materials.matid,MATERIALS.MATName 'Material', SUM(MUD_TRATMENT.QTY) as Total,MUD_TRATMENT.PackingQTY 'P-QTy' ,MUD_TRATMENT.UnitName 'UnitT' , PackingQTYNewValue 'PQTY ', UnitNewValue 'Unit'   " +
                    " FROM RIGS, WELLS, REPORTS, MUD_TRATMENT, MATERIALS, CATEGORY, SUBCATEGORY " +
                    " where REPORTS.RIGID = rigs.RIGID and " +
                    " reports.WELLID = WELLS.WELLID  and " +
                    " MUD_TRATMENT.MATID = MATERIALS.MATID and " +
                    " MUD_TRATMENT.REPORTID = REPORTS.REPORTID and " +
                    " CATEGORY.CatID = SUBCATEGORY.Catid and " +
                    " SUBCATEGORY.subid = MATERIALS.SubID and " +
                    " REPORTS.Date >= @C2 and REPORTS.Date <= @C3  ";



                string GroupQuery = " GROUP BY MATERIALS.MATName,MUD_TRATMENT.PackingQTY,MUD_TRATMENT.UnitName ,PackingQTYNewValue,UnitNewValue, materials.matid ";
                string GroupQuery2 = " GROUP BY MATERIALS.MATName,MUD_TRATMENT.PackingQTY,MUD_TRATMENT.UnitName,PackingQTYNewValue,UnitNewValue, materials.matid ";


                // " GROUP BY  MATERIALS.MATName,MUD_TRATMENT.PackingQTY,MUD_TRATMENT.UnitName; ";





                /*Query builder **/


                if ((int)CatComboBox.SelectedValue != 0)
                {

                    SQuery2 = SQuery2 + " and CATEGORY.catid = " + CatComboBox.SelectedValue + GroupQuery;
                    DataTable dt2 = new DataTable();

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


                                    dataGridView1.DataSource = dt2;
                                    //this.dataGridView2.Columns[3].Width = 50;
                                    //this.dataGridView2.Columns[1].Width = 70;
                                    this.dataGridView1.Columns["matid"].Visible = false;
                                    this.dataGridView1.Columns["P-QTy"].Visible = false;
                                    this.dataGridView1.Columns["UnitT"].Visible = false;



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


                                            dataGridView1.DataSource = dt2;
                                            //this.dataGridView2.Columns[3].Width = 50;
                                            //this.dataGridView2.Columns[1].Width = 70;
                                            this.dataGridView1.Columns["matid"].Visible = false;
                                            this.dataGridView1.Columns["P-QTy"].Visible = false;
                                            this.dataGridView1.Columns["UnitT"].Visible = false;

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


                                                    dataGridView1.DataSource = dt2;
                                                    //this.dataGridView2.Columns[3].Width = 50;
                                                    //this.dataGridView2.Columns[1].Width = 70;
                                                    this.dataGridView1.Columns["matid"].Visible = false;
                                                    this.dataGridView1.Columns["P-QTy"].Visible = false;
                                                    this.dataGridView1.Columns["UnitT"].Visible = false;

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


                                    dataGridView1.DataSource = dt42;
                                    this.dataGridView1.Columns[3].Width = 50;
                                    this.dataGridView1.Columns[1].Width = 70;





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
                    dt23.Rows.Clear();





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


                                    dataGridView1.DataSource = dt23;
                                    this.dataGridView1.Columns[3].Width = 50;
                                    this.dataGridView1.Columns[1].Width = 70;





                                    con.Close();

                                }
                            }


                        }

                    }
                    // dt2 sec();


                }
                else
                {
                    //DataTable dt3 = new DataTable();
                    ////    dataGridView2.DataSource = null;
                    //dt3.Rows.Clear();
                    //using (SqlConnection con = new SqlConnection("Data Source=192.168.1.8;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                    //{
                    //    using (SqlCommand cmd = new SqlCommand(SQuery2, con))
                    //    {

                    //        cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                    //        cmd.Parameters["@C2"].Value = dateTimePicker1.Value;

                    //        cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Date));
                    //        cmd.Parameters["@C3"].Value = dateTimePicker2.Value;
                    //        con.Open();
                    //        //  MessageBox.Show(SQuery);

                    //        using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                    //        {

                    //            using (dt3)
                    //            {
                    //                ada.Fill(dt3);
                    //              //  dataGridView1.DataSource = dt3;
                    //                this.dataGridView1.Columns[7].Visible = true;

                    //            }
                    //            //     dt3.Rows.Clear();
                    //            con.Close();
                    //        }
                    //    }

                    //}
                //    MessageBox.Show("Please choose Category", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);


                }

                int a = 0;
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    {
                        a += Convert.ToInt32(r.Cells[2].Value);
                        totqty.Text = a.ToString();
                        this.dataGridView1.Columns["matid"].Visible = false;
                        this.dataGridView1.Columns["P-QTy"].Visible = false;
                        this.dataGridView1.Columns["UnitT"].Visible = false;

                    }
                }

            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            refreshdataMaterialCategory();

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            refreshdataMaterialCategory();

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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        // update
        private void button4_Click(object sender, EventArgs e)
        {
            ///category 


            if (txt_NamePQTY.Text != "" || txt_NameUnitName.Text != "")
            {
                cmd = new SqlCommand("update MUD_TRATMENT set PackingQTYNewValue=@qty, UnitNewValue=@unit where matid=@id", con);
                con.Open();
                cmd.Parameters.AddWithValue("@id", ID);
                cmd.Parameters.AddWithValue("@qty", txt_NamePQTY.Text);
                cmd.Parameters.AddWithValue("@unit", txt_NameUnitName.Text);
                if (DialogResult.Yes == MessageBox.Show("Do You Want Update ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                {
                    // do what u want
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Record Updated Successfully");
                    con.Close();
                    button1.PerformClick();
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
                        txt_NamePQTY.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                        txt_NameUnitName.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();


                    }
                }
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            PrintPreviewDialog ppd = new PrintPreviewDialog();

            PrintDocument pd = new PrintDocument();

            pd.PrintPage += new PrintPageEventHandler(print);

            ppd.Document = pd;

            ppd.ShowDialog();

        }
        void print(object sender, PrintPageEventArgs e)
        {


            int width = 1000;
            int height = 1800;

            //bill_groupbox.Width,bill_groupbox.Height

            Bitmap bmp = new Bitmap(width, height);

            System.Drawing.Rectangle rec = new System.Drawing.Rectangle(0, 0, groupBox2.Width, height);

            groupBox2.DrawToBitmap(bmp, rec);


            e.Graphics.DrawImage(bmp, rec);

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            this.Hide();
            frm1.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            this.Hide();
            frm3.Show();
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            Form2 frm2 = new Form2();
            this.Hide();
            frm2.Show();
        }
    }
}

