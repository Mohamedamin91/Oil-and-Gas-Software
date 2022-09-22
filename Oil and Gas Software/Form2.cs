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
        public Form2()
        {
            InitializeComponent();
        }
        //public void BindTotal()
        //{
        //    SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@");
        //    con.Open();
        //    SqlCommand cmd1 = new SqlCommand(" select  COUNT(distinct KeywordName)'material' from KEYWORDS ", con);
        //    SqlCommand cmd2 = new SqlCommand(" select  count (distinct Rign)  'RigNo' from Rigs  ", con);
        //    SqlCommand cmd3 = new SqlCommand(" select  count (distinct Rigname) 'Well No' from Rigs ", con);
        //    SqlCommand cmd4 = new SqlCommand(" select  COUNT(distinct Rigname)'Well No (type) Contain data' from rigs where Contain = 1 ", con);
        //    SqlCommand cmd5 = new SqlCommand(" select  COUNT(distinct Rigname)'Well No (type) Non Contain data' from rigs where Contain = 0 ", con);
        //    SqlCommand cmd6 = new SqlCommand(" select  COUNT(distinct CatID)'Category' from Category", con);
        //    SqlCommand cmd7 = new SqlCommand(" select  COUNT( distinct Subid)'SubCategory' from Subcatogory", con);
        //    SqlCommand cmd8 = new SqlCommand(" select  COUNT(distinct KeywordName)'Material Without' from KEYWORDS where CatID = 0 ", con);
        //    SqlCommand cmd9 = new SqlCommand(" select COUNT ( distinct RigID) from FILES ", con);


        //    cmd1.ExecuteNonQuery();
        //    cmd2.ExecuteNonQuery();
        //    cmd3.ExecuteNonQuery();
        //    cmd4.ExecuteNonQuery();
        //    cmd5.ExecuteNonQuery();
        //    cmd6.ExecuteNonQuery();
        //    cmd7.ExecuteNonQuery();
        //    cmd8.ExecuteNonQuery();
        //    cmd9.ExecuteNonQuery();
        //    reader = cmd1.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label4.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();
        //    con.Open();
        //    reader = cmd2.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label5.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();
        //    con.Open();
        //    reader = cmd3.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label6.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();
        //    con.Open();
        //    reader = cmd4.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label7.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();
        //    con.Open();
        //    reader = cmd5.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label10.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();
        //    con.Open();
        //    con.Close();
        //    con.Open();
        //    reader = cmd6.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label11.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();
        //    con.Open();
        //    con.Close();
        //    con.Open();
        //    reader = cmd7.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label13.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();
        //    con.Open();
        //    con.Close();
        //    con.Open();
        //    reader = cmd8.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label16.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();
        //    con.Open();
        //    reader = cmd9.ExecuteReader();
        //    if (reader.HasRows)
        //    {
        //        while (reader.Read())
        //        {
        //            label17.Text = reader[0].ToString();
        //        }
        //    }
        //    con.Close();

        //}
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
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select distinct RigID, RIGN from Rigs where RIGN !=''    ", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {
                dr.ItemArray = new object[] { 0, "--Select Rigs--" };
                dt.Rows.InsertAt(dr, 0);
                RigComboBox.ValueMember = "RigID";
                RigComboBox.DisplayMember = "RIGN";
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
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@");
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
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@");
            con.Open();
            SqlCommand cmd = new SqlCommand("select CatID,CatName from Category order by CatName", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {
                dr.ItemArray = new object[] { 0, "--Select Category--" };
                dt.Rows.InsertAt(dr, 0);
                this.CatComboBox.ValueMember = "CatID";
               this. CatComboBox.DisplayMember = "CatName";
               this. CatComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
               this. CatComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
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
                MessageBox.Show("Please choose a folder to import 'Materials'  ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void refreshdataMaterial()
        {
            DataRow dr;
            SqlConnection con = new SqlConnection(@"Data Source=192.168.1.105;Initial Catalog=OILREPORT;Persist Security Info=True;User ID=sa;password=Ram72763@");
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

            SqlCommand cmd = new SqlCommand("select RigID,Rigname from Rigs where Rigname !='' order by Rigname ", con);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dr = dt.NewRow();
            if (dt != null)
            {

                dr.ItemArray = new object[] { 0, "All" };
                dt.Rows.InsertAt(dr, 0);
                WellComboBox.ValueMember = "RigID";
                WellComboBox.DisplayMember = "Rigname";
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
            refreshdataMaterialCategory();
        }
    }
}
