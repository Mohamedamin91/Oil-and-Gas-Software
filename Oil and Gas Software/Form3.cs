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
        //ID variable used in Updating and Deleting Record  
        int ID = 0;
        public Form3()
        {
          
            InitializeComponent();
            int id;


        }
        //Display Data in DataGridView  
        private void DisplayData()
        {
            con.Open();
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("select * from CATEGORY", con);
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        //Clear Data  
        private void ClearData()
        {
            txt_Name.Text = "";
            ID = 0;
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            int ID = 0;
            DisplayData();
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
            if (metroCheckBox1.Checked == true)
            {
                label2.Enabled = true;
                comboBox1.Enabled = true;

            }
            else
            {
                label2.Enabled = false;
                comboBox1.Enabled = false;
            }
        }
        //Insert Record  
        private void metroButton1_Click(object sender, EventArgs e)
        {
            if (txt_Name.Text != "" )
            {
                cmd = new SqlCommand("insert into CATEGORY(CatName) values(@name)", con);
                con.Open();
                cmd.Parameters.AddWithValue("@name", txt_Name.Text);
                if (DialogResult.Yes == MessageBox.Show("Do You Want Update ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
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
                        txt_Name.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    }
                }
            }

           
            
           

        }
        //Update Record  
        private void metroButton2_Click(object sender, EventArgs e)
        {
            if (txt_Name.Text != "" )
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
        //Delete Record  
        private void metroButton3_Click(object sender, EventArgs e)
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
    }
}
