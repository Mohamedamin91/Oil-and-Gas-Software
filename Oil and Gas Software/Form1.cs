using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Aspose.Cells;
using MetroFramework.Forms;
using System.Text.RegularExpressions;
using System.IO;

namespace Oil_and_Gas_Software
{
    public partial class Form1 : MetroForm
    {
        OpenFileDialog opf = new OpenFileDialog();
        SaveFileDialog svg = new SaveFileDialog();
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        DataTable dt = new DataTable();
        SqlDataReader reader;
        public Form1()
        {
            InitializeComponent();
        }
        public void BindGV()
        {
            dt.Rows.Clear();
            using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT " +
                    " [MATName] as 'Material',[Val]as'Qty',PackingQTY as 'Packaging Quantity',UnitName as 'Unit',[Rigname] 'Well No (Type)',[Date] " +
                    "FROM [OILREPORT2].[dbo].[FILES],OILREPORT2.dbo.MATERIALS,OILREPORT2.dbo.Rigs" +
                    " where REPORTS.MATID = MATERIALS.MATID and FILES.RigID = Rigs.RigID  ", con))
                {

                    using (SqlDataAdapter ada = new SqlDataAdapter(cmd))
                    {
                        using (dt)
                        {
                            ada.Fill(dt);
                            dataGridView1.Visible = true;
                            dataGridView1.DataSource = dt;
                          //  RowsNuumlblNEW.Visible = true;
                          //  RowsNuumlblNEW.Text = dataGridView1.Rows.Count.ToString("N0");
                        }
                    }

                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            BindGV();
        }

        private void BrowseBtn_Click(object sender, EventArgs e)
        {
            // dt.Rows.Clear();
            //method to remove MATERIALS
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                int MaterialID;
                int RigID;
                txtfilepath.Text = fbd.SelectedPath;
                DirectoryInfo dir = new DirectoryInfo(fbd.SelectedPath);

                foreach (var file in dir.GetFiles())
                {
                    string extractedfullDATE = "";
                    string extractedDATEONLY = "";
                    string extractedRIGNAME = "";
                    string extractedRIGDATA = "";
                    string extractedRIGDepth = "";
                    string last24 = "";
                    string DaysSince = "";
                    string extratcRIGNameNEWONE = "";
                    var Depth = "";
                    var Rigname = "";
                    RigID = 0;
                    var workbook = new Workbook(file.FullName);
                    workbook.Save(file.FullName + ".txt");
                    using (var sr1 = new StreamReader(file.FullName, true))
                    {
                        string FullData = File.ReadAllText(workbook.FileName);
                        int From10 = FullData.LastIndexOf("MUD TREATMENT") + "MUD TREATMENT".Length;
                        int To10 = FullData.LastIndexOf("BIT DATA");
                        string check = "";
                        if (From10 != -1 && To10 != -1)
                        {
                            check = FullData.Substring(From10, To10 - From10);
                            if (check.Contains(")"))
                            {
                                /**  extracting Depth*/
                                int From0 = FullData.IndexOf("Depth") + "Depth".Length;
                                int To0 = FullData.IndexOf("Liner Size");
                                extractedRIGDepth = FullData.Substring(From0, To0 - From0);
                                extractedRIGDepth = extractedRIGDepth.Replace("\"", "");
                                extractedRIGDepth = Regex.Replace(extractedRIGDepth, @"\s+", " ");
                                extractedRIGDepth = extractedRIGDepth.Replace(") ", ")" + System.Environment.NewLine);
                                extractedRIGDepth = extractedRIGDepth.Trim();
                                var input = extractedRIGDepth;
                                Depth = Regex.Replace(input.Split()[0], @"[^0-9a-zA-Z\ ]+", "");
                                /**  extracting Depth*/
                                /**start Exrtracting process for reg name , reg date , mud data */
                                int From1 = FullData.IndexOf("Date") + "Date".Length;
                                int To1 = FullData.IndexOf("Well No");
                                extractedfullDATE = FullData.Substring(From1, To1 - From1);
                                // delete the day and just keep the date
                                extractedDATEONLY = extractedfullDATE.Substring(extractedfullDATE.Length - 15, 15);
                                extractedDATEONLY = extractedDATEONLY.Replace("' ", "");
                                int From2 = FullData.IndexOf("Well No (Type) :") + "Well No (Type) :".Length;
                                int To2 = FullData.IndexOf("Charge #");
                                extractedRIGNAME = FullData.Substring(From2, To2 - From2);
                                // remove between bractise /** to 
                                extractedRIGNAME = Regex.Replace(extractedRIGNAME, @"\([^)]*\)", "");
                                extractedRIGNAME = extractedRIGNAME.Replace(")", "");
                                extractedRIGNAME = extractedRIGNAME.Replace(";", "");
                                extractedRIGNAME = extractedRIGNAME.Replace(",", "");
                                extractedRIGNAME = extractedRIGNAME.Replace(" '' '' ", "");
                                extractedRIGNAME = extractedRIGNAME.Replace("\"", "");
                                int space1 = extractedRIGNAME.IndexOf(" ");
                                Rigname = (extractedRIGNAME.Substring(0, space1));
                                Rigname = Rigname.TrimStart();
                                Rigname = Rigname.TrimEnd();
                                Rigname = Rigname.Trim();

                                int From3 = FullData.LastIndexOf("MUD TREATMENT") + "MUD TREATMENT".Length;
                                int To3 = FullData.LastIndexOf("BIT DATA");
                                extractedRIGDATA = FullData.Substring(From3, To3 - From3);
                                /**End Exrtracting process for reg name , reg date , mud data */
                                /** Start Styilng extracted mud data */
                                extractedRIGDATA = extractedRIGDATA.Replace("\"", "");
                                extractedRIGDATA = Regex.Replace(extractedRIGDATA, @"\s+", " ");
                                extractedRIGDATA = extractedRIGDATA.Replace(") ", ")" + System.Environment.NewLine);
                                extractedRIGDATA = extractedRIGDATA.Trim();
                                /** end Styilng extracted mud data */
                                /**  extracting last 24*/
                                int From4 = FullData.IndexOf("Last 24 hr operations") + "Last 24 hr operations".Length;
                                int To4 = FullData.IndexOf("Next 24 hr plan");
                                last24 = FullData.Substring(From4, To4 - From4);
                                last24 = last24.Replace(",", System.Environment.NewLine);
                                last24 = last24.TrimStart();
                                last24 = last24.TrimEnd();
                                /**  extracting last 24*/
                                /**  extracting DaysSince*/
                                int From5 = FullData.IndexOf("Days Since Spud/Comm (Date)") + "Days Since Spud/Comm (Date)".Length;
                                int To5 = FullData.IndexOf("Formation tops");
                                DaysSince = FullData.Substring(From5, To5 - From5);
                                DaysSince = DaysSince.TrimStart();
                                DaysSince = DaysSince.TrimEnd();
                                DaysSince = Regex.Replace(DaysSince, @"\s", "");
                                /**  extracting DaysSince*/

                                /*start extrat RIGNA**//***/
                                string FinalString0 = "";

                                int From6 = FullData.LastIndexOf("Wellbores:") + "Wellbores:".Length;
                                int To6 = FullData.LastIndexOf("Foreman(s)");
                                FinalString0 = FullData.Substring(From6, To6 - From6);
                                List<string> EXtractRIGNAMELIST = FinalString0.Split('\n').ToList();
                                EXtractRIGNAMELIST.RemoveAt(0);
                                foreach (var word in EXtractRIGNAMELIST)
                                {
                                    extratcRIGNameNEWONE = word.ToString();
                                }
                                extratcRIGNameNEWONE = extratcRIGNameNEWONE.TrimStart();
                                extratcRIGNameNEWONE = extratcRIGNameNEWONE.TrimEnd();
                                extratcRIGNameNEWONE = extratcRIGNameNEWONE.Trim();
                                /*end extrat RIGNA**//***/

                                /** start insert rig info contain mudtreatment*/

                                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                                {
                                    //SqlCommand check_RIG1 = new SqlCommand("SELECT Count(*) FROM [Rigs] WHERE ([Rigname] = @C1 and RIGN = NULL)", con);
                                    //con.Open();
                                    //check_RIG1.Parameters.AddWithValue("@C1", Rigname);
                                    //int RIGExist1 = (int)check_RIG1.ExecuteScalar();
                                    //if (RIGExist1 == 0)
                                    //{
                                    //    SqlCommand cmd = new SqlCommand(" UPDATE [Rigs] SET [RIGN] = @C1 WHERE [Rigname] = @C2 ", con);
                                    //    cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                    //    cmd.Parameters["@C1"].Value = extratcRIGNameNEWONE.ToString();
                                    //    cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.VarChar));
                                    //    cmd.Parameters["@C2"].Value = Rigname.ToString();
                                    //    cmd.ExecuteNonQuery();
                                    //    con.Close();
                                    //}

                                    // date +
                                    SqlCommand check_RIG2 = new SqlCommand("SELECT Count(*) FROM [Rigs] WHERE ([Rigname] = @C1 and RIGN= @C2 )", con);
                                    con.Open();
                                    check_RIG2.Parameters.AddWithValue("@C1", Rigname);
                                    check_RIG2.Parameters.AddWithValue("@C2", extratcRIGNameNEWONE.ToString());
                                    int RIGExist2 = (int)check_RIG2.ExecuteScalar();
                                    if (RIGExist2 == 0)
                                    {
                                        using (SqlCommand cmd = new SqlCommand("INSERT INTO Rigs (Rigname,Depth,last24,DaysSince,Contain,RIGN) VALUES (@C1,@C2,@C3,@C4,1,@C5)", con))
                                        {
                                            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                            cmd.Parameters["@C1"].Value = Rigname.ToString();
                                            cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.VarChar));
                                            cmd.Parameters["@C2"].Value = Depth.ToString();
                                            cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.VarChar));
                                            cmd.Parameters["@C3"].Value = last24.ToString();
                                            cmd.Parameters.Add(new SqlParameter("@C4", SqlDbType.VarChar));
                                            cmd.Parameters["@C4"].Value = DaysSince.ToString();
                                            cmd.Parameters.Add(new SqlParameter("@C5", SqlDbType.VarChar));
                                            cmd.Parameters["@C5"].Value = extratcRIGNameNEWONE.ToString();
                                            cmd.ExecuteNonQuery();
                                            using (SqlCommand cmd1 = new SqlCommand("SELECT (RigID)  FROM  Rigs WHERE Rigname=@C1", con))
                                            {
                                                cmd1.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                                cmd1.Parameters["@C1"].Value = Rigname.ToString();

                                                RigID = (int)cmd1.ExecuteScalar();
                                            }
                                            //   MessageBox.Show("The new Materials has been added successfully .", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        }
                                    }

                                    else
                                    {
                                        using (SqlCommand cmd = new SqlCommand("SELECT (RigID)  FROM  Rigs WHERE Rigname=@C1", con))
                                        {
                                            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                            cmd.Parameters["@C1"].Value = Rigname.ToString();
                                            RigID = (int)cmd.ExecuteScalar();
                                        }

                                    }
                                    con.Close();
                                }

                                ///** end insert rig info contain mudtreatment */

                                /** start section (Mud Treatment )  */
                                List<string> FinalString3LIST = extractedRIGDATA.Split('\n').ToList();
                                foreach (var word in FinalString3LIST)
                                {
                                    //  MessageBox.Show(word);
                                    MaterialID = 0;
                                    string strDate = extractedDATEONLY;
                                    string[] dateString = strDate.Split('/');
                                    DateTime enter_date = Convert.ToDateTime(dateString[0] + "/" + dateString[1] + "/" + dateString[2]);
                                    var newenter_date = enter_date.ToShortDateString();

                                    int qous = word.IndexOf("(");
                                    int space = word.LastIndexOf(" ");
                                    var MValue = (word.Substring(space, qous - space));
                                    MValue = MValue.Replace("(", " ");
                                    var keyword = (word.Substring(word.IndexOf(word), space));
                                    keyword = keyword.TrimStart();
                                    keyword = keyword.TrimEnd();
                                    keyword = keyword.Trim();

                                    /*extract value between  brackets */
                                    int start = word.IndexOf("(") + 1;
                                    int end = word.IndexOf(")", start);
                                    string brackets = word.Substring(start, end - start);
                                    Regex re = new Regex("([0-9]+)([A-Z]+)");
                                    Match result2 = re.Match(brackets);
                                    string PackingQTY = result2.Groups[1].Value;
                                    string UnitName = result2.Groups[2].Value;

                                    /*extract value between  brackets */
                                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                                    {

                                        SqlCommand check_KEYWORD = new SqlCommand("SELECT Count(*) FROM [MATERIALS] WHERE ([MATName] = @C1)", con);
                                        con.Open();
                                        check_KEYWORD.Parameters.AddWithValue("@C1", keyword.ToString());
                                        int KEYWORDExist = (int)check_KEYWORD.ExecuteScalar();
                                        if (KEYWORDExist == 0)
                                        {
                                            using (SqlCommand cmd = new SqlCommand("INSERT INTO MATERIALS(MATName,UnitName,PackingQTY) VALUES (@C1,@C2,@C3)", con))
                                            {
                                                cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                                cmd.Parameters["@C1"].Value = keyword.ToString();
                                                cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.VarChar));
                                                cmd.Parameters["@C2"].Value = UnitName.ToString();
                                                cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.VarChar));
                                                cmd.Parameters["@C3"].Value = PackingQTY;
                                                cmd.ExecuteNonQuery();
                                                using (SqlCommand cmd1 = new SqlCommand("SELECT (MATID)  FROM  MATERIALS WHERE MATName=@C1", con))
                                                {
                                                    cmd1.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                                    cmd1.Parameters["@C1"].Value = keyword.ToString();

                                                    MaterialID = (int)cmd1.ExecuteScalar();
                                                }
                                                //   MessageBox.Show("The new Materials has been added successfully .", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            }
                                        }
                                        else
                                        {
                                            using (SqlCommand cmd = new SqlCommand("SELECT (MATID)  FROM  MATERIALS WHERE MATName=@C1", con))
                                            {
                                                cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                                cmd.Parameters["@C1"].Value = keyword.ToString();
                                                MaterialID = (int)cmd.ExecuteScalar();
                                            }
                                            //  MessageBox.Show("Info : The Material already Exist", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        }
                                    }
                                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                                    {
                                        SqlCommand cmd0 = new SqlCommand(" UPDATE [MATERIALS] SET [PackingQTY] = 1 WHERE [PackingQTY]= 0 ", con);
                                        SqlCommand cmd1 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] = 1 /*,[SubID]=1 */ WHERE [MATName]= 'BARITE' or [MATName]= 'BA-NF' or [MATName]= 'BA-AM' or [MATName]= 'BA-AL'  or [MATName]= 'BA-AF' or [MATName]= 'BA-OR' or [MATName]= 'KI-BARITE' or [MATName]= 'BA-PERF' or [MATName]= 'BA-ESNAAD' or [MATName]= 'BA-BAR'  or[MATName]= 'BA-IBC' or [MATName]= 'BA-OM'  or[MATName]= 'BA-AGMED' or [MATName]= 'BA-MID'  or[MATName]= 'BA-POUR' or [MATName]= 'BA-ATAD'  or[MATName]= 'BA' ", con);
                                        SqlCommand cmd2 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] = 2 /*,[SubID]=2*/   WHERE [MATName]= 'BA-NF-BULK' or [MATName]= 'BA-BA-BULK' or [MATName]= 'BA-AR-BULK' or [MATName]= 'BA-DMF-BULK'or[MATName]= 'BA BULK'", con);
                                        SqlCommand cmd3 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] = 3 /*,[SubID]=3*/   WHERE [MATName]= 'CABR2' or [MATName]= 'CABR2-MI' or [MATName]= 'CABR2-HAL' or [MATName]= 'CABR2-TET'or [MATName]= 'CABR2-OS' or [MATName]= 'CABR2-JOR' or [MATName]= 'CABR2-AGR' or [MATName]= 'CABR2-SHA' or[MATName]= 'CABR2-JIA' or [MATName]= 'CABR2-WEI'  or[MATName]= 'CABR2-ALB'", con);
                                        SqlCommand cmd4 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] = 4 /*,[SubID]=4 */  WHERE [MATName]= 'CABR2-SHO' ", con);
                                        SqlCommand cmd5 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] = 5 /*,[SubID]=5  */ WHERE [MATName]= 'CACL2-80CC' or [MATName]= 'CACL2-80' or [MATName]= 'CACL2-80ME' or [MATName]= 'CACL2-80TET'or [MATName]= 'CACL2-80SO' or[MATName]= 'CACL2-80TAN' or[MATName]= 'CACL2-80DW' or[MATName]= 'CACL2-80GC' or[MATName]= 'CACL2-80QH' or [MATName]= 'CACL2-80WE'  or[MATName]= 'CACL2-80CH'or [MATName]= 'CACL2-80IN'  or[MATName]= 'CACL2-80TEEU' or [MATName]= 'CACL2-80LIA' ", con);
                                        SqlCommand cmd6 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] = 6 /*,[SubID]=6 */  WHERE [MATName]= 'CACL2-98' or [MATName]= 'CACL2-98-BB' or [MATName]= 'CACL2-98CH-BB' or [MATName]= 'CACL2-98BA'or [MATName]= 'CACL2-98BH-BB' or[MATName]= 'CACL2-98DW' or[MATName]= 'CACL2-98TA-BB' or[MATName]= 'CACL2-98JB-BB' or[MATName]= 'CACL2-98WE-BB' or [MATName]= 'CACL2-98JB'  or[MATName]= 'CACL2-98TCE'or [MATName]= 'CACL2-98TE-BB'  or[MATName]= 'CACL2-98IN-BB' ", con);
                                        SqlCommand cmd7 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] = 7 /*,[SubID]=7 */  WHERE [MATName]= 'LIG-OBM' or [MATName]= 'TNATHN'or [MATName]= 'LIGNITE' or [MATName]= 'CACL2-98BA' ", con);
                                        SqlCommand cmd8 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] =8  /*,[SubID]=8 */  WHERE [MATName]= 'MRBL-C-NF-BB' or [MATName]= 'MRBL-C-SEP' or [MATName]= 'MRBL-C-SEP-BB' or [MATName]= 'MRBL-C-NF'or [MATName]= 'MRBL-C' or[MATName]= 'MRBL-C-BH-BB' or[MATName]= 'MRBL-C-BH' ", con);
                                        SqlCommand cmd9 = new SqlCommand(" UPDATE [MATERIALS] SET [CatID] =9  /*,[SubID]=9*/   WHERE [MATName]= 'MRBL-F-BH' or [MATName]= 'MRBL-F-SEP' or [MATName]= 'MRBL-F-NF-BB' or [MATName]= 'MRBL-F-SEP-BB'or [MATName]= 'MRBL-F-NF' or[MATName]= 'MRBL-F' or[MATName]= 'MRBL-C-BH'or[MATName]= 'MRBL-F-BH-BB'or[MATName]= 'MRBL-F-AEC' or [MATName]= 'MRBL-F-AEC-BB'  or[MATName]= 'MRBL-F-TP-BB' ", con);
                                        SqlCommand cmd10 = new SqlCommand("UPDATE [MATERIALS] SET [CatID] =10 /*,[SubID]=10*/  WHERE [MATName]= 'MRBL-M-NF-BB' or [MATName]= 'MRBL-M-SEP' or [MATName]= 'MRBL-M-BH' or [MATName]= 'MRBL-MED-BB'or[MATName]= 'MRBL-M-MI' or[MATName]= 'MRBL-M-MI-BB' or[MATName]= 'MRBL-M-NF'or[MATName]= 'MRBL-M-SEP-BB'or[MATName]= 'MRBL-M'  ", con);
                                        SqlCommand cmd11 = new SqlCommand("UPDATE [MATERIALS] SET [CatID] =11 /*,[SubID]=11*/  WHERE [MATName]= 'GELTONE' or [MATName]= 'CARBOGEL-II' or [MATName]= 'OILGEL'or [MATName]= 'NAFGEL' ", con);
                                        SqlCommand cmd12 = new SqlCommand("UPDATE [MATERIALS] SET [CatID] =12 /*,[SubID]=12*/  WHERE [MATName]= 'DURATONE' or [MATName]= 'CARBOTROL' or [MATName]= 'VRSLIG'or [MATName]= 'NAFTROL' ", con);
                                        SqlCommand cmd13 = new SqlCommand("UPDATE [MATERIALS] SET [CatID] =13 /*,[SubID]=13 */ WHERE [MATName]= 'RESINEX II' or [MATName]= 'GMPRO-RX' or [MATName]= 'RENZI_SPNH' ", con);
                                        SqlCommand cmd14 = new SqlCommand("UPDATE [MATERIALS] SET [CatID] =14 /*,[SubID]=14*/  WHERE [MATName]= 'NACL' or [MATName]= 'NACL-SAR' or [MATName]= 'NACL-RY' or [MATName]= 'NACL-GC' or[MATName]= 'NACL-SEP' ", con);
                                        SqlCommand cmd15 = new SqlCommand("UPDATE [MATERIALS] SET [CatID] =15 /*,[SubID]=15*/  WHERE [MATName]= 'NACL-DEL' ", con);
                                        SqlCommand cmd16 = new SqlCommand("UPDATE [MATERIALS] SET [CatID] =0 /*,[SubID]= 107*/ WHERE [CatID]  IS NULL /*or [SubID] IS NULL*/ ", con);
                                        SqlCommand cmd17 = new SqlCommand("UPDATE MATERIALS SET MATERIALS.SubID = B.Subid FROM Category A, SUBCATEGORY B   WHERE  A.CatID = B.Catid and  a.CatID = MATERIALS.CatID ", con);
                                        /* for update subid in table keyword automaticly  regarding */
                                        /**       UPDATE MATERIALS SET MATERIALS.SubID = B.Subid FROM Category A, SUBCATEGORY B   WHERE  A.CatID = B.Catid and  a.CatID = MATERIALS.CatID*/


                                        con.Open();
                                        cmd0.ExecuteNonQuery();
                                        cmd1.ExecuteNonQuery();
                                        cmd2.ExecuteNonQuery();
                                        cmd3.ExecuteNonQuery();
                                        cmd4.ExecuteNonQuery();
                                        cmd5.ExecuteNonQuery();
                                        cmd6.ExecuteNonQuery();
                                        cmd7.ExecuteNonQuery();
                                        cmd8.ExecuteNonQuery();
                                        cmd9.ExecuteNonQuery();
                                        cmd10.ExecuteNonQuery();
                                        cmd11.ExecuteNonQuery();
                                        cmd12.ExecuteNonQuery();
                                        cmd13.ExecuteNonQuery();
                                        cmd14.ExecuteNonQuery();
                                        cmd15.ExecuteNonQuery();
                                        cmd16.ExecuteNonQuery();
                                        cmd17.ExecuteNonQuery();

                                    }
                                    using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                                    {
                                        SqlCommand check_FILE = new SqlCommand(" SELECT Count(*) from REPORTS,MATERIALS where REPORTS.MATID=MATERIALS.MATID and" +
                                            "  RigID=@C3 and date=@C4 and MATName=@C1", con);
                                        con.Open();
                                        check_FILE.Parameters.AddWithValue("@C1", keyword.ToString());
                                        check_FILE.Parameters.AddWithValue("@C3", RigID.ToString());
                                        check_FILE.Parameters.AddWithValue("@C4", enter_date.ToString());
                                        int FILEExist = (int)check_FILE.ExecuteScalar();
                                        if (FILEExist == 0)
                                        {

                                            using (SqlCommand cmd = new SqlCommand("INSERT INTO REPORTS(MATID,Val,RigID,Date,Contain) VALUES (@C1,@C2,@C3,@C4,'1')", con))
                                            {

                                                cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
                                                cmd.Parameters["@C1"].Value = MaterialID;

                                                cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.NVarChar));
                                                cmd.Parameters["@C2"].Value = MValue;

                                                cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.Int));
                                                cmd.Parameters["@C3"].Value = RigID;

                                                cmd.Parameters.Add(new SqlParameter("@C4", SqlDbType.Date));
                                                cmd.Parameters["@C4"].Value = enter_date;


                                                cmd.ExecuteNonQuery();
                                                //   MessageBox.Show(" inserted rig  " + RigID.ToString());

                                            }
                                        }
                                        else
                                        {
                                        }
                                    }

                                }
                                /** end  section (Mud Treatment )  */
                            }
                            else
                            {
                                /** start insert rig info non contain mudtreatment*/
                                /**  extracting Depth*/
                                int From0 = FullData.IndexOf("Depth") + "Depth".Length;
                                int To0 = FullData.IndexOf("Liner Size");
                                extractedRIGDepth = FullData.Substring(From0, To0 - From0);
                                extractedRIGDepth = extractedRIGDepth.Replace("\"", "");
                                extractedRIGDepth = Regex.Replace(extractedRIGDepth, @"\s+", " ");
                                extractedRIGDepth = extractedRIGDepth.Replace(") ", ")" + System.Environment.NewLine);
                                extractedRIGDepth = extractedRIGDepth.Trim();
                                var input = extractedRIGDepth;
                                Depth = Regex.Replace(input.Split()[0], @"[^0-9a-zA-Z\ ]+", "");
                                /**start Exrtracting process for reg name , reg date , mud data */
                                int From1 = FullData.IndexOf("Date") + "Date".Length;
                                int To1 = FullData.IndexOf("Well No");
                                extractedfullDATE = FullData.Substring(From1, To1 - From1);
                                // delete the day and just keep the date
                                extractedDATEONLY = extractedfullDATE.Substring(extractedfullDATE.Length - 15, 15);
                                extractedDATEONLY = extractedDATEONLY.Replace("' ", "");

                                int From2 = FullData.IndexOf("Well No (Type) :") + "Well No (Type) :".Length;
                                int To2 = FullData.IndexOf("Charge #");
                                extractedRIGNAME = FullData.Substring(From2, To2 - From2);
                                // remove between bractise /** to 
                                extractedRIGNAME = Regex.Replace(extractedRIGNAME, @"\([^)]*\)", "");
                                extractedRIGNAME = extractedRIGNAME.Replace(")", "");
                                extractedRIGNAME = extractedRIGNAME.Replace(";", "");
                                extractedRIGNAME = extractedRIGNAME.Replace(",", "");
                                extractedRIGNAME = extractedRIGNAME.Replace(" '' '' ", "");
                                extractedRIGNAME = extractedRIGNAME.Replace("\"", "");
                                int space1 = extractedRIGNAME.IndexOf(" ");
                                Rigname = (extractedRIGNAME.Substring(0, space1));
                                Rigname = Rigname.TrimStart();
                                Rigname = Rigname.TrimEnd();
                                Rigname = Rigname.Trim();

                                /**  extracting last 24*/
                                int From4 = FullData.IndexOf("Last 24 hr operations") + "Last 24 hr operations".Length;
                                int To4 = FullData.IndexOf("Next 24 hr plan");
                                last24 = FullData.Substring(From4, To4 - From4);
                                last24 = last24.Replace(",", System.Environment.NewLine);
                                last24 = last24.TrimStart();
                                last24 = last24.TrimEnd();
                                /**  extracting last 24*/

                                /**  extracting DaysSince*/
                                int From5 = FullData.IndexOf("Days Since Spud/Comm (Date)") + "Days Since Spud/Comm (Date)".Length;
                                int To5 = FullData.IndexOf("Formation tops");
                                DaysSince = FullData.Substring(From5, To5 - From5);
                                DaysSince = DaysSince.TrimStart();
                                DaysSince = DaysSince.TrimEnd();
                                DaysSince = Regex.Replace(DaysSince, @"\s", "");
                                /**  extracting DaysSince*/


                                DateTime enter_date;

                                if (extractedDATEONLY.Contains(string.Empty))
                                {
                                    string strDate = extractedDATEONLY;
                                    string[] dateString = strDate.Split('/');

                                    enter_date = new DateTime(1900, 01, 01);


                                }
                                else
                                {
                                    string strDate = extractedDATEONLY;
                                    string[] dateString = strDate.Split('/');
                                    enter_date = Convert.ToDateTime(dateString[1] + "/" + dateString[0] + "/" + dateString[2]);

                                }

                                /*start extrat RIGNA**//***/
                                string FinalString0 = "";

                                int From6 = FullData.LastIndexOf("Wellbores:") + "Wellbores:".Length;
                                int To6 = FullData.LastIndexOf("Foreman(s)");
                                FinalString0 = FullData.Substring(From6, To6 - From6);
                                List<string> EXtractRIGNAMELIST = FinalString0.Split('\n').ToList();
                                EXtractRIGNAMELIST.RemoveAt(0);
                                foreach (var word in EXtractRIGNAMELIST)
                                {
                                    extratcRIGNameNEWONE = word.ToString();
                                }
                                extratcRIGNameNEWONE = extratcRIGNameNEWONE.TrimStart();
                                extratcRIGNameNEWONE = extratcRIGNameNEWONE.TrimEnd();
                                extratcRIGNameNEWONE = extratcRIGNameNEWONE.Trim();


                                /*end extrat RIGNA**//***/
                                //string strDate = extractedDATEONLY;

                                //string[] dateString = strDate.Split('/');

                                //DateTime enter_date = Convert.ToDateTime(dateString[0] + "/" + dateString[1] + "/" + dateString[2]);











                                ///** start insert rig info non contain mudtreatment */
                                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                                {
                                    //SqlCommand check_RIG1 = new SqlCommand("SELECT Count(*) FROM [Rigs] WHERE ([Rigname] = @C1 and RIGN = NULL)", con);
                                    //con.Open();
                                    //check_RIG1.Parameters.AddWithValue("@C1", Rigname);
                                    //int RIGExist1 = (int)check_RIG1.ExecuteScalar();
                                    //if (RIGExist1 == 0)
                                    //{
                                    //    SqlCommand cmd = new SqlCommand(" UPDATE [Rigs] SET [RIGN] = @C1 WHERE [Rigname] = @C2 ", con);
                                    //    cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                    //    cmd.Parameters["@C1"].Value = extratcRIGNameNEWONE.ToString();
                                    //    cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.VarChar));
                                    //    cmd.Parameters["@C2"].Value = Rigname.ToString();
                                    //    cmd.ExecuteNonQuery();
                                    //    con.Close();
                                    //}




                                    SqlCommand check_RIG = new SqlCommand("SELECT Count(*) FROM [Rigs] WHERE ([Rigname] = @C1)", con);
                                    con.Open();
                                    check_RIG.Parameters.AddWithValue("@C1", Rigname);
                                    int RIGExist = (int)check_RIG.ExecuteScalar();
                                    if (RIGExist == 0)
                                    {
                                        using (SqlCommand cmd = new SqlCommand("INSERT INTO Rigs (Rigname,Depth,last24,DaysSince,Contain,RIGN) VALUES (@C1,@C2,@C3,@C4,0,@C5)", con))
                                        {
                                            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                            cmd.Parameters["@C1"].Value = Rigname.ToString();
                                            cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.VarChar));
                                            cmd.Parameters["@C2"].Value = Depth.ToString();
                                            cmd.Parameters.Add(new SqlParameter("@C3", SqlDbType.VarChar));
                                            cmd.Parameters["@C3"].Value = last24.ToString();
                                            cmd.Parameters.Add(new SqlParameter("@C4", SqlDbType.VarChar));
                                            cmd.Parameters["@C4"].Value = DaysSince.ToString();
                                            cmd.Parameters.Add(new SqlParameter("@C5", SqlDbType.VarChar));
                                            cmd.Parameters["@C5"].Value = extratcRIGNameNEWONE.ToString();


                                            cmd.ExecuteNonQuery();
                                            using (SqlCommand cmd1 = new SqlCommand("SELECT (RigID)  FROM  Rigs WHERE Rigname=@C1", con))
                                            {
                                                cmd1.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                                cmd1.Parameters["@C1"].Value = Rigname.ToString();

                                                RigID = (int)cmd1.ExecuteScalar();
                                            }

                                        }
                                    }
                                    else
                                    {
                                        using (SqlCommand cmd = new SqlCommand("SELECT (RigID)  FROM  Rigs WHERE Rigname=@C1", con))
                                        {
                                            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.VarChar));
                                            cmd.Parameters["@C1"].Value = Rigname.ToString();
                                            RigID = (int)cmd.ExecuteScalar();
                                        }

                                    }
                                }
                                ///** end insert rig info non contain mudtreatment */  
                                using (SqlConnection con = new SqlConnection("Data Source=192.168.1.105;Initial Catalog=OILREPORT2;Persist Security Info=True;User ID=sa;password=Ram72763@"))
                                {
                                    con.Open();
                                    SqlCommand check_FILE = new SqlCommand(" SELECT Count(*) from files,Rigs where files.RigID=Rigs.RigID and" +
                                          "  files.RigID=@C3 and date=@C4 ", con);

                                    check_FILE.Parameters.AddWithValue("@C3", RigID.ToString());
                                    check_FILE.Parameters.AddWithValue("@C4", enter_date.ToString());
                                    int FILEExist = (int)check_FILE.ExecuteScalar();
                                    if (FILEExist == 0)
                                    {

                                        using (SqlCommand cmd = new SqlCommand("INSERT INTO Files(RigID,Date,Contain) VALUES (@C1,@C2,'0')", con))
                                        {
                                            cmd.Parameters.Add(new SqlParameter("@C1", SqlDbType.Int));
                                            cmd.Parameters["@C1"].Value = RigID;

                                            cmd.Parameters.Add(new SqlParameter("@C2", SqlDbType.Date));
                                            cmd.Parameters["@C2"].Value = enter_date;
                                            cmd.ExecuteNonQuery();
                                            //   MessageBox.Show(" inserted rig  " + RigID.ToString());
                                        }

                                    }
                                    else
                                    {
                                    }
                                }


                            }
                        }

                        File.Delete(workbook.FileName);

                    }
                }

            }
            else
            {
                //DialogResult res1 = MessageBox.Show("Are you sure you want to Delete", "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                //if (res1 == DialogResult.Cancel)
                //{
                //    MessageBox.Show("You have clicked Cancel Button");
                //    //Some task…
                //}

            }
            MessageBox.Show("The data has been exported successfully", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //BindTotal();
            BindGV();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
