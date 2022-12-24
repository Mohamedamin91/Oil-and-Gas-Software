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
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Tulpep.NotificationWindow;
using Timer = System.Windows.Forms.Timer;
using System.Timers;
using HtmlAgilityPack;

namespace Oil_and_Gas_Software
{
    public partial class Form1 : MetroForm
    {

        OpenFileDialog opf = new OpenFileDialog();
        DataTable dt = new DataTable();
        SQLCONNECTION SQLCONN = new SQLCONNECTION();
        public String lastRunDate = "";





        public Form1()
        {
           
             InitializeComponent();
            // Task.Delay(new TimeSpan(0, 0,10)).ContinueWith(o => { BrowseBtn.PerformClick(); });
            //     Task.Delay(new TimeSpan(15, 00,00)).ContinueWith(o => { BrowseBtn.PerformClick(); });
            timer1.Interval = 1000;
            timer1.Enabled = true;




        }
        public void BindGV()
        {
            dt.Rows.Clear();
            SQLCONN.OpenConection();
            dataGridView1.Visible = true;
            dataGridView1.DataSource = SQLCONN.ShowDataInGridViewORCombobox(" select rigs.Rigname 'Rig',WELLS.Wellname 'Well No',MATERIALS.MATName 'Materials', MUD_TRATMENT.QTY'QTY'," +
                    "MUD_TRATMENT.PackingQTY'PQTY',MUD_TRATMENT.UnitName 'Unit',reports.Date 'Date' from RIGS,WELLS,REPORTS,MUD_TRATMENT,MATERIALS" +
                    " where  REPORTS.RIGID = rigs.RIGID and reports.WELLID = WELLS.WELLID  and  MUD_TRATMENT .MATID = MATERIALS.MATID and  MUD_TRATMENT .REPORTID = REPORTS.REPORTID order by [Well No] ");
            SQLCONN.CloseConnection();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Interval = 1000;
            timer1.Enabled = true;
            //   BindGV();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            BindGV();
        }



        private async void BrowseBtn_Click(object sender, EventArgs e)
        {

          
            var allFirstFiles
                   // = Directory.EnumerateFiles("D:\\Aramco Reports", "*", SearchOption.AllDirectories);
                     = Directory.EnumerateFiles("E:\\Main Test", "*", SearchOption.AllDirectories);
            foreach (string currentFirstFile in allFirstFiles)
            {

                /** create new file name to comapre with date in database */
                currentFirstFile.Trim();
                currentFirstFile.TrimEnd();
                currentFirstFile.TrimStart();
                string newfilename = currentFirstFile;
                newfilename = newfilename.Replace("MR", "").ToString();
                newfilename = newfilename.Replace(".zip", "").ToString();
               // newfilename = newfilename.Replace("D:\\Aramco Reports\\", "").ToString();
                newfilename = newfilename.Replace("E:\\Main Test\\", "").ToString();
                newfilename.Trim();
                newfilename.TrimEnd();
                newfilename.TrimStart();

                SQLCONN.OpenConection();
                SqlDataReader drnew = SQLCONN.DataReader("select * from reports where date = '" + newfilename + "' ");
                if (drnew.HasRows)
                {

                }
                else
                {

                    ///**unzip process*/
                    string zipFilePath = currentFirstFile;
                   // string extractionPath = "D:\\Extracted" + ".zip";
                   string extractionPath = "E:\\Extracted" + ".zip";
                    extractionPath = extractionPath.Replace(".zip", "");
                    extractionPath = extractionPath.Trim();
                    extractionPath = extractionPath.TrimStart();
                    extractionPath = extractionPath.TrimEnd();
                    ZipFile.ExtractToDirectory(zipFilePath, extractionPath);
                    string sourceDirectory = extractionPath;
                    txtfilepath.Text = extractionPath;
                    SQLCONN.OpenConection();
                    try
                    {

                        var allFiles
                          = Directory.EnumerateFiles(sourceDirectory, "*", SearchOption.AllDirectories);
                        //  DisableMouse();

                        foreach (string currentFile in allFiles)
                        {


                            string fileName = currentFile.Substring(sourceDirectory.Length + 1);

                            int MaterialID;
                            int RigID;
                            int WellID;
                            int ReportID;

                            DirectoryInfo dir = new DirectoryInfo(sourceDirectory);


                            ReportID = 0;
                            RigID = 0;
                            WellID = 0;

                            //* count convert and extract for mcontain mud */

                            //**/
                            HtmlWeb web = new HtmlWeb();
                          //  MessageBox.Show(currentFile);
                            if (currentFile.Contains("index") ==true  )
                            {
                            }
                            else
                            {
                                var doc = await Task.Factory.StartNew(() => web.Load(currentFile.ToString()));

                                var DateNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"AutoNumber1\"]/thead/tr[1]/td/table/tbody/tr[1]/td[2]/table/tbody/tr[3]/td");
                                var Wellnode = doc.DocumentNode.SelectSingleNode("//*[@id=\"AutoNumber1\"]/thead/tr[1]/td/table/tbody/tr[1]/td[3]/table[1]/tbody/tr[2]/td");
                                var Rignode = doc.DocumentNode.SelectSingleNode("//*[@id=\"AutoNumber1\"]/thead/tr[1]/td/table/tbody/tr[1]/td[4]/table/tbody/tr[2]/td");
                                var Depthnode = doc.DocumentNode.SelectSingleNode("//*[@id=\"AutoNumber1\"]/thead/tr[2]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td");
                                var Last24node = doc.DocumentNode.SelectSingleNode("//*[@id=\"AutoNumber1\"]/thead/tr[2]/td/table/tbody/tr[1]/td[2]/table/tbody/tr[2]/td");
                                var DaysSincenode = doc.DocumentNode.SelectSingleNode("//*[@id=\"AutoNumber1\"]/thead/tr[2]/td/table/tbody/tr[1]/td[5]/table/tbody/tr[2]/td");
                                SQLCONN.OpenConection();
                                /** start insert rig info contain mudtreatment*/

                                /** opt*/

                                SqlParameter paramextratcRIGnode = new SqlParameter("@C1", SqlDbType.NVarChar);
                                paramextratcRIGnode.Value = Rignode.InnerText.ToString();
                                SqlDataReader dr = SQLCONN.DataReader("SELECT RigID FROM [Rigs] WHERE [Rigname]= '" + Rignode.InnerText.Trim().ToString() + "'");
                                dr.Read();
                                if (dr.HasRows)
                                {
                                    dr.Dispose();
                                    dr.Close();
                                    dr = SQLCONN.DataReader("SELECT (RigID)  FROM  Rigs WHERE rigname= '" + Rignode.InnerText.Trim().ToString() + "'");
                                    dr.Read();
                                    RigID = int.Parse(dr["RigID"].ToString());
                                    dr.Dispose();
                                    dr.Close();

                                }
                                else
                                {
                                    dr.Dispose();
                                    dr.Close();
                                    SQLCONN.ExecuteQueries("INSERT INTO Rigs (rigname) VALUES (@C1)", paramextratcRIGnode);
                                    dr = SQLCONN.DataReader("SELECT (RigID)  FROM  Rigs WHERE rigname= '" + Rignode.InnerText.Trim().ToString() + "'");
                                    dr.Read();
                                    RigID = int.Parse(dr["RigID"].ToString());
                                    dr.Dispose();
                                    dr.Close();

                                }

                                /**opt*/
                                ///** end insert rig info contain mudtreatment */

                                /** start insert well info contain mudtreatment*/

                                ///**opt*/
                                dr.Dispose();
                                dr.Close();
                                SqlParameter paramWellnode = new SqlParameter("@C2", SqlDbType.NVarChar);
                                paramWellnode.Value = Wellnode.InnerText.Trim().ToString();
                                dr = SQLCONN.DataReader("SELECT Wellid FROM [wells] WHERE [wellname]= '" + Wellnode.InnerText.Trim().ToString() + "'");
                                dr.Read();
                                if (dr.HasRows)
                                {
                                    dr.Dispose();
                                    dr.Close();
                                    dr = SQLCONN.DataReader("SELECT (Wellid) FROM [wells] WHERE [wellname]= '" + Wellnode.InnerText.Trim().ToString() + "'");
                                    dr.Read();
                                    WellID = int.Parse(dr["Wellid"].ToString());
                                    dr.Dispose();
                                    dr.Close();

                                }
                                else
                                {
                                    dr.Dispose();
                                    dr.Close();
                                    SQLCONN.ExecuteQueries("INSERT INTO wells (wellname) VALUES (@C2)", paramWellnode);
                                    dr = SQLCONN.DataReader("SELECT (Wellid)  FROM  wells WHERE wellname = '" + Wellnode.InnerText.Trim().ToString() + "'");
                                    dr.Read();
                                    WellID = int.Parse(dr["Wellid"].ToString());
                                    dr.Dispose();
                                    dr.Close();
                                }
                                /**opt*/
                                ///** end insert well info contain mudtreatment */



                                foreach (HtmlNode table in doc.DocumentNode.SelectNodes("//*[@id=\"AutoNumber1\"]/tbody/tr[2]/td/table/tbody/tr[2]"))
                                {
                                    foreach (HtmlNode cell in table.SelectNodes("td|th"))
                                    {
                                        if (cell.InnerText == string.Empty)
                                        {
                                            cell.RemoveAll();
                                        }
                                        else
                                        {
                                            string cellvalue = "";
                                            System.Text.RegularExpressions.Regex regex =
                                               new System.Text.RegularExpressions.Regex(@"(<br />|<br/>|</ br>|</br>|<br>|<span>|</span>)");
                                            cell.InnerHtml = regex.Replace(cell.InnerHtml, "   ");
                                            cell.InnerText.TrimStart().TrimEnd().ToString();
                                            cellvalue = cell.InnerText.ToString();
                                            cellvalue = cellvalue.Replace("MUD TREATMENT", string.Empty).ToString();
                                            List<string> cellvalueList = cellvalue.Split('\n').ToList();
                                            cellvalueList.RemoveAll(string.IsNullOrWhiteSpace);

                                            foreach (var word in cellvalueList)
                                            {
                                                string finalword = word.ToString();

                                                finalword = finalword.Trim();
                                                finalword = finalword.TrimStart();
                                                finalword = finalword.TrimEnd();

                                                int qous = finalword.IndexOf("(");
                                                int space = finalword.IndexOf("   ");
                                                string MValue = (finalword.Substring(space, qous - space));
                                                MValue = MValue.Replace("(", " ");
                                                var keyword = (finalword.Substring(finalword.IndexOf(finalword), space));
                                                keyword = keyword.TrimStart();
                                                keyword = keyword.TrimEnd();
                                                keyword = keyword.Trim();

                                                // extract value between brackets */
                                                int start = finalword.LastIndexOf("(") + 1;
                                                int end = finalword.IndexOf(")", start);
                                                string brackets = finalword.Substring(start, end - start);
                                                Regex re = new Regex("([0-9]+)([A-Z]+)");
                                                Match result2 = re.Match(brackets);
                                                string PackingQTY = result2.Groups[1].Value;
                                                string UnitName = result2.Groups[2].Value;
                                                if (UnitName.Length <= 0)
                                                {
                                                    int From = finalword.IndexOf("(") + "(".Length;
                                                    int To = finalword.IndexOf(")");
                                                    UnitName = finalword.Substring(From, To - From);
                                                }
                                                //*extract value between  brackets 
                                                SqlParameter paramkeyword = new SqlParameter("@C3", SqlDbType.NVarChar);
                                                paramkeyword.Value = keyword;

                                                dr = SQLCONN.DataReader("SELECT (MATID)  FROM  MATERIALS WHERE MATName= '" + keyword + "'");
                                                dr.Read();
                                                if (dr.HasRows)
                                                {
                                                    dr.Dispose();
                                                    dr.Close();
                                                    dr = SQLCONN.DataReader("SELECT (MATID)  FROM  MATERIALS WHERE MATName= '" + keyword + "'");
                                                    dr.Read();
                                                    MaterialID = int.Parse(dr["MATID"].ToString());
                                                    dr.Dispose();
                                                    dr.Close();

                                                }
                                                else
                                                {
                                                    dr.Dispose();
                                                    dr.Close();
                                                    SQLCONN.ExecuteQueries("INSERT INTO MATERIALS(MATName) VALUES (@C3)", paramkeyword);
                                                    dr = SQLCONN.DataReader("SELECT (MATID)  FROM  MATERIALS WHERE MATName=  '" + keyword + "'");
                                                    dr.Read();
                                                    MaterialID = int.Parse(dr["MATID"].ToString());
                                                    dr.Dispose();
                                                    dr.Close();
                                                }



                                                /** opt */
                                                /** extract material**/

                                                /** Update Category and sub cat**/

                                                /** opt */


                                                SQLCONN.ExecuteQueries("UPDATE [MUD_TRATMENT] SET [PackingQTY] = 1 WHERE [PackingQTY]= 0");
                                                SQLCONN.ExecuteQueries("UPDATE [MATERIALS] SET [CatID] = 1 /*,[SubID]=1 */ WHERE [MATName]= 'BARITE' or [MATName]= 'BA-NF' or [MATName]= 'BA-AM' or [MATName]= 'BA-AL'  or [MATName]= 'BA-AF' or [MATName]= 'BA-OR' or [MATName]= 'KI-BARITE' or [MATName]= 'BA-PERF' or [MATName]= 'BA-ESNAAD' or [MATName]= 'BA-BAR'  or[MATName]= 'BA-IBC' or [MATName]= 'BA-OM'  or[MATName]= 'BA-AGMED' or [MATName]= 'BA-MID'  or[MATName]= 'BA-POUR' or [MATName]= 'BA-ATAD'  or[MATName]= 'BA' ");
                                                SQLCONN.ExecuteQueries(" UPDATE [MATERIALS] SET [CatID] = 2 /*,[SubID]=2*/   WHERE [MATName]= 'BA-NF-BULK' or [MATName]= 'BA-BA-BULK' or [MATName]= 'BA-AR-BULK' or [MATName]= 'BA-DMF-BULK'or[MATName]= 'BA BULK'");
                                                SQLCONN.ExecuteQueries(" UPDATE [MATERIALS] SET [CatID] = 3 /*,[SubID]=3*/   WHERE [MATName]= 'CABR2' or [MATName]= 'CABR2-MI' or [MATName]= 'CABR2-HAL' or [MATName]= 'CABR2-TET'or [MATName]= 'CABR2-OS' or [MATName]= 'CABR2-JOR' or [MATName]= 'CABR2-AGR' or [MATName]= 'CABR2-SHA' or[MATName]= 'CABR2-JIA' or [MATName]= 'CABR2-WEI'  or[MATName]= 'CABR2-ALB'");
                                                SQLCONN.ExecuteQueries(" UPDATE [MATERIALS] SET [CatID] = 4 /*,[SubID]=4 */  WHERE [MATName]= 'CABR2-SHO' ");
                                                SQLCONN.ExecuteQueries(" UPDATE [MATERIALS] SET [CatID] = 5 /*,[SubID]=5  */ WHERE [MATName]= 'CACL2-80CC' or [MATName]= 'CACL2-80' or [MATName]= 'CACL2-80ME' or [MATName]= 'CACL2-80TET'or [MATName]= 'CACL2-80SO' or[MATName]= 'CACL2-80TAN' or[MATName]= 'CACL2-80DW' or[MATName]= 'CACL2-80GC' or[MATName]= 'CACL2-80QH' or [MATName]= 'CACL2-80WE'  or[MATName]= 'CACL2-80CH'or [MATName]= 'CACL2-80IN'  or[MATName]= 'CACL2-80TEEU' or [MATName]= 'CACL2-80LIA' ");
                                                SQLCONN.ExecuteQueries(" UPDATE [MATERIALS] SET [CatID] = 6 /*,[SubID]=6 */  WHERE [MATName]= 'CACL2-98' or [MATName]= 'CACL2-98-BB' or [MATName]= 'CACL2-98CH-BB' or [MATName]= 'CACL2-98BA'or [MATName]= 'CACL2-98BH-BB' or[MATName]= 'CACL2-98DW' or[MATName]= 'CACL2-98TA-BB' or[MATName]= 'CACL2-98JB-BB' or[MATName]= 'CACL2-98WE-BB' or [MATName]= 'CACL2-98JB'  or[MATName]= 'CACL2-98TCE'or [MATName]= 'CACL2-98TE-BB'  or[MATName]= 'CACL2-98IN-BB' ");
                                                SQLCONN.ExecuteQueries(" UPDATE [MATERIALS] SET [CatID] = 7 /*,[SubID]=7 */  WHERE [MATName]= 'LIG-OBM' or [MATName]= 'TNATHN'or [MATName]= 'LIGNITE' or [MATName]= 'CACL2-98BA' ");
                                                SQLCONN.ExecuteQueries(" UPDATE [MATERIALS] SET [CatID] =8  /*,[SubID]=8 */  WHERE [MATName]= 'MRBL-C-NF-BB' or [MATName]= 'MRBL-C-SEP' or [MATName]= 'MRBL-C-SEP-BB' or [MATName]= 'MRBL-C-NF'or [MATName]= 'MRBL-C' or[MATName]= 'MRBL-C-BH-BB' or[MATName]= 'MRBL-C-BH' ");
                                                SQLCONN.ExecuteQueries(" UPDATE [MATERIALS] SET [CatID] =9  /*,[SubID]=9*/   WHERE [MATName]= 'MRBL-F-BH' or [MATName]= 'MRBL-F-SEP' or [MATName]= 'MRBL-F-NF-BB' or [MATName]= 'MRBL-F-SEP-BB'or [MATName]= 'MRBL-F-NF' or[MATName]= 'MRBL-F' or[MATName]= 'MRBL-C-BH'or[MATName]= 'MRBL-F-BH-BB'or[MATName]= 'MRBL-F-AEC' or [MATName]= 'MRBL-F-AEC-BB'  or[MATName]= 'MRBL-F-TP-BB' ");
                                                SQLCONN.ExecuteQueries("UPDATE [MATERIALS] SET [CatID] =10 /*,[SubID]=10*/  WHERE [MATName]= 'MRBL-M-NF-BB' or [MATName]= 'MRBL-M-SEP' or [MATName]= 'MRBL-M-BH' or [MATName]= 'MRBL-MED-BB'or[MATName]= 'MRBL-M-MI' or[MATName]= 'MRBL-M-MI-BB' or[MATName]= 'MRBL-M-NF'or[MATName]= 'MRBL-M-SEP-BB'or[MATName]= 'MRBL-M'  ");
                                                SQLCONN.ExecuteQueries("UPDATE [MATERIALS] SET [CatID] =11 /*,[SubID]=11*/  WHERE [MATName]= 'GELTONE' or [MATName]= 'CARBOGEL-II' or [MATName]= 'OILGEL'or [MATName]= 'NAFGEL' ");
                                                SQLCONN.ExecuteQueries("UPDATE [MATERIALS] SET [CatID] =12 /*,[SubID]=12*/  WHERE [MATName]= 'DURATONE' or [MATName]= 'CARBOTROL' or [MATName]= 'VRSLIG'or [MATName]= 'NAFTROL' ");
                                                SQLCONN.ExecuteQueries("UPDATE [MATERIALS] SET [CatID] =13 /*,[SubID]=13 */ WHERE [MATName]= 'RESINEX II' or [MATName]= 'GMPRO-RX' or [MATName]= 'RENZI_SPNH' ");
                                                SQLCONN.ExecuteQueries("UPDATE [MATERIALS] SET [CatID] =14 /*,[SubID]=14*/  WHERE [MATName]= 'NACL' or [MATName]= 'NACL-SAR' or [MATName]= 'NACL-RY' or [MATName]= 'NACL-GC' or[MATName]= 'NACL-SEP' ");
                                                SQLCONN.ExecuteQueries("UPDATE [MATERIALS] SET [CatID] =15 /*,[SubID]=15*/  WHERE [MATName]= 'NACL-DEL' ");
                                                SQLCONN.ExecuteQueries("UPDATE [MATERIALS] SET [CatID] =16 /*,[SubID]= 107*/ WHERE [CatID]  IS NULL /*or [SubID] IS NULL*/ ");
                                                SQLCONN.ExecuteQueries("UPDATE MATERIALS SET MATERIALS.SubID = B.Subid FROM Category A, SUBCATEGORY B   WHERE  A.CatID = B.Catid and  a.CatID = MATERIALS.CatID ");



                                                /** Update Category and sub cat**/

                                                /** check Dublicate Reports and insert new reports  **/

                                                /** opt */
                                                SqlParameter paramkeywordID = new SqlParameter("@C33", SqlDbType.Int);
                                                paramkeywordID.Value = MaterialID;
                                                SqlParameter paramextratcRIGID = new SqlParameter("@C11", SqlDbType.Int);
                                                paramextratcRIGID.Value = RigID;
                                                SqlParameter paramWellID = new SqlParameter("@C22", SqlDbType.Int);
                                                paramWellID.Value = WellID;
                                                SqlParameter paramDate = new SqlParameter("@C4", SqlDbType.NVarChar);
                                                paramDate.Value = DateNode.InnerText.ToString();
                                                SqlParameter paramDepth = new SqlParameter("@C5", SqlDbType.NVarChar);
                                                paramDepth.Value = Depthnode.InnerText.Trim().ToString(); ;
                                                SqlParameter paramDAYSSINCE = new SqlParameter("@C6", SqlDbType.NVarChar);
                                                paramDAYSSINCE.Value = DaysSincenode.InnerText.Trim().ToString(); ;
                                                SqlParameter paramLAST24 = new SqlParameter("@C7", SqlDbType.NVarChar);
                                                paramLAST24.Value = Last24node.InnerText.ToString(); ;

                                                SqlParameter paramQty = new SqlParameter("@C9", SqlDbType.NVarChar);
                                                paramQty.Value = MValue;

                                                SqlParameter paramunit = new SqlParameter("@C10", SqlDbType.NVarChar);
                                                paramunit.Value = UnitName;
                                                SqlParameter paramPackingQTY = new SqlParameter("@C12", SqlDbType.NVarChar);
                                                paramPackingQTY.Value = PackingQTY;

                                                SqlParameter paramunit2 = new SqlParameter("@C13", SqlDbType.NVarChar);
                                                paramunit2.Value = UnitName;
                                                SqlParameter paramPackingQTY2 = new SqlParameter("@C14", SqlDbType.NVarChar);
                                                paramPackingQTY2.Value = PackingQTY;


                                                dr = SQLCONN.DataReader("SELECT REPORTS.WELLID,REPORTS.RIGID,date from REPORTS,RIGS,WELLS where REPORTS.RIGID=RIGS.RIGID and  REPORTS.WELLID=WELLS.WELLID AND " +
                                                        "  REPORTS.WELLID='" + WellID + "' and date='" + DateNode.InnerText.Trim().ToString() + "'and REPORTS.RIGID= '" + RigID + "'");
                                                dr.Read();
                                                if (dr.HasRows)
                                                {
                                                    dr.Dispose();
                                                    dr.Close();


                                                }
                                                else
                                                {
                                                    dr.Dispose();
                                                    dr.Close();
                                                    SQLCONN.ExecuteQueries("INSERT INTO REPORTS(Date,WellID,RIGID,DEPTH,DAYSSINCE,LAST24) VALUES (@C4,@C22,@C11,@C5,@C6,@C7)", paramDate, paramWellID, paramextratcRIGID, paramDepth, paramDAYSSINCE, paramLAST24);
                                                    dr.Dispose();
                                                    dr.Close();
                                                }

                                                /** opt */
                                                /** check Dublicate Reports and insert new reports  **/

                                                /** check Dublicate data in mudtreatment and insert new data in mudtreatment  **/

                                                /**opt*/

                                                dr = SQLCONN.DataReader("SELECT REPORTID FROM  Reports WHERE wellid='" + WellID + "'and REPORTS.RIGID= '" + RigID + "'and date= '" + DateNode.InnerText.Trim().ToString() + "'");
                                                dr.Read();
                                                ReportID = int.Parse(dr["REPORTID"].ToString());
                                                SqlParameter paramReportID = new SqlParameter("@C8", SqlDbType.Int);
                                                paramReportID.Value = ReportID;

                                                if (dr.HasRows)
                                                {
                                                    dr.Dispose();
                                                    dr.Close();
                                                    dr = SQLCONN.DataReader("SELECT MUD_TRATMENT.REPORTID from REPORTS,MUD_TRATMENT where MUD_TRATMENT.REPORTID ='" + ReportID + "' and MUD_TRATMENT.MATID = '" + MaterialID + "' AND MUD_TRATMENT.QTY = '" + MValue + "' and REPORTS.date = '" + DateNode.InnerText.ToString() + "' and UnitName ='" + UnitName + "' and PackingQTY='" + PackingQTY + "'  and Reports.reportid = MUD_TRATMENT.REPORTID");
                                                    dr.Read();
                                                    if (dr.HasRows)
                                                    {
                                                        dr.Dispose();
                                                        dr.Close();
                                                    }
                                                    else
                                                    {
                                                        dr.Dispose();
                                                        dr.Close();
                                                        SQLCONN.ExecuteQueries("insert into  MUD_TRATMENT ( REPORTID,MATID,QTY,UnitName,PackingQTY,UnitNewValue,PackingQTYNewValue) values (@C8,@C33,@C9,@C10,@C12,@C13,@C14)", paramReportID, paramkeywordID, paramQty, paramunit, paramPackingQTY, paramunit2, paramPackingQTY2);
                                                    }
                                                }
                                                else
                                                {
                                                    dr.Dispose();
                                                    dr.Close();
                                                }
                                            }

                                        }

                                    }
                                }

                                Application.DoEvents();

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.WindowState = FormWindowState.Maximized;

                        MessageBox.Show(ex.Message);

                    }
                    //BindTotal();
                    SQLCONN.CloseConnection();
                  
                    BindGV();
                    /** delete extracted files*/
                 //Directory.Delete("D:\\Extracted", true);
               Directory.Delete("E:\\Extracted", true);
             
                    //     var NewCreatedDirectory = Directory.CreateDirectory("D:\\Extracted");
             
          var NewCreatedDirectory = Directory.CreateDirectory("E:\\Extracted");
                    /** delete extracted files*/
                    ///**unzip process*/               

                }

                drnew.Dispose();
                drnew.Close();
                SQLCONN.CloseConnection();



             





            }
            PopupNotifier popup = new PopupNotifier();
            popup.TitleText = "Oil and Gas Software";
            popup.ContentText = "The data has been exported successfully";
            this.WindowState = FormWindowState.Maximized;
            popup.Popup();// show

          //  MessageBox.Show("The data has been exported successfully", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public static void DeleteDirectory(string target_dir)
        {
            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }

           // Directory.Delete(target_dir, false);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
           


        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Form2 frm2 = new Form2();
            this.Hide();
            frm2.Show();
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

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            Form4 frm4 = new Form4();
            this.Hide();
            frm4.Show();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            for(int i =1; i<=100; i++)
            {
                Thread.Sleep(10);
                backgroundWorker1.WorkerReportsProgress = true;
                backgroundWorker1.ReportProgress(i);

            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
           
         
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void txtfilepath_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
          
        }

        public  void Checkbtn_Click(object sender, EventArgs e)
        {
            var directory = new DirectoryInfo("E:\\test3");
            var myFile = directory.GetFiles()
                        .OrderByDescending(f => f.LastWriteTime)
                        .First();

            string newfilename = myFile.Name;
            newfilename = newfilename.Replace("MR", "").ToString();
            newfilename = newfilename.Replace(".zip", "").ToString();
            newfilename.Trim();
            newfilename.TrimEnd();
            newfilename.TrimStart();
            SQLCONN.OpenConection();
            SqlDataReader dr = SQLCONN.DataReader("select * from reports where date = '"+  newfilename  +"' ");
            if (dr.HasRows)
            {
               // MessageBox.Show("Found");
            }
            else
            {
              //  MessageBox.Show("Not Found");

            }
            dr.Dispose();
            dr.Close();
            SQLCONN.CloseConnection();



        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (lastRunDate != System.DateTime.Now.ToString("yyyy-MM-dd"))
            {
                String str = System.DateTime.Now.ToString("h:mm tt");

                if (str.Equals("8:00 AM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                   BrowseBtn.PerformClick();
                }
                if (str.Equals("9:00 AM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("10:00 AM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("11:00 AM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("12:00 AM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("13:00 PM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("14:00 PM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("15:00 PM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("16:00 PM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("17:00 PM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }
                if (str.Equals("18:00 PM"))
                {
                    lastRunDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    BrowseBtn.PerformClick();
                }

            }
        }
    }
}