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

namespace Oil_and_Gas_Software
{
    public partial class Form1 : MetroForm
    {
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        OpenFileDialog opf = new OpenFileDialog();
        DataTable dt = new DataTable();
        SQLCONNECTION SQLCONN = new SQLCONNECTION();



        public Form1()
        {
            InitializeComponent();

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
         //   BindGV();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            BindGV();
        }
        Rectangle BoundRect;
        Rectangle OldRect = Rectangle.Empty;
     


        private void EnableMouse()
        {
            Cursor.Clip = OldRect;
            Cursor.Show();
           // Application.RemoveMessageFilter(this.panel1);
        }
        public bool PreFilterMessage(ref Message m)
        {
            if (m.Msg == 0x201 || m.Msg == 0x202 || m.Msg == 0x203) return true;
            if (m.Msg == 0x204 || m.Msg == 0x205 || m.Msg == 0x206) return true;
            return false;
        }
        private void DisableMouse()
        {
            OldRect = Cursor.Clip;
            // Arbitrary location.
            BoundRect = new Rectangle(50, 50, 1, 1);
            Cursor.Clip = BoundRect;
            Cursor.Hide();
           // Application.AddMessageFilter(this);
        }


        private void BrowseBtn_Click(object sender, EventArgs e)
        {
          

            if (opf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ///**unzip process*/
                string zipFilePath = opf.FileName;
                string extractionPath = opf.FileName + " ";
                extractionPath = extractionPath.Replace(".zip", "");
                extractionPath = extractionPath.Trim();
                extractionPath = extractionPath.TrimStart();
                extractionPath = extractionPath.TrimEnd();

                ZipFile.ExtractToDirectory(zipFilePath, extractionPath);
                ///**unzip process*/
                string sourceDirectory =extractionPath;
                txtfilepath.Text = extractionPath;
                // this.WindowState = FormWindowState.Minimized;
               

                SQLCONN.OpenConection();

        //        var watch3 = System.Diagnostics.Stopwatch.StartNew();


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


                        string extractedfullDATE = "";
                        string extractedDATEONLY = "";
                        string extractedWELLNAME = "";
                        string extractedRIGDATA = "";
                        string extractedRIGDepth = "";
                        string last24 = "";
                        string DaysSince = "";
                        string extratcRIGNO = "";
                        var Depth = "";
                        var Wellname = "";
                        ReportID = 0;
                        RigID = 0;
                        WellID = 0;

                        //* count convert and extract for mcontain mud */
                      
                        //**/
                       var workbook = new Workbook(currentFile);
                      
                       workbook.Save(currentFile + ".txt");

                        using (var sr1 = new StreamReader(currentFile, true))
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
                                    extractedWELLNAME = FullData.Substring(From2, To2 - From2);
                                    // remove between bractise /** to 
                                    extractedWELLNAME = Regex.Replace(extractedWELLNAME, @"\([^)]*\)", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace(")", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace(";", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace(",", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace(" '' '' ", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace("\"", "");
                                    int space1 = extractedWELLNAME.IndexOf(" ");
                                    Wellname = (extractedWELLNAME.Substring(0, space1));
                                    Wellname = Wellname.TrimStart();
                                    Wellname = Wellname.TrimEnd();
                                    Wellname = Wellname.Trim();

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

                                    ///*start extrat RIGNA**//***/
                                    string FinalString0 = "";
                                    int From6 = FullData.LastIndexOf("Wellbores:") + "Wellbores:".Length;
                                    int To6 = FullData.LastIndexOf("Foreman(s)");


                                    FinalString0 = FullData.Substring(From6, To6 - From6);
                                    List<string> EXtractRIGNAMELIST = FinalString0.Split('\n').ToList();
                                    EXtractRIGNAMELIST.RemoveAt(0);
                                    foreach (var word in EXtractRIGNAMELIST)
                                    {
                                        extratcRIGNO = word.ToString();
                                    }
                                    extratcRIGNO = extratcRIGNO.TrimStart();
                                    extratcRIGNO = extratcRIGNO.TrimEnd();
                                    extratcRIGNO = extratcRIGNO.Trim();

                               


                                    /*end extrat RIGNA**//***/





                                    /** start insert rig info contain mudtreatment*/

                                    /** opt*/

                                    SqlParameter paramextratcRIGNOName = new SqlParameter("@C1", SqlDbType.NVarChar);
                                    paramextratcRIGNOName.Value = extratcRIGNO;


                                    SqlDataReader dr = SQLCONN.DataReader("SELECT RigID FROM [Rigs] WHERE [Rigname]= '" + extratcRIGNO + "'");
                                    dr.Read();
                                    if (dr.HasRows)
                                    {
                                        dr.Dispose();
                                        dr.Close();
                                        dr = SQLCONN.DataReader("SELECT (RigID)  FROM  Rigs WHERE rigname= '" + extratcRIGNO + "'");
                                        dr.Read();
                                        RigID = int.Parse(dr["RigID"].ToString());
                                        dr.Dispose();
                                        dr.Close();

                                    }
                                    else
                                    {
                                        dr.Dispose();
                                        dr.Close();
                                        SQLCONN.ExecuteQueries("INSERT INTO Rigs (rigname) VALUES (@C1)", paramextratcRIGNOName);
                                        dr = SQLCONN.DataReader("SELECT (RigID)  FROM  Rigs WHERE rigname= '" + extratcRIGNO + "'");
                                        dr.Read();
                                        RigID = int.Parse(dr["RigID"].ToString());
                                        dr.Dispose();
                                        dr.Close();

                                    }

                                    /**opt*/
                                    ///** end insert rig info contain mudtreatment */

                                    /** start insert well info contain mudtreatment*/

                                    ///**opt*/
                                    SqlParameter paramWellname = new SqlParameter("@C2", SqlDbType.NVarChar);
                                    paramWellname.Value = Wellname;



                                    dr = SQLCONN.DataReader("SELECT Wellid FROM [wells] WHERE [wellname]= '" + Wellname + "'");
                                    dr.Read();
                                    if (dr.HasRows)
                                    {
                                        dr.Dispose();
                                        dr.Close();
                                        dr = SQLCONN.DataReader("SELECT (Wellid) FROM [wells] WHERE [wellname]= '" + Wellname + "'");
                                        dr.Read();
                                        WellID = int.Parse(dr["Wellid"].ToString());
                                        dr.Dispose();
                                        dr.Close();

                                    }
                                    else
                                    {
                                        dr.Dispose();
                                        dr.Close();
                                        SQLCONN.ExecuteQueries("INSERT INTO wells (wellname) VALUES (@C2)", paramWellname);
                                        dr = SQLCONN.DataReader("SELECT (Wellid)  FROM  wells WHERE wellname = '" + Wellname + "'");
                                        dr.Read();
                                        WellID = int.Parse(dr["Wellid"].ToString());
                                        dr.Dispose();
                                        dr.Close();
                                    }
                                    /**opt*/
                                    ///** end insert well info contain mudtreatment */


                                    /** start section (Mud Treatment )  */
                                    List<string> FinalString3LIST = extractedRIGDATA.Split('\n').ToList();
                                    foreach (var word in FinalString3LIST)
                                    {
                                        //  MessageBox.Show(word);
                                        MaterialID = 0;
                                        string strDate = extractedDATEONLY;
                                        string[] dateString = strDate.Split('/');
                                          DateTime enter_date = Convert.ToDateTime(dateString[0] + "/" + dateString[1] + "/" + dateString[2]);
                                          enter_date.ToShortDateString();

                                          enter_date.ToString("yyyy-MM-dd");

                                        int qous = word.IndexOf("(");
                                        int space = word.LastIndexOf(" ");
                                        var MValue = (word.Substring(space, qous - space));
                                        MValue = MValue.Replace("(", " ");
                                        var keyword = (word.Substring(word.IndexOf(word), space));
                                        keyword = keyword.TrimStart();
                                        keyword = keyword.TrimEnd();
                                        keyword = keyword.Trim();

                                        /*extract value between  brackets */
                                        int start = word.LastIndexOf("(") + 1;
                                        int end = word.IndexOf(")", start);
                                        string brackets = word.Substring(start, end - start);
                                        Regex re = new Regex("([0-9]+)([A-Z]+)");
                                        Match result2 = re.Match(brackets);
                                        string PackingQTY = result2.Groups[1].Value;
                                        string UnitName = result2.Groups[2].Value;
                                        if ( UnitName.Length <= 0)
                                        {
                                            int From = word.IndexOf("(") + "(".Length;
                                            int To = word.IndexOf(")");
                                            UnitName = word.Substring(From, To - From);
                                        }

                                        /*extract value between  materials */
                                        /** extract material**/


                                        /** opt */
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
                                        SqlParameter paramDate = new SqlParameter("@C4", SqlDbType.Date);
                                        paramDate.Value = enter_date;
                                        SqlParameter paramDepth = new SqlParameter("@C5", SqlDbType.NVarChar);
                                        paramDepth.Value = Depth;
                                        SqlParameter paramDAYSSINCE = new SqlParameter("@C6", SqlDbType.NVarChar);
                                        paramDAYSSINCE.Value = DaysSince;
                                        SqlParameter paramLAST24 = new SqlParameter("@C7", SqlDbType.NVarChar);
                                        paramLAST24.Value = last24;

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
                                                "  REPORTS.WELLID='" + WellID + "' and date='" + enter_date + "'and REPORTS.RIGID= '" + RigID + "'");
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

                                        dr = SQLCONN.DataReader("SELECT REPORTID FROM  Reports WHERE wellid='" + WellID + "'and REPORTS.RIGID= '" + RigID + "'and date= '" + enter_date + "'");
                                        dr.Read();
                                        ReportID = int.Parse(dr["REPORTID"].ToString());
                                        SqlParameter paramReportID = new SqlParameter("@C8", SqlDbType.Int);
                                        paramReportID.Value = ReportID;

                                        if (dr.HasRows)
                                        {
                                            dr.Dispose();
                                            dr.Close();
                                            dr = SQLCONN.DataReader("SELECT MUD_TRATMENT.REPORTID from REPORTS,MUD_TRATMENT where MUD_TRATMENT.REPORTID ='" + ReportID + "' and MUD_TRATMENT.MATID = '" + MaterialID + "' AND MUD_TRATMENT.QTY = '" + MValue + "' and REPORTS.date = '" + enter_date + "' and UnitName ='" + UnitName + "' and PackingQTY='" + PackingQTY + "'  and Reports.reportid = MUD_TRATMENT.REPORTID");
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

                                        /**opt*/

                                        /** check Dublicate data in mudtreatment and insert new data in mudtreatment  **/
                                    }
                                    /** end  section (Mud Treatment )  */
                                    //watch.Stop();
                                    //label4.Text = watch.Elapsed.ToString();




                                }
                                else
                                {

                                    /**calcltate for extract non contain data*/
                                    /**calcltate for extract non contain data*/
                                  
                                    //if (!watch2.IsRunning) // checks if it is not running
                                    //    watch2.Start(); // Start the counter from where it stopped

                                    //for (int j = 0; j < 1000; j++)
                                    //{
                                    //    Console.Write(j);
                                    //}

                                    /** for check files with out mud treayment : select * from reports where dayssince = "0.0(@)"*/
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
                                    extractedDATEONLY = extractedDATEONLY.TrimStart();
                                    extractedDATEONLY = extractedDATEONLY.TrimEnd();
                                    extractedDATEONLY = extractedDATEONLY.Trim();

                                    int From2 = FullData.IndexOf("Well No (Type) :") + "Well No (Type) :".Length;
                                    int To2 = FullData.IndexOf("Charge #");
                                    extractedWELLNAME = FullData.Substring(From2, To2 - From2);
                                    // remove between bractise /** to 
                                    extractedWELLNAME = Regex.Replace(extractedWELLNAME, @"\([^)]*\)", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace(")", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace(";", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace(",", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace(" '' '' ", "");
                                    extractedWELLNAME = extractedWELLNAME.Replace("\"", "");
                                    int space1 = extractedWELLNAME.IndexOf(" ");
                                    Wellname = (extractedWELLNAME.Substring(0, space1));
                                    Wellname = Wellname.TrimStart();
                                    Wellname = Wellname.TrimEnd();
                                    Wellname = Wellname.Trim();

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

                                    if (extractedDATEONLY.Length <= 0)
                                    {
                                        string strDate = extractedDATEONLY;
                                        string[] dateString = strDate.Split('/');

                                        enter_date = new DateTime(1900, 01, 01);


                                    }
                                    else
                                    {
                                        string strDate = extractedDATEONLY;
                                        string[] dateString = strDate.Split('/');
                                        enter_date = Convert.ToDateTime(dateString[0] + "/" + dateString[1] + "/" + dateString[2]);
                                        enter_date.ToShortDateString();
                                        enter_date.ToString("yyyy-MM-dd");

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
                                        extratcRIGNO = word.ToString();
                                    }
                                    extratcRIGNO = extratcRIGNO.TrimStart();
                                    extratcRIGNO = extratcRIGNO.TrimEnd();
                                    extratcRIGNO = extratcRIGNO.Trim();

                                    //watch2.Stop();
                                
                                    // label5.Text = watch2.Elapsed.ToString();

                                    /*end extrat RIGNA**//***/
                                    /** calcute insert non contain */
                                    /**calcltate for extract non contain data*/
                                
                                    //if (!watch2.IsRunning) // checks if it is not running
                                    //watch2.Start(); // Start the counter from where it stopped

                                    //for (int j = 0; j < 1000; j++)
                                    //{
                                    //    Console.Write(j);

                                    //}

                                    ///** start insert rig info non contain mudtreatment */

                                    ///**opt*/
                                    SqlParameter paramextratcRIGNO1 = new SqlParameter("@C1", SqlDbType.NVarChar);
                                    paramextratcRIGNO1.Value = extratcRIGNO;
                                    SqlDataReader dr = SQLCONN.DataReader("SELECT RigID FROM [Rigs] WHERE [Rigname]= '" + extratcRIGNO + "'");
                                    dr.Read();
                                    if (dr.HasRows)
                                    {
                                        dr.Dispose();
                                        dr.Close();
                                        dr = SQLCONN.DataReader("SELECT (RigID)  FROM  Rigs WHERE rigname= '" + extratcRIGNO + "'");
                                        dr.Read();
                                        RigID = int.Parse(dr["RigID"].ToString());
                                        dr.Dispose();
                                        dr.Close();

                                    }
                                    else
                                    {
                                        dr.Dispose();
                                        dr.Close();
                                        SQLCONN.ExecuteQueries("INSERT INTO Rigs (rigname) VALUES (@C1)", paramextratcRIGNO1);
                                        dr = SQLCONN.DataReader("SELECT (RigID)  FROM  Rigs WHERE rigname= '" + extratcRIGNO + "'");
                                        dr.Read();
                                        RigID = int.Parse(dr["RigID"].ToString());
                                        dr.Dispose();
                                        dr.Close();

                                    }

                                    /**opt*/
                                    ///** end insert rig info non contain mudtreatment */  
                                    ///
                                    ///** start insert well info non contain mudtreatment */
                                    ///**opt*/
                                    SqlParameter paramWellname1 = new SqlParameter("@C2", SqlDbType.NVarChar);
                                    paramWellname1.Value = Wellname;

                                    dr = SQLCONN.DataReader("SELECT Wellid FROM [wells] WHERE [wellname]= '" + Wellname + "'");
                                    dr.Read();
                                    if (dr.HasRows)
                                    {
                                        dr.Dispose();
                                        dr.Close();
                                        dr = SQLCONN.DataReader("SELECT (Wellid) FROM [wells] WHERE [wellname]= '" + Wellname + "'");
                                        dr.Read();
                                        WellID = int.Parse(dr["Wellid"].ToString());
                                        dr.Dispose();
                                        dr.Close();

                                    }
                                    else
                                    {
                                        dr.Dispose();
                                        dr.Close();
                                        SQLCONN.ExecuteQueries("INSERT INTO wells (wellname) VALUES (@C2)", paramWellname1);
                                        dr = SQLCONN.DataReader("SELECT (Wellid)  FROM  wells WHERE wellname = '" + Wellname + "'");
                                        dr.Read();
                                        WellID = int.Parse(dr["Wellid"].ToString());
                                        dr.Dispose();
                                        dr.Close();
                                    }

                                    /**opt*/
                                    ///** end insert well info non contain mudtreatment */  
                                    /** insert report non contain mudtreatment**/
                                    /** opt */
                                    SqlParameter paramextratcRIGID = new SqlParameter("@C11", SqlDbType.Int);
                                    paramextratcRIGID.Value = RigID;
                                    SqlParameter paramWellID = new SqlParameter("@C22", SqlDbType.Int);
                                    paramWellID.Value = WellID;
                                    SqlParameter paramDate = new SqlParameter("@C4", SqlDbType.Date);
                                    paramDate.Value = enter_date;
                                    SqlParameter paramDepth = new SqlParameter("@C5", SqlDbType.NVarChar);
                                    paramDepth.Value = Depth;
                                    SqlParameter paramDAYSSINCE = new SqlParameter("@C6", SqlDbType.NVarChar);
                                    paramDAYSSINCE.Value = DaysSince;
                                    SqlParameter paramLAST24 = new SqlParameter("@C7", SqlDbType.NVarChar);
                                    paramLAST24.Value = last24;




                                    dr = SQLCONN.DataReader("SELECT REPORTS.WELLID,REPORTS.RIGID,date from REPORTS,RIGS,WELLS where REPORTS.RIGID=RIGS.RIGID and  REPORTS.WELLID=WELLS.WELLID AND " +
                                            "  REPORTS.WELLID='" + WellID + "' and date='" + enter_date + "'and REPORTS.RIGID= '" + RigID + "'");
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

                                    /** end report non contain mudtreatment**/









                                //    watch2.Stop();


                                //    label6.Text = watch2.Elapsed.ToString();

                                }
                              


                        }

                            File.Delete(workbook.FileName);

                        }
                        Application.DoEvents();

                    }



                    PopupNotifier popup = new PopupNotifier();
                    popup.TitleText = "Oil and Gas Software";
                    popup.ContentText = "The data has been exported successfully";
                    this.WindowState = FormWindowState.Maximized;



                    //  EnableMouse();



                    popup.Popup();// show

                   MessageBox.Show("The data has been exported successfully", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);





                }
                catch (Exception ex)
                {
                    this.WindowState = FormWindowState.Maximized;

                    MessageBox.Show(ex.Message);



                }
                //BindTotal();
                SQLCONN.CloseConnection();

                //watch3.Stop();
                //label13.Text = watch3.Elapsed.ToString();
            }
            else
            {

            }


          //  var watch4 = System.Diagnostics.Stopwatch.StartNew();

            BindGV();

         //   watch4.Stop();
         //   label5.Text = watch4.Elapsed.ToString();



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
            //    label1.Visible = label2.Visible = true;
            //progressBar1.Value = e.ProgressPercentage;
            //label1.Text = e.ProgressPercentage.ToString() + "%";
            //if (label1.Text == "100%")
            //{
            //    label2.Text = "The data has been exported successfully";
            //}

         
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
    }
}