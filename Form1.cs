using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using ExcelLibrary.SpreadSheet;
using Newtonsoft.Json.Linq;


namespace MedicSchedulerMileage
{
    public partial class FrmMain : Form
    {
        string FileName = "";
        string filePath = "";
        string strFullPath = "";
        public FrmMain()
        {
            InitializeComponent();
        }


        public string GetOfficeVersion()
        {
            string sVersion = string.Empty;
            Microsoft.Office.Interop.Excel.Application appVersion = new Microsoft.Office.Interop.Excel.Application();
            appVersion.Visible = false;
            string sBitness = ((Microsoft.Office.Interop.Excel.Application)(appVersion)).OperatingSystem;
            //switch (appVersion.Version.ToString())
            //{
            //    case "7.0":
            //        sVersion = "95";
            //        break;
            //    case "8.0":
            //        sVersion = "97";
            //        break;
            //    case "9.0":
            //        sVersion = "2000";
            //        break;
            //    case "10.0":
            //        sVersion = "2002";
            //        break;
            //    case "11.0":
            //        sVersion = "2003";
            //        break;
            //    case "12.0":
            //        sVersion = "2007";
            //        break;
            //    case "14.0":
            //        sVersion = "2010";
            //        break;
            //    default:
            //        sVersion = "Too Old!";
            //        break;
            //}
            // MessageBox.Show("MS office version: " + sVersion + " " + sBitness);
            return sBitness;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFDFile = new OpenFileDialog();

            if (OFDFile.ShowDialog() == DialogResult.OK)
            {
                FileName = OFDFile.FileName.Substring(OFDFile.FileName.LastIndexOf("\\") + 1);
                filePath = OFDFile.FileName.Replace(FileName, "");
                txtInputfile.Text = OFDFile.FileName.Trim();
                strFullPath = filePath + FileName;
            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            tbProgress.Visible = true;
            //tbProgress.Value = 0;
            ExtractData(strFullPath);
            //Thread objThread1 = new Thread(LoadProgress);
            //objThread1.Start();

            //Thread objThread = new Thread(ExtractData);
            //objThread.Start(strFullPath);
            //tbProgress.Visible = false;
            //tbProgress.Value = 100;
        }

        public void ExtractData(object objFileName)
        {
            string distance = "";
            string Tolls = "";
            int stepsCount = 0;
            string strFileName = objFileName as string;
            if (strFileName != "")
            {

                if (strFileName.ToLower().Contains(".xls") || strFileName.ToLower().Contains(".xlsx"))
                {
                    string connString;
                    if (strFileName.ToLower().Contains(".xlsx"))
                        connString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 12.0;HDR=Yes;IMEX=2\";";
                    else
                        connString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=2\";";


                    //string query = "Select * From [Sheet1$]";
                    OleDbConnection connObj = new OleDbConnection(connString);
                    try
                    {
                        connObj.Open();
                    }
                    catch (Exception ex)
                    {
                        if (connObj != null)
                            connObj.Close();

                        if (ex.Message.ToLower().Contains("provider is not registered"))
                        {
                            string sOfficeBitness = GetOfficeVersion();

                            if (sOfficeBitness.Contains("64"))
                            {
                                Process process = Process.Start(Path.Combine(Application.StartupPath, "AccessDatabaseEngine_x64.exe"));
                                if (process != null)
                                    process.WaitForExit();
                            }
                            else
                            {
                                Process process = Process.Start(Path.Combine(Application.StartupPath, "AccessDatabaseEngine.exe"));
                                if (process != null)
                                    process.WaitForExit();
                            }
                        }
                        else if (ex.Message.ToLower().Contains("80040154 class not registered"))
                        {
                            //MessageBox.Show("Please make sure that your IP address is authorized to use the API key.", "MileageTracker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //return;

                            AlertMessage objAlert = new AlertMessage();
                            objAlert.ShowDialog(this);
                            //this.Close();
                            return;

                        }
                        else
                            MessageBox.Show(ex.Message);

                        return;
                    }

                    DataTable dtObj = new DataTable();
                    DataTable dtObjClone = new DataTable();
                    int iCount = 0;

                    DataTable dtExcelSchema = connObj.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();

                    string query = "SELECT * FROM [" + SheetName + "] ";
                    OleDbCommand cmdObj = new OleDbCommand(query, connObj);
                    OleDbDataAdapter daObj = new OleDbDataAdapter(cmdObj);
                    DataTable dtInputData = new DataTable();
                    daObj.Fill(dtInputData);
                    connObj.Close();

                    DataColumn newCol = new DataColumn("Mileage", typeof(string));
                    dtInputData.Columns.Add(newCol);
                    DataColumn newCol2 = new DataColumn("Tolls", typeof(string));
                    dtInputData.Columns.Add(newCol2);

                    // dtObj = dtInputData.Clone();
                    dtObj = dtInputData.AsEnumerable().CopyToDataTable();
                    // dtObjClone = dtObj.AsEnumerable().CopyToDataTable();

                    string sFileName = Guid.NewGuid().ToString() + ".xls";
                    string originalpath = Application.StartupPath + "\\ExportFiles\\" + sFileName;

                    if (dtObj.Rows.Count > 0)
                    {

                        Workbook workbook = new Workbook();
                        Worksheet worksheet = new Worksheet("Export Data");

                        //Headers
                        for (int i = 0; i < dtObj.Columns.Count; i++)
                        {

                            worksheet.Cells[0, i] = new Cell(dtObj.Columns[i].ColumnName);
                        }

                        // Content.  
                        for (int i = 0; i < 29001; i++)
                        {
                            for (int j = 0; j < dtObj.Columns.Count; j++)
                            {
                                worksheet.Cells[i + 1, j] = new Cell(dtObj.Rows[i][j].ToString());
                            }
                        }


                        //worksheet.Cells.ColumnWidth[0, 1] = 3000;
                        workbook.Worksheets.Add(worksheet);
                        workbook.Save(originalpath);

                        Workbook book = Workbook.Load(originalpath);
                        Worksheet sheet = book.Worksheets[0];
                    }

                    //if (dtInputData.Rows.Count > 1000)
                    //{
                    //    int iLoopcount = (dtInputData.Rows.Count % 1000) != 0 ? ((dtInputData.Rows.Count / 1000) + 1) : (dtInputData.Rows.Count / 1000);

                    //    for (int iSplitCount = 0; iSplitCount < 3; iSplitCount++)
                    //    {
                    //        dtObj.Rows.Clear();
                    //        if (iCount == 0)
                    //            dtInputData.AsEnumerable().Take(1000).CopyToDataTable(dtObj, LoadOption.Upsert);
                    //        else
                    //            dtInputData.AsEnumerable().Skip(iCount * 1000).Take(1000).CopyToDataTable(dtObj, LoadOption.Upsert);

                    //        tbProgress.Minimum = 1;
                    //        tbProgress.Maximum = dtObj.Rows.Count;
                    //        tbProgress.Value = 1;
                    //        tbProgress.Step = 1;

                    //        if (dtObj.Rows.Count > 0)
                    //        {
                    //            for (int i = 0; i < dtObj.Rows.Count; i++)
                    //            {
                    //                bool flag = false;
                    //                if (dtObj.Rows[i]["puaddr"].ToString() != "" && dtObj.Rows[i]["daddr"].ToString() != "")
                    //                {
                    //                    string Origin = "";
                    //                    string Destination = "";
                    //                    string Status = "";

                    //                    // AIzaSyDrdmH-HSG4o7F_Dnn93jQzsHc60RW2wj4
                    //                    //to check in other IP comment below line
                    //                    string API_KEY = "AIzaSyDrdmH-HSG4o7F_Dnn93jQzsHc60RW2wj4"; //"AIzaSyBIAXvn_JfN4eK2_9SNctV7fi72EKBpDTs"; //

                    //                    Origin = dtObj.Rows[i]["puaddr"].ToString() + " " + dtObj.Rows[i]["pucity"].ToString() + " " + dtObj.Rows[i]["pust"].ToString() + " " + dtObj.Rows[i]["puzip"].ToString();

                    //                    Destination = dtObj.Rows[i]["daddr"].ToString() + " " + dtObj.Rows[i]["dcity"].ToString() + " " + dtObj.Rows[i]["dst"].ToString() + " " + dtObj.Rows[i]["dzip"].ToString();

                    //                    //to check in other IP comment below line
                    //                    string url = "https://maps.googleapis.com/maps/api/directions/json?origin=" + Origin.Replace("#", "") + "&destination=" + Destination.Replace("#", "") + "&sensor=false&key=" + API_KEY + "&userIp=" + UserIP;

                    //                    //to check in valid IP uncomment below line
                    //                    //  string url = "http://maps.googleapis.com/maps/api/directions/json?origin=" + Origin.Replace("#", "") + "&destination=" + Destination.Replace("#", "") + "&sensor=false";

                    //                    string requesturl = url;
                    //                    string content = fileGetJSON(requesturl);

                    //                    if (content.ToUpper() == "DENIED")
                    //                    {
                    //                        MessageBox.Show("Please make sure that your IP address is authorized to use the API key.", "MileageTracker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //                        return;
                    //                    }
                    //                    else if (content == "unable to connect to server")
                    //                    {
                    //                        dtObj.Rows[i]["Mileage"] = "unable to connect to server";
                    //                        if (flag)
                    //                            dtObj.Rows[i]["Tolls"] = "";
                    //                        else
                    //                            dtObj.Rows[i]["Tolls"] = "";
                    //                    }
                    //                    else
                    //                    {
                    //                        try
                    //                        {
                    //                            JObject _Jobj = JObject.Parse(content);
                    //                            Status = (string)_Jobj.SelectToken("status");
                    //                            tbProgress.PerformStep();
                    //                            if (Status == "OK")
                    //                            {
                    //                                distance = (string)_Jobj.SelectToken("routes[0].legs[0].distance.text");
                    //                                stepsCount = (int)_Jobj.SelectToken("routes[0].legs[0].steps").Count();

                    //                                for (int j = 0; j < stepsCount; j++)
                    //                                {
                    //                                    Tolls = (string)_Jobj.SelectToken("routes[0].legs[0].steps[" + j + "].html_instructions");
                    //                                    if (Tolls.Contains("toll") || Tolls.Contains("Toll"))
                    //                                    {
                    //                                        flag = true;
                    //                                    }
                    //                                }
                    //                                dtObj.Rows[i]["Mileage"] = distance;
                    //                                if (flag)
                    //                                    dtObj.Rows[i]["Tolls"] = "Yes";
                    //                                else
                    //                                    dtObj.Rows[i]["Tolls"] = "No";
                    //                            }
                    //                            else if (Status == "OVER_QUERY_LIMIT")
                    //                            {
                    //                                string ErrorMsg = (string)_Jobj.SelectToken("error_message");
                    //                                //MessageBox.Show(i + " Calcualtion crossed over limit at this time" + ErrorMsg);
                    //                                dtObj.Rows[i]["Mileage"] = ErrorMsg;
                    //                                if (flag)
                    //                                    dtObj.Rows[i]["Tolls"] = "";
                    //                                else
                    //                                    dtObj.Rows[i]["Tolls"] = "";
                    //                            }
                    //                            else
                    //                            {
                    //                                string ErrorMsg = (string)_Jobj.SelectToken("error_message");
                    //                                dtObj.Rows[i]["Mileage"] = ErrorMsg;
                    //                                if (flag)
                    //                                    dtObj.Rows[i]["Tolls"] = "";
                    //                                else
                    //                                    dtObj.Rows[i]["Tolls"] = "";
                    //                            }
                    //                        }
                    //                        catch (Exception ex)
                    //                        {
                    //                            dtObj.Rows[i]["Mileage"] = ex;
                    //                            if (flag)
                    //                                dtObj.Rows[i]["Tolls"] = "";
                    //                            else
                    //                                dtObj.Rows[i]["Tolls"] = "";
                    //                        }

                    //                    }
                    //                }
                    //            }
                    //            dtObjClone.Merge(dtObj, true);
                    //            //dtObj.AsEnumerable().CopyToDataTable(dtObjClone, LoadOption.Upsert);
                    //        }

                    //        iCount++;
                    //    }

                    //    #region Excel Format

                    //    if (dtObjClone.Rows.Count > 0)
                    //    {

                    //        Workbook workbook = new Workbook();
                    //        Worksheet worksheet = new Worksheet("Export Data");

                    //        //Headers
                    //        for (int i = 0; i < dtObjClone.Columns.Count; i++)
                    //        {

                    //            worksheet.Cells[0, i] = new Cell(dtObjClone.Columns[i].ColumnName);
                    //        }

                    //        // Content.  
                    //        for (int i = 0; i < dtObjClone.Rows.Count; i++)
                    //        {
                    //            for (int j = 0; j < dtObjClone.Columns.Count; j++)
                    //            {
                    //                worksheet.Cells[i + 1, j] = new Cell(dtObjClone.Rows[i][j].ToString());
                    //            }
                    //        }


                    //        //worksheet.Cells.ColumnWidth[0, 1] = 3000;
                    //        workbook.Worksheets.Add(worksheet);
                    //        workbook.Save(originalpath);

                    //        Workbook book = Workbook.Load(originalpath);
                    //        Worksheet sheet = book.Worksheets[0];
                    //    }

                    //    #endregion
                    //}
                    //else
                    //{

                    //    tbProgress.Minimum = 1;
                    //    tbProgress.Maximum = dtObj.Rows.Count;
                    //    tbProgress.Value = 1;
                    //    tbProgress.Step = 1;

                    //    if (dtObj.Rows.Count > 0)
                    //    {

                    //        for (int i = 0; i < dtObj.Rows.Count; i++)
                    //        {
                    //            bool flag = false;
                    //            if (dtObj.Rows[i]["puaddr"].ToString() != "" && dtObj.Rows[i]["daddr"].ToString() != "")
                    //            {
                    //                string Origin = "";
                    //                string Destination = "";
                    //                string Status = "";

                    //                //to check in other IP comment below line
                    //                string API_KEY = "AIzaSyDrdmH-HSG4o7F_Dnn93jQzsHc60RW2wj4";// "AIzaSyBIAXvn_JfN4eK2_9SNctV7fi72EKBpDTs"; //

                    //                Origin = dtObj.Rows[i]["puaddr"].ToString() + " " + dtObj.Rows[i]["pucity"].ToString() + " " + dtObj.Rows[i]["pust"].ToString() + " " + dtObj.Rows[i]["puzip"].ToString();

                    //                Destination = dtObj.Rows[i]["daddr"].ToString() + " " + dtObj.Rows[i]["dcity"].ToString() + " " + dtObj.Rows[i]["dst"].ToString() + " " + dtObj.Rows[i]["dzip"].ToString();

                    //                //to check in other IP comment below line
                    //                string url = "https://maps.googleapis.com/maps/api/directions/json?origin=" + Origin.Replace("#", "") + "&destination=" + Destination.Replace("#", "") + "&sensor=false&key=" + API_KEY + "&userIp=" + UserIP;

                    //                //to check in valid IP uncomment below line
                    //                // string url = "http://maps.googleapis.com/maps/api/directions/json?origin=" + Origin.Replace("#", "") + "&destination=" + Destination.Replace("#", "") + "&sensor=false";

                    //                string requesturl = url;
                    //                string content = fileGetJSON(requesturl);

                    //                if (content.ToUpper() == "DENIED")
                    //                {
                    //                    MessageBox.Show("Please make sure that your IP address is authorized to use the API key.", "MileageTracker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //                    return;
                    //                }
                    //                else if (content == "unable to connect to server")
                    //                {
                    //                    dtObj.Rows[i]["Mileage"] = "unable to connect to server";
                    //                    if (flag)
                    //                        dtObj.Rows[i]["Tolls"] = "";
                    //                    else
                    //                        dtObj.Rows[i]["Tolls"] = "";
                    //                }
                    //                else
                    //                {
                    //                    try
                    //                    {
                    //                        JObject _Jobj = JObject.Parse(content);
                    //                        Status = (string)_Jobj.SelectToken("status");
                    //                        tbProgress.PerformStep();
                    //                        if (Status == "OK")
                    //                        {
                    //                            distance = (string)_Jobj.SelectToken("routes[0].legs[0].distance.text");
                    //                            stepsCount = (int)_Jobj.SelectToken("routes[0].legs[0].steps").Count();

                    //                            for (int j = 0; j < stepsCount; j++)
                    //                            {
                    //                                Tolls = (string)_Jobj.SelectToken("routes[0].legs[0].steps[" + j + "].html_instructions");
                    //                                if (Tolls.Contains("toll") || Tolls.Contains("Toll"))
                    //                                {
                    //                                    flag = true;
                    //                                }
                    //                            }
                    //                            dtObj.Rows[i]["Mileage"] = distance;
                    //                            if (flag)
                    //                                dtObj.Rows[i]["Tolls"] = "Yes";
                    //                            else
                    //                                dtObj.Rows[i]["Tolls"] = "No";
                    //                        }
                    //                        else if (Status == "OVER_QUERY_LIMIT")
                    //                        {
                    //                            string ErrorMsg = (string)_Jobj.SelectToken("error_message");
                    //                            //MessageBox.Show(i + " Calcualtion crossed over limit at this time" + ErrorMsg);
                    //                            dtObj.Rows[i]["Mileage"] = ErrorMsg;
                    //                            if (flag)
                    //                                dtObj.Rows[i]["Tolls"] = "";
                    //                            else
                    //                                dtObj.Rows[i]["Tolls"] = "";
                    //                        }
                    //                        else
                    //                        {
                    //                            string ErrorMsg = (string)_Jobj.SelectToken("error_message");
                    //                            dtObj.Rows[i]["Mileage"] = ErrorMsg;
                    //                            if (flag)
                    //                                dtObj.Rows[i]["Tolls"] = "";
                    //                            else
                    //                                dtObj.Rows[i]["Tolls"] = "";
                    //                        }
                    //                    }
                    //                    catch (Exception Ex1)
                    //                    {
                    //                        dtObj.Rows[i]["Mileage"] = Ex1;
                    //                        if (flag)
                    //                            dtObj.Rows[i]["Tolls"] = "";
                    //                        else
                    //                            dtObj.Rows[i]["Tolls"] = "";
                    //                    }
                    //                }
                    //            }
                    //        }

                    //        #region Excel Format

                    //        if (dtObj.Rows.Count > 0)
                    //        {

                    //            Workbook workbook = new Workbook();
                    //            Worksheet worksheet = new Worksheet("Export Data");

                    //            //Headers
                    //            for (int i = 0; i < dtObj.Columns.Count; i++)
                    //            {

                    //                worksheet.Cells[0, i] = new Cell(dtObj.Columns[i].ColumnName);
                    //            }

                    //            // Content.  
                    //            for (int i = 0; i < dtObj.Rows.Count; i++)
                    //            {
                    //                for (int j = 0; j < dtObj.Columns.Count; j++)
                    //                {
                    //                    worksheet.Cells[i + 1, j] = new Cell(dtObj.Rows[i][j].ToString());
                    //                }
                    //            }


                    //            //worksheet.Cells.ColumnWidth[0, 1] = 3000;
                    //            workbook.Worksheets.Add(worksheet);
                    //            workbook.Save(originalpath);

                    //            Workbook book = Workbook.Load(originalpath);
                    //            Worksheet sheet = book.Worksheets[0];
                    //        }
                    //        #endregion
                    //    }
                    //}

                    try
                    {
                        Process.Start(originalpath);
                        this.Close();
                        Application.Exit();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + ", Please find excel file in the directory choosen.");
                    }

                }
                else
                {
                    MessageBox.Show("Please import input file in Excel file Format!");
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();



        }
        //public void LoadProgress()
        //{

        //    tbProgress.Invoke(new Action(() => tbProgress.Visible = true));


        //}

        protected string fileGetJSON(string fileName)
        {
            string _sData = string.Empty;
            string me = string.Empty;
            try
            {
                if (fileName.ToLower().IndexOf("http:") > -1 || fileName.ToLower().IndexOf("https:") > -1)
                {
                    System.Net.WebClient wc = new System.Net.WebClient();
                    byte[] response = wc.DownloadData(fileName);
                    _sData = System.Text.Encoding.ASCII.GetString(response);

                }
                else
                {
                    //int len = fileName.Length;
                    System.IO.StreamReader sr = new System.IO.StreamReader(fileName);
                    _sData = sr.ReadToEnd();
                    sr.Close();

                }
            }
            catch (Exception ex)
            {
                if (ex.Message.ToLower().Contains("illegal") || ex.Message.ToLower().Contains("the specified path, file name, or both are too long"))
                {
                    _sData = "DENIED";
                }
                else
                    _sData = "unable to connect to server";
            }
            return _sData;
        }




    }
}
