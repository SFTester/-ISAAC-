using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Microsoft.SqlServer;
using System.Xml;

namespace BIS4._1_API_Application
{
    public class DataProcessor : Form
    {
    public DataProcessor()
        {
         initiateLogging();
        }

    public virtual string   DPName { get; set; }
    public virtual string   LOG_Allowed { get; set; }
    public virtual int      RAM_LOGLength { get; set; }
    public virtual bool     LOG_Overwrite { get; set; }
    public delegate void    INFORMER (object sender, string inf, int bar1, int bar2, int bar3);
    public event            INFORMER OnNeedSomething;
    public event            INFORMER updateProgressBarsMaximum;
    public event            INFORMER updateProgressBarsCurrent;
    public string           msgError        = "";
    public string           msgErrorType    = "";
    public string           dataType        = "";

    public DataTable RAM_APPLog = new DataTable("RAM_APPLog");
    public DataTable RAM_APILog = new DataTable("RAM_APILog");
    public DataTable RAM_SRVLog = new DataTable("SRV_APPLog");
    public DataTable RAM_EXTLog = new DataTable("RAM_EXTLog");

    public int RAM_APPLog_Rec_N = 0;
    public int RAM_APILog_Rec_N = 0;
    public int RAM_SRVLog_Rec_N = 0;
    public int RAM_EXTLog_Rec_N = 0;

    private void DataProcessor_Load(object sender, EventArgs e)
    {

    }

    public void initiateLogging()
    {
    RAM_APPLog.Columns.Add("RAM LOG. APP. DUMP START --------------");
    RAM_APILog.Columns.Add("RAM LOG. API. DUMP START --------------");
    RAM_SRVLog.Columns.Add("RAM LOG. SRV. DUMP START --------------");
    RAM_EXTLog.Columns.Add("RAM LOG. EXT. DUMP START --------------");
    RAM_APPLog_Rec_N = 0;
    RAM_APILog_Rec_N = 0;
    RAM_SRVLog_Rec_N = 0;
    RAM_EXTLog_Rec_N = 0;
    }

    private void InitializeComponent()
    {
        //this.SuspendLayout();
        // 
        // DataProcessor
        // 
        //this.ClientSize = new System.Drawing.Size(284, 261);
        //this.Name = "DataProcessor";
        //this.Load += new System.EventHandler(this.DataProcessor_Load);
        //this.ResumeLayout(false);

    }

    public string txtReadFromFile(string nameFile)                                  //reads string from file
    {
        dataType = "TXT from FILE";
        string txt = "";
        if (File.Exists(nameFile))
        {
            try
            {
                FileInfo txtFile = new FileInfo(nameFile);
                txt = txtFile.OpenText().ReadToEnd();
            }
            catch (Exception e)
            {
                msgError = "txtReadFromFile. \r\n" +
                            dataType + "\r\n" +
                            e.Message + "\r\n" +
                            "1. Check if file has proper format\r\n" +
                "Options: Abort the application or Retry to start new operation or Ignore this message";
                msgErrorType = "GENERAL EXCEPTION";
                global.Data.logger(msgError, "10000100");
                DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); return ""; }
                if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); return ""; }
            }
        }
        else
        {
            string RESULT = "txtReadFromFile; " + "FILE; " + "; READ; " + " File not found:" + nameFile; global.Data.logger(RESULT, "10000100");
        }
        return txt;
    }
    
    public bool txtSaveToFile(string txt, string nameFile)                          //saves string to file
    {
        dataType = "TXT to FILE";
        if (!String.IsNullOrEmpty(nameFile))
        {
            try
            {
                if (File.Exists(nameFile)) File.Delete(nameFile);

                if (!File.Exists(nameFile))
                {
                    using (FileStream fs = new FileStream(nameFile, FileMode.Append, FileAccess.Write))
                    {
                        Encoding utf8WithoutBom = new UTF8Encoding(false);
                        StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                        sw.WriteLine(txt, utf8WithoutBom);
                        sw.Close(); fs.Close();
                        global.Data.logger("txtSaveToFile. File is saved; " + nameFile, "00001000");
                        return true;
                    }
                }
                else
                {
                    msgError = "txtSaveToFile. Can not replace existing file: " + nameFile + "\r\nPossibly, file is in use now, try a few seconds later";
                    global.Data.logger(msgError, "00001000");
                    MessageBox.Show(msgError);
                    return false;
                }
            }
            catch (Exception e)
            {
                msgError = "txtSaveToFile. \r\n" +
                            dataType + "\r\n" +
                            e.Message + "\r\n" +
                            "1. Check if file can be saved in this place\r\n" +
                "Options: Abort the application or Retry to start new operation or Ignore this message";
                msgErrorType = "GENERAL EXCEPTION";
                global.Data.logger(msgError, "10000100");
                DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                if (DialogResult == DialogResult.Abort)     { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                if (DialogResult == DialogResult.Retry)     { global.Data.logger(DialogResult.ToString(), "10000100"); return false; }
                if (DialogResult == DialogResult.Ignore)    { global.Data.logger(DialogResult.ToString(), "10000100"); return false; }
            }
        }
        else
        {
            string RESULT = "txtSaveToFile; " + "FILE; " + "; SAVE; " + " File not found:" + nameFile; global.Data.logger(RESULT, "10000100");
            return false;
        }
        return false;
    }

    public DataTable dtReadFromFile(string nameFile, bool F_HASHEADERS)             //reads datatable from csv-file
    {
        dataType = "DataTable from FILE";
        DataTable dtRead = new DataTable("dt");
        dtRead.Clear(); dtRead.Columns.Clear();
        string Line;
        string[] strArrIn = new string[1000];
        if (File.Exists(nameFile)) try
            {
                //if (updateProgressBarsCurrent != null) updateProgressBarsMaximum(this, dataType, System.IO.File.ReadAllLines(nameFile).Length, -1, -1);
                //if (updateProgressBarsCurrent != null) updateProgressBarsCurrent(this, dataType, 0, -1, -1);
                updateProgressBarsMaximum(this, dataType, System.IO.File.ReadAllLines(nameFile).Length, -1, -1);
                updateProgressBarsCurrent(this, dataType, 0, -1, -1);

                using (FileStream fs = new FileStream(nameFile, FileMode.Open))
                {
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamReader sr = new StreamReader(fs, utf8WithoutBom);
                    DataColumn[] strArrColumnsIn = new DataColumn[1000];

                    int i = 0; int j = 0;                                                                           // i = column, j = row
                    while (sr.EndOfStream != true)
                    {
                        Line = sr.ReadLine();                                                                       // read current row from file
                        // add columns to datatable:
                        if (j == 0 && F_HASHEADERS)                                                                 // if row #0 of file has headers...
                        {
                            strArrIn = Line.Split(';');                                                             // split row into cells
                            for (i = 0; i < strArrIn.Length; i++)                                                   // add headers taken from row #0
                            {
                                strArrColumnsIn[i] = new DataColumn(strArrIn[i].Trim(), typeof(String));
                                dtRead.Columns.Add(strArrColumnsIn[i]);
                            }
                        }

                        if (j == 0 && !F_HASHEADERS)                                                                // if row #0 of file has no headers...
                        {
                            strArrIn = Line.Split(';');
                            for (i = 0; i < strArrIn.Length; i++)                                                   // add numeric headers to the beginning of file (from row #0)
                            {
                                strArrColumnsIn[i] = new DataColumn(i.ToString(), typeof(String));
                                dtRead.Columns.Add(strArrColumnsIn[i]);
                            }
                        }

                        if (j != 0)                                                                                 // put all other (then #0) rows into datatable
                        {
                            dtRead.NewRow();
                            dtRead.Rows.Add(Line.Split(';'));
                        }

                        j++;
                        //if (updateProgressBarsCurrent != null) updateProgressBarsCurrent(this, dataType, j, -1, -1);
                        updateProgressBarsCurrent(this, dataType, j, -1, -1);
                    }
                    global.Data.FILE_ROWS = dtRead.Rows.Count + 1;
                    global.Data.FILE_COLUMNS = dtRead.Columns.Count;
                    sr.Close(); fs.Close();
                    global.Data.logger("dtReadFromFile. DataTable is filled from file; " + nameFile + "; " + global.Data.FILE_COLUMNS + " columns and " + global.Data.FILE_ROWS + " rows", "00001000");
                }
            }
            catch (Exception e)
            {
                msgError = "dtReadFromFile. \r\n" +
                            dataType + "\r\n" +
                            e.Message + "\r\n" +
                            "1. Check if all involved files have headers\r\n" + 
                            "2. Check if all involved files has consistent line structure (MEMO field can be an issue)\r\n" +
                "Options: Abort the application or Retry to start new operation or Ignore this message";
                msgErrorType = "GENERAL EXCEPTION";
                global.Data.logger(msgError, "10000100");
                DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); }
                if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); }
            }
        return dtRead;
    }

    public bool dtSaveToFile(DataTable dt, string nameFile, bool overWrite)         //saves datatable to csv-file
    {
        dataType = "DT to FILE";
        if (!String.IsNullOrEmpty(nameFile)) 
            try
            {
                if (File.Exists(nameFile) && overWrite) File.Delete(nameFile);

                using (FileStream fs = new FileStream(nameFile, FileMode.Append, FileAccess.Write))
                {
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);

                    string headers = "";
                    foreach (DataColumn col in dt.Columns) headers += col.Caption + ";";                        // read headers
                    sw.WriteLine(headers.Substring(0, headers.Length - 1));                                     // write headers to file (without final ";")
                    foreach (DataRow row in dt.Rows) sw.WriteLine(String.Join(";", row.ItemArray));             // read and write data to file
                    sw.Close(); fs.Close();
                    global.Data.logger("dtSaveToFile. File is saved; " + nameFile + "; " + dt.Columns.Count + " columns and " + dt.Rows.Count + " data rows", "00001000");
                }
            }
            catch (Exception e)
            {
                msgError = "dtSaveToFile. \r\n" +
                            dataType + "\r\n" +
                            e.Message + "\r\n" +
                            "General failure\r\n" +
                "Options: Abort the application or Retry to start new operation or Ignore this message";
                msgErrorType = "GENERAL EXCEPTION";
                global.Data.logger(msgError, "10000100");
                DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); return false; }
                if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); return true; }
            }
        return false;
    }

    public DataTable dtReadFromDgv(DataGridView dgv)
    {
        dataType = "DGV to DT";
        DataTable dtRead = new DataTable("dt");
        dtRead.Rows.Clear(); dtRead.Clear(); dtRead.Columns.Clear();
        if (true) try
            {
                for (int i = 0; i < dgv.Columns.Count; i++) dtRead.Columns.Add(dgv.Columns[i].Name);
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    DataRow newRow = dtRead.NewRow();
                    for (int i = 0; i < dgv.Columns.Count; i++) newRow[i] = row.Cells[i].Value;
                    dtRead.Rows.Add(newRow);
                }
                dtRead.Rows[dtRead.Rows.Count-1].Delete();                          // removes last row which is always empty
            }
            catch (Exception e)
            {
                msgError = "dtReadFromDgv. \r\n" +
                            dataType + "\r\n" +
                            e.Message + "\r\n" +
                            "General failure\r\n" +
                "Options: Abort the application or Retry to start new operation or Ignore this message";
                msgErrorType = "GENERAL EXCEPTION";
                global.Data.logger(msgError, "10000100");
                DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); }
                if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); }
            }
        return dtRead;
    } 

    public void dgvSaveToFile(DataGridView dgv, string saveToFile)                  //saves datagridview to csv-file
    {
        dataType = "DGV to FILE";
        if (!String.IsNullOrEmpty(saveToFile))
            {
            dtSaveToFile(dtReadFromDgv(dgv), saveToFile, true);
            global.Data.logger("dgvSaveToFile. File is updated; " + saveToFile + "; " + dgv.ColumnCount + " columns and " + dgv.RowCount + " rows", "00001000");
            }
    }
    
    public string dtFindDB(string connFile, string sqlRqFile, string Data1, string Data2, string Data3, string Data4)   //looks for and returns single found value from SQL database
    {
        string connString   = "";
        string sqlRq        = "";
        string dataFound    = "";

        connString = txtReadFromFile(connFile);
        sqlRq = txtReadFromFile(sqlRqFile);

        if (!String.IsNullOrEmpty(connString) && !String.IsNullOrEmpty(sqlRq))
        {
            try
            {
                SqlConnection con;
                using (con = new SqlConnection())
                {
                    con.ConnectionString = connString;
                    con.Open();
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = con;
                    cmd.Parameters.AddWithValue("@Data1", Data1);
                    cmd.Parameters.AddWithValue("@Data2", Data2);
                    cmd.Parameters.AddWithValue("@Data3", Data3);
                    cmd.Parameters.AddWithValue("@Data4", Data4);
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = sqlRq;
                    object result = cmd.ExecuteScalar();
                    if (result != null) dataFound = result.ToString();
                    con.Close();
                }
                con.Dispose();
            }
            catch (SqlException se) 
            { 
                string RESULT = "dtFindDB; " + "SQL DATA; " + "; FIND; " + "General failure" + se.ToString(); 
                global.Data.logger(RESULT, "00001000");
                global.Data.logger("dtFindDB; " + "SQL DATA; " + "; FIND; ERROR", "10000100"); 
                return ""; 
            }
        }
        else
        {
            string RESULT = "dtFindDB; " + "SQL DATA; " + "; FIND; " + " ConnString or sqlRq is empty"; global.Data.logger(RESULT, "10000100"); return ""; ;
        }

        return dataFound;
    }

    public DataTable dtReadDB(string connFile, string sqlRqFile, string Data1, string Data2, string Data3, string Data4) //reads datatable from SQL database
    {
        DataTable dtRead = new DataTable("dt");
        dtRead.Clear(); dtRead.Columns.Clear();

        string connString = "";
        string sqlRq = "";

        connString = txtReadFromFile(connFile);
        sqlRq = txtReadFromFile(sqlRqFile);

        if (!String.IsNullOrEmpty(connString) && !String.IsNullOrEmpty(sqlRq)) try
            {
                SqlConnection con;
                using (con = new SqlConnection())
                {
                    con.ConnectionString = connString;
                    con.Open();
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = con;
                    cmd.Parameters.AddWithValue("@Data1", Data1);
                    cmd.Parameters.AddWithValue("@Data2", Data2);
                    cmd.Parameters.AddWithValue("@Data3", Data3);
                    cmd.Parameters.AddWithValue("@Data4", Data4);
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = sqlRq;
                    using (SqlDataReader dr = cmd.ExecuteReader()) 
                    dtRead.Load(dr);
                    con.Close();
                }
            }
            catch (SqlException se) { string RESULT = "dtReadDB; " + "SQL DATA; " + "; SEND; " + "General failure" + se.ToString(); global.Data.logger(RESULT, "10000100"); }

        return dtRead;
    }

    public string readXML(string path, int par)                                     // reads one item from xml file
    {
        string parValue = "";
        if (File.Exists(path)) try
            {
                XmlDocument docXML = new XmlDocument();
                docXML.Load(path);
                parValue = docXML.DocumentElement.ChildNodes[par].InnerText;
            }
            catch (Exception e) { string RESULT = "readXML; " + "XML DATA; " + "; READ; " + "General failure" + e.ToString(); global.Data.logger(RESULT, "10000100"); }
        return parValue;
    }     

    public uint DAY  (string value, string dtFormat)                                // extracts and returnes day from data
        {
            uint val = 0;
            if (dtFormat == "yyyy-MM-dd HH:mm:ss.000")  val = Convert.ToUInt32(value.Substring(8, 2));
            if (dtFormat == "dd.MM.yyyy HH:mm:ss")      val = Convert.ToUInt32(value.Substring(0, 2));
            if (dtFormat == "yyyyMMdd")                 val = Convert.ToUInt32(value.Substring(6, 2));
            if (dtFormat == "")
            {
                if (value.Length == 23)                 val = Convert.ToUInt32(value.Substring(8, 2));            //"yyyy-MM-dd HH:mm:ss.000"
                if (value.Length == 19)                 val = Convert.ToUInt32(value.Substring(0, 2));            //"dd.MM.yyyy HH:mm:ss"
                if (value.Length == 8)                  val = Convert.ToUInt32(value.Substring(6, 2));            //"yyyyMMdd"
            }
            return val;
        }
    public uint MONTH(string value, string dtFormat)                                // extracts and returnes month from data
        {
            uint val = 0;
            if (dtFormat == "yyyy-MM-dd HH:mm:ss.000")  val = Convert.ToUInt32(value.Substring(5, 2));
            if (dtFormat == "dd.MM.yyyy HH:mm:ss")      val = Convert.ToUInt32(value.Substring(3, 2));
            if (dtFormat == "yyyyMMdd")                 val = Convert.ToUInt32(value.Substring(4, 2));
            if (dtFormat == "")
            {
                if (value.Length == 23)                 val = Convert.ToUInt32(value.Substring(5, 2));            //"yyyy-MM-dd HH:mm:ss.000"
                if (value.Length == 19)                 val = Convert.ToUInt32(value.Substring(3, 2));            //"dd.MM.yyyy HH:mm:ss"
                if (value.Length == 8)                  val = Convert.ToUInt32(value.Substring(4, 2));            //"yyyyMMdd"
            }
            return val;
        }
    public uint YEAR (string value, string dtFormat)                                // extracts and returnes year from data
        {
            uint val = 0;
            if (dtFormat == "yyyy-MM-dd HH:mm:ss.000")  val = Convert.ToUInt32(value.Substring(0, 4));
            if (dtFormat == "dd.MM.yyyy HH:mm:ss")      val = Convert.ToUInt32(value.Substring(6, 4));
            if (dtFormat == "yyyyMMdd")                 val = Convert.ToUInt32(value.Substring(0, 4));
            if (dtFormat == "")
            {
                if (value.Length == 23)                 val = Convert.ToUInt32(value.Substring(0, 4));            //"yyyy-MM-dd HH:mm:ss.000"
                if (value.Length == 19)                 val = Convert.ToUInt32(value.Substring(6, 4));            //"dd.MM.yyyy HH:mm:ss"
                if (value.Length == 8)                  val = Convert.ToUInt32(value.Substring(0, 4));            //"yyyyMMdd"
            }
            return val;
        }

    public bool logger(string logMsg, string logMode)           // logMode indicates where the logMsg will be logged, LOG_Allowed indicates which logs are turned on or off
    {
        try
        {
            if (logMode.Substring(1, 1) == "1")                                                                                 //mode = *1****** = RAM logging enabled
            {
                if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("00000001", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 01000001 Application RAM log
                {
                    RAM_APPLog.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; r; " + logMsg);
                    RAM_APPLog_Rec_N++;
                    if (RAM_APPLog_Rec_N > RAM_LOGLength) dumpRAMLog(0);                                                         // dump RAM Log to HDD
                }

                if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("00000010", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 01000010 API RAM log
                {
                    RAM_APILog.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; r; " + logMsg);
                    RAM_APILog_Rec_N++;
                    if (RAM_APILog_Rec_N > RAM_LOGLength) dumpRAMLog(1);                                                         // dump RAM Log to HDD
                }

                if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("00000100", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 01000100 Service RAM log
                {
                    RAM_SRVLog.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; r; " + logMsg);
                    RAM_SRVLog_Rec_N++;
                    if (RAM_SRVLog_Rec_N > RAM_LOGLength) dumpRAMLog(2);                                                         // dump RAM Log to HDD
                }

                if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("10000000", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 10000000 Console
                {
                    Console.WriteLine(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; " + logMsg);
                }

                return true;
            }
            else                                                                                                                //mode = *0****** = HDD logging enabled
            {
                if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("00000001", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 00000001 Application HDD log
                {
                    string PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGAPP;
                    FileStream fs = new FileStream(PATH, FileMode.Append, FileAccess.Write);
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                    sw.WriteLine(DateTime.Now + "; " + logMsg);
                    sw.Close(); fs.Close();
                }
                if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("00000010", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 00000010 API log
                {
                    string PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGAPI;
                    FileStream fs = new FileStream(PATH, FileMode.Append, FileAccess.Write);
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                    sw.WriteLine(DateTime.Now + "; " + logMsg);
                    sw.Close(); fs.Close();
                }
                if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("00000100", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 00000100 Service log
                {
                    string PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGSRV;
                    FileStream fs = new FileStream(PATH, FileMode.Append, FileAccess.Write);
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                    sw.WriteLine(DateTime.Now + "   -------------------------------------------");
                    sw.WriteLine(logMsg);
                    sw.Close(); fs.Close();
                }
                if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("00001000", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 00001000 Extended log
                {
                    string PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGAPP;
                    FileStream fs = new FileStream(PATH, FileMode.Append, FileAccess.Write);
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                    sw.WriteLine(DateTime.Now + "; " + logMsg);
                    sw.Close(); fs.Close();
                }

            }

            if ((Convert.ToInt32(logMode, 2) & Convert.ToInt32("10000000", 2) & Convert.ToInt32(LOG_Allowed, 2)) > 0)       //mode = 10000000 Console
            {
                Console.WriteLine(DateTime.Now + "; " + logMsg);
            }

            return true;
        }
        catch { return false; }



    }

    public void dumpRAMLog(int typeOfLog)
    {
        string PATH = "";
        
        // save RAM_APPLog to HDD logfile
        if (typeOfLog == 0 || typeOfLog == 9)
        {
            PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGAPP;
            RAM_APPLog.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; RAMLog; RAM APP LOG SAVED to HDD logfile");
            dtSaveToFile(RAM_APPLog, PATH, false);
            RAM_APPLog_Rec_N = 0;
            RAM_APPLog.Rows.Clear();
        }

        // save RAM_APILog to HDD logfile
        if (typeOfLog == 1 || typeOfLog == 9)
        {
            PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGAPI;
            RAM_APILog.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; RAMLog; RAM API LOG SAVED to HDD logfile");
            dtSaveToFile(RAM_APILog, PATH, false);
            RAM_APILog_Rec_N = 0;
            RAM_APILog.Rows.Clear();
        }

        // save RAM_SRVLog to HDD logfile
        if (typeOfLog == 2 || typeOfLog == 9)
        {
            PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGSRV;
            RAM_SRVLog.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; RAMLog; RAM SRV LOG SAVED to HDD logfile");
            dtSaveToFile(RAM_SRVLog, PATH, false);
            RAM_SRVLog_Rec_N = 0;
            RAM_SRVLog.Rows.Clear();
        }
    }

    #region DP. APPENDIX----------------------------------------------------------------------
    public void interClass()
    {
        MessageBox.Show("dp: " + DPName);
        if (OnNeedSomething != null) OnNeedSomething(this, "hi", 0, 0, 0);
    }
#endregion-------------------
    }
}
