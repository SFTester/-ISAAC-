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

namespace BIS4._1_API_Application
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var ace = new AccessEngine();
            Application.Run(new FormMain());
        }
    }

}

namespace global
{
    class Data
    {
        public static string    WORKFILESPATH   = "";
        public static string    LOGAPP          = "log_Application.csv";
        public static string    LOGAPI          = "log_API.csv";
        public static string    LOGSRV          = "log_Service.csv";
        public static int       FILE_COLUMNS    = 0;
        public static int       FILE_ROWS       = 0;
        public static int       REC_N           = 0;
        public static bool      F_LOGAPP        = true;
        public static bool      F_LOGAPP_EXT    = false;
        public static bool      F_LOGAPI        = true;
        public static bool      F_LOGSRV        = true;

        public static void      logger(string msg, string mode)
        {
            try
            {
                if ((Convert.ToInt32(mode, 2) & Convert.ToInt32("00000001", 2))>0 && F_LOGAPP)              //mode = 00000001 Application log
                {
                    string PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGAPP;
                    FileStream fs = new FileStream(PATH, FileMode.Append, FileAccess.Write);
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                    sw.WriteLine(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; " + msg);
                    sw.Close(); fs.Close();
                }
                if ((Convert.ToInt32(mode, 2) & Convert.ToInt32("00000010", 2)) > 0 && F_LOGAPI)            //mode = 00000010 API log
                {
                    string PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGAPI;
                    FileStream fs = new FileStream(PATH, FileMode.Append, FileAccess.Write);
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                    sw.WriteLine(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; " + msg);
                    sw.Close(); fs.Close();
                }
                if ((Convert.ToInt32(mode, 2) & Convert.ToInt32("00000100", 2)) > 0 && F_LOGSRV)            //mode = 00000100 Service log
                {
                    string PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGSRV;          
                    FileStream fs = new FileStream(PATH, FileMode.Append, FileAccess.Write);
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                    sw.WriteLine(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "   -------------------------------------------");
                    sw.WriteLine(msg);
                    sw.Close(); fs.Close();
                }
                if ((Convert.ToInt32(mode, 2) & Convert.ToInt32("00001000", 2)) > 0 && F_LOGAPP_EXT)        //mode = 00001000 Extended log
                {
                    string PATH = global.Data.WORKFILESPATH + @"\_logs\" + global.Data.LOGAPP;
                    FileStream fs = new FileStream(PATH, FileMode.Append, FileAccess.Write);
                    Encoding utf8WithoutBom = new UTF8Encoding(false);
                    StreamWriter sw = new StreamWriter(fs, utf8WithoutBom);
                    sw.WriteLine(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; " + msg);
                    sw.Close(); fs.Close();
                }
                if ((Convert.ToInt32(mode, 2) & Convert.ToInt32("10000000", 2)) > 0)                        //mode = 10000000 Console
                {
                    Console.WriteLine(DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss.fff") + "; " + msg);
                }
            }
            catch { }
        }
    }

} 



