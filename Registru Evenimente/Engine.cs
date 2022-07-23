using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Remoting.Messaging;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Configuration;

namespace Registru_Evenimente
{
   public static class Engine
    {
        public static _Application excel = new _Excel.Application();
        public static Workbook wb;
        public static  Worksheet ws;
        public static DateTime date = new DateTime();
        public static ConnectionStringSettings setting = ConfigurationManager.ConnectionStrings["Registru_Evenimente.Properties.Settings.Setting"];
        public static int id;
        public static string excelPath= @"C:\Registru";
        public static void Excel(string p, int s)
        {
            wb = excel.Workbooks.Open(p);
            ws = excel.Worksheets[s];
        }
        public static bool VerifyExcel(string cod, int n)
        {
            while (ws.Cells[n, 3].Value != null)
            {
                if (cod == (string)ws.Cells[n, 3].Value) return true;
                n++;
            }
            return false;
        }
        public static void CloseExcel()
        {
            excel.Workbooks.Close();
            excel.Quit();
        }
    }
}
