using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Agent2._0
{
    class Excel
    {
        string path;

        public string FileSerialNum { get; }

        public string FileDate { get; }

        public String FileResult { get; }

        _Application xlApp;
        public Workbook wb { get; set; }
        public Worksheet ws { get; set; }

        public Excel(string path, int sheet)
        {
            this.path = path;

            try
            {
                string FileName = Path.GetFileNameWithoutExtension(path);
                string[] parts = FileName.Split("_");
                FileSerialNum = parts[0];
                FileDate = GetDateFromFileName(parts[1]);
                FileResult = parts[2];
            }
            catch(Exception e)
            {
                Console.WriteLine("Error: File name wrong format");
                Console.WriteLine("Exception: " + e.Message);
                FileSerialNum = "";
                FileDate = "";
                FileResult = "";
            }

            xlApp = new _Excel.Application();
            wb = xlApp.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }
        public static string GetDateFromFileName(string rawDate)
        {
            DateTime dt;
            try
            {
                dt = DateTime.ParseExact(rawDate, "ddMMyyyy", CultureInfo.InvariantCulture);
                return dt.ToString("yyyy-MM-dd");
            }
            catch (FormatException e)
            {
                Console.WriteLine("Read date error: " + e.Message);
                return "";
            }
        }

        public string ReadCell(int row, int col)
        {
            if (row > 0 && col > 0 && (ws.Cells[row, col]).Value2 != null)
            {
                return ws.Cells[row, col].Text;
            }
            else
                return "";
        }

        public _Excel.Range Find(string range, string text)
        {
            try
            {
                return ws.Range[range].Find(text);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
            return null;
        }

        public void Close()
        {
            try
            {
                KillExcel(this.xlApp);
                //System.Threading.Thread.Sleep(100);
            }
            catch(Exception e)
            {
                Console.WriteLine("Error: " + e.StackTrace);
            }
        }
        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
        private static void KillExcel(_Excel._Application theApp)
        {
            int id = 0;
            IntPtr intptr = new IntPtr(theApp.Hwnd);
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out id);
                p = System.Diagnostics.Process.GetProcessById(id);
                if (p != null)
                {
                    p.Kill();
                    p.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("KillExcel:" + ex.Message);
            }
        }
    }
}
