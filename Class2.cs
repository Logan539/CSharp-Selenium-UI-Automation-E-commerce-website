using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ShoppingCart_Excel
{
    public class Input_Class
    {
        public static Application xlapp = new Application();
        
        public static Workbook xlwb = xlapp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        public static Worksheet xlws= xlwb.Worksheets[1];
        public static Excel.Range xlRange = xlws.UsedRange;
        public int rowCount = xlRange.Rows.Count;
        public int colCount = xlRange.Columns.Count;
        public List<string> price = new List<string>();
        public List<string> discounted_price = new List<string>();
        public List<string> product_code = new List<string>();
        public List<string> site_quant = new List<string>();
        static string[] q_value = { "2", "5", "10", "5", "20", "10", "5", "5", "1", "1", "10", "5" }; 
        public List<string> quant = new List<string>(q_value);

        public List<string> w = new List<string>();
        public string[] y;
        public string filepath = @"ExcelFilePath-" + DateTime.Now.ToString("dd-MM-yyyy");

        /*public void Createfile()
        {
            xlapp.Visible = true;
            xlwb = xlapp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            try
            {
                xlws = xlwb.Worksheets[1];
                xlwb.Worksheets[1].Name = "MySheet";
                Console.WriteLine("firstExcel" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss-fff") + ".xlsx");
                xlwb.SaveAs(@"ExcelFIlepath\Input"+ DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + "");
                xlwb.Close();
                xlapp.Quit();
                Marshal.ReleaseComObject(xlws);
                Marshal.ReleaseComObject(xlwb);
                Marshal.ReleaseComObject(xlapp);
            }
            catch (Exception exHandle)
            {
                Console.WriteLine("Exception: " + exHandle.Message);
                Console.ReadLine();
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                    process.Kill();
            }*/
        
    }
}
