using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PYNExcel
{
    internal class WriteExcelTool
    {
        private List<StatisticalData> data;
        public WriteExcelTool(List<StatisticalData> t)
        { 
           this.data = t;
        }

        public List<StatisticalData> Data { get => data; set => data = value; }

        public string writeExcel()
        {
            string fileName = "";
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                return fileName;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "产品";
            xlWorkSheet.Cells[1, 2] = "公司抬头";
            xlWorkSheet.Cells[1, 3] = "清关类型";
            xlWorkSheet.Cells[1, 4] = "已结关额度";
            xlWorkSheet.Cells[1, 5] = "清关中额度";
            xlWorkSheet.Cells[1, 6] = "总额度";
            xlWorkSheet.Cells[1, 7] = "增量补贴（万元）";
            for (int index = 0;index < data.Count;index++)
            {
                xlWorkSheet.Cells[index+2, 1] = data[index].GoodsName;
                xlWorkSheet.Cells[index + 2, 2] = data[index].CompanyName;
                xlWorkSheet.Cells[index + 2, 3] = data[index].CustomstType;
                xlWorkSheet.Cells[index + 2, 4] = data[index].ClearedQuota/10000;
                xlWorkSheet.Cells[index + 2, 5] = data[index].CustomsClearance / 10000;
                xlWorkSheet.Cells[index + 2, 6] = data[index].CustomsTotal / 10000;
                xlWorkSheet.Cells[index + 2, 7] = data[index].IncrementalSubsidy / 10000;
            }
            fileName = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + Guid.NewGuid().ToString() + ".xls";
            xlWorkBook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+ Guid.NewGuid().ToString()+ ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return fileName;
        }
    }
}
