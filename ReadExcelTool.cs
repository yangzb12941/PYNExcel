using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PYNExcel
{
    internal class ReadExcelTool
    {

        private Workbook wb = null;

        public ReadExcelTool(string strFileName)
        {
            object missing = System.Reflection.Missing.Value;
            Application excel = new Application();//lauch excel application
            excel.Visible = false; excel.UserControl = true;
            // 以只读的形式打开EXCEL文件
            this.wb = excel.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
             missing, missing, missing, true, missing, missing, missing, missing, missing);
        }

        public Worksheet getWorksheet(string sheetName) 
        {
            //取得第一个工作薄
            Worksheet ws = (Worksheet)this.wb.Worksheets.get_Item(sheetName);
            return ws;
        }

        public Workbook getWorkbook() 
        {
            return this.wb;
        }
    }
}
