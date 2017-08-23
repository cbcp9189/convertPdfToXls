using System;
using System.Collections.Generic;
using System.Reflection;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using ClosedXML.Excel;


namespace WindowsFormsApplication1.util
{
    class ExcelUtil
    {
        
        //导入excel文件
        public List<String> createChildExcel(string excelFileName)
        {
            List<String> pathList = new List<string>();
            // copy a sheet from one workbook to another:
            var wbSource = new XLWorkbook(excelFileName);
            int size = wbSource.Worksheets.Count;
            for (int i = 1; i <= size; i++)
            {
                string tempPath = Path.ChangeExtension(excelFileName, i + ".xlsx");
                var wb = new XLWorkbook();
                wb.Worksheets.Add("Sheet1");
                wb.SaveAs(tempPath);
                wbSource.Worksheet(i).CopyTo(wb, "table1", 1);
                wb.SaveAs(tempPath,true);
                pathList.Add(tempPath);
            }
            return pathList;
        }

    }
}
