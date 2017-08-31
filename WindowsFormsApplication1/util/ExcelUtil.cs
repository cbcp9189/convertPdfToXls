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
using System.Text.RegularExpressions;
using WindowsFormsApplication1.entity;


namespace WindowsFormsApplication1.util
{
    class ExcelUtil
    {
        
        //导入excel文件
        public List<KeyValEntity> createChildExcel(string excelFileName)
        {
            List<KeyValEntity> kyeList = new List<KeyValEntity>();
            List<String> pathList = new List<string>();
            // 根据excel的sheet生成多个excel
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

            //获取excel中的文本
            foreach (String path in pathList) {
                KeyValEntity kvey = new KeyValEntity();
                XLWorkbook workBook = new XLWorkbook(path);
                IXLWorksheet workSheet = workBook.Worksheet(1);
                var rows = workSheet.RowsUsed();
                StringBuilder text = new StringBuilder("");
                foreach (var row in rows)
                {
                    //遍历所有的Cells
                    foreach (var cell in row.Cells())
                    {
                        if (cell.DataType == XLCellValues.DateTime)
                        {
                                String val = cell.RichText.ToString().Replace("@", "");
                                text.Append(val + " ");
                        }
                        else
                        {
                            text.Append(cell.RichText.ToString() + " ");
                        }
                        
                    }
                }
                kvey.key = path;
                kvey.value = Regex.Replace(text.ToString(), "\\s+", " ");
                kyeList.Add(kvey);
            }
            return kyeList;
        }

    }
}
