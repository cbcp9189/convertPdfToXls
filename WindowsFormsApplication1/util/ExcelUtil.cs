﻿using System;
using System.Collections.Generic;
using System.Reflection;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public List<IXLWorksheet> getExcelSheetList(string excelFileName)
        {
            try
            {
                List<IXLWorksheet> sheetList = new List<IXLWorksheet>();
                // 根据excel的sheet生成txt
                var wbSource = new XLWorkbook(excelFileName);
                int size = wbSource.Worksheets.Count;
                for (int i = 1; i <= size; i++)
                {
                    sheetList.Add(wbSource.Worksheet(i));
                }
                return sheetList;
            }
            catch (Exception ex) {
                throw ex;
            }
        }

        public String getExcelSheetText(IXLWorksheet sheet)
        {
            try
            {
                var rows = sheet.RowsUsed();
                StringBuilder text = new StringBuilder("");
                foreach (var row in rows)
                {
                    //遍历所有的Cells
                    foreach (var cell in row.Cells())
                    {
                        text.Append(cell.RichText.ToString() + " ");
                    }
                }
                return Regex.Replace(text.ToString(), "\\s+", " ");
            }
            catch (Exception ex) {
                throw ex;
            }
        }

        public void createExcelBySheet(IXLWorksheet sheet, String excelPath) 
        {
            try
            {
                // 根据excel的sheet生成excel
                var wb = new XLWorkbook();
                wb.Worksheets.Add("Sheet1");
                wb.SaveAs(excelPath);
                sheet.CopyTo(wb, "table1", 1);
                wb.SaveAs(excelPath, true);
            }
            catch (Exception ex) {
                throw ex; 
            }
        }

        public void createExcelBySheetList(List<IXLWorksheet> sheetList,String excelPath,int startIndex,int endIndex)
        {
            try{
                var newWork = new XLWorkbook();
                newWork.Worksheets.Add("table1");
                var newSheet = newWork.Worksheet(1);
                int local = 1;
                for (int i = startIndex; i <= endIndex; i++) 
                {
                    var ws = sheetList[i];
                    var firstTableCell = ws.FirstCell();
                    var lastTableCell = ws.LastCellUsed();
                
                    var rngData = ws.Range(firstTableCell.Address, lastTableCell.Address);
                    //设置样式
                    newSheet.Style.Alignment = ws.Style.Alignment;
                    newSheet.Style.Border = ws.Style.Border;
                    newSheet.Style.Font = ws.Style.Font;
                    newSheet.Style.Fill = ws.Style.Fill;
                    newSheet.Style.NumberFormat = ws.Style.NumberFormat;
               
                    int num = ws.LastRowUsed().RowNumber();
                    newSheet.Cell(local, 1).Value = rngData;
                    //设置宽度和列的高度
                    for (int j = 1; j <= ws.LastColumnUsed().ColumnNumber(); j++) 
                    {
                        newSheet.Column(j).Width = ws.Column(j).Width;
                    }
                    for (int j = local; j < ws.LastRowUsed().RowNumber(); j++)
                    {
                        newSheet.Row(j).Height = ws.Row(j).Height;
                    }
                    local += num;
                }
                newWork.SaveAs(excelPath,true);
            }
            catch (Exception ex) {
                throw ex; 
            }
        }

        public static void TestExcel(string excelFileName)
        {
            

            //获取excel中的文本
                XLWorkbook workBook = new XLWorkbook(excelFileName);
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
                            //cell.Style.DateFormat.Format = "yyyy-MM-dd";
                            var hell = (DateTime)cell.Value;
                            var hell1 = cell.GetFormattedString();
                            var hell2 = cell.GetString();
                            DateTime dateTime1 = (DateTime)cell.Value;
                            DateTime dateTime2 = cell.GetDateTime();
                            DateTime dateTime3 = cell.GetValue<DateTime>();
                            String val = cell.RichText.ToString().Replace("@", "");
                            text.Append(val + " ");
                        }
                        else
                        {
                            text.Append(cell.RichText.ToString() + " ");
                        }

                    }
                }
        }

        public String getExcelSheetTextDemo(IXLWorksheet sheet)
        {
            try
            {
                var rows = sheet.RowsUsed();
                StringBuilder text = new StringBuilder("");
                foreach (var row in rows)
                {
                    //遍历所有的Cells
                    foreach (var cell in row.Cells())
                    {
                        Console.WriteLine(cell.Value + " " + cell.RichText);
                        //cell.SetDataType(XLCellValues.Text);
                        //Console.WriteLine(cell.Value+" "+cell.RichText);
                        if (cell.DataType == XLCellValues.DateTime)
                        {
                            String val = cell.RichText.ToString();
                            text.Append(val + " ");
                        }
                        else
                        {
                            text.Append(cell.RichText.ToString() + " ");
                        }

                    }
                }
                Console.WriteLine(text);
                return Regex.Replace(text.ToString(), "\\s+", " ");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
