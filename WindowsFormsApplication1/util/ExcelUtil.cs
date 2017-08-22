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
        public static void Import(string strFileName)
        {
            String path1 = @"C:\Users\Administrator\Desktop\8\21\549479070614_11.xlsx";
            String path2 = @"C:\Users\Administrator\Desktop\8\PDFConvertXls\547891379459.xlsx";
            var wb = new XLWorkbook(path2);
            //var wsSource = wb.Worksheet(1);
            // Copy the worksheet to a new sheet in this workbook
            

            // We're going to open another workbook to show that you can
            // copy a sheet from one workbook to another:
            String newPath = Path.ChangeExtension(strFileName, "searchable.xlsx");
            var wbSource = new XLWorkbook(path1);
            wbSource.Worksheet(2).CopyTo(wb, "Copy From Other");
            //wb.Worksheet(0).Delete();
            // Save the workbook with the 2 copies
            //wb.SaveAs("CopyingWorksheets.xlsx");
            wb.Save(true);

          

           
        }

        public static void createExcel(string excelFileName)
        {
            //List excelList = new List();
            String path2 = @"C:\Users\Administrator\Desktop\8\PDFConvertXls\547891379459.xlsx";
            var wb = new XLWorkbook(path2);
            var excelWork = new XLWorkbook(excelFileName);
            int size = excelWork.Worksheets.Count;
            for (int i = 1; i <=size; i++)
            {
                excelWork.Worksheet(i).CopyTo(wb, "table"+i);
                string tempPath = Path.ChangeExtension(excelFileName, i + ".xlsx");
                Console.WriteLine(tempPath);
                try
                {
                    wb.SaveAs(tempPath);

                }
                catch (Exception e) 
                { 

                }
                
            }
            //wb.SaveAs("CopyingWorksheets.xlsx");
            //wb.Save(true);
        }

        public static void createExcel1(string excelFileName)
        {
            //List excelList = new List();
          
            var excelWork = new XLWorkbook(excelFileName);
            int size = excelWork.Worksheets.Count;
            for (int i = 2; i <= size; i++)
            {
                excelWork.Worksheet(i).Delete();
                try
                {
                    if (i == size) {
                        string tempPath = Path.ChangeExtension(excelFileName, i + ".xlsx");
                        Console.WriteLine(tempPath);
                        excelWork.SaveAs(tempPath);
                    }
                }
                catch (Exception e)
                {

                }

            }
            //wb.SaveAs("CopyingWorksheets.xlsx");
            //wb.Save(true);
        }


       
    public static void createExcel2(string excelFileName)
    {
        var wb = new XLWorkbook("CopyingRanges.xlsx");
        var wsSource1 = wb.Worksheet(1);
        // Copy the worksheet to a new sheet in this workbook
        wsSource1.CopyTo("Copy");

        // We're going to open another workbook to show that you can
        // copy a sheet from one workbook to another:
        var wbSource = new XLWorkbook("CopyingWorksheets.xlsx");
        wbSource.Worksheet(1).CopyTo(wb, "Copy From Other1", 1);

        // Save the workbook with the 2 copies
        wb.Save(true);
        
       
    }

    }
}
