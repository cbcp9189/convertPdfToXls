using SolidFramework.Converters;
using SolidFramework.Converters.Plumbing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsFormsApplication1.util;

namespace WindowsFormsApplication1
{
    class SolidConvertUtil
    {
        
        public void pdfConvertExcel(String path, String excelPath)
        {
            try
            {
                using (PdfToExcelConverter converter = new PdfToExcelConverter())
                {
                    // Add files to convert. 
                    converter.AddSourceFile(path);
                    
                    //This combines all tables onto one sheet 
                    converter.SingleTable = 0;

                    //This gets Non Table Content 
                    converter.TablesFromContent = false;

                    //convert the file, calling it the same name but with a different extention , setting overwrite to true 
                    converter.ConvertTo(excelPath, true);
                    Console.WriteLine("end............");
                    
                }

                

            }
            catch (Exception ex) 
            {
                Console.WriteLine(ex.Message);
            }
        }

        public ConversionStatus pdfConvertExcel2(String path,String xls)
        {
            
            Console.WriteLine(xls);
            try
            {
                using (PdfToExcelConverter converter = new PdfToExcelConverter())
                {
                    // Add files to convert. 
                    converter.AddSourceFile(path);

                    //This combines all tables onto one sheet 
                    converter.SingleTable = 0;

                    //This gets Non Table Content 
                    converter.TablesFromContent = false;
                    return converter.ConvertTo(xls, true);

                    //convert the file, calling it the same name but with a different extention , setting overwrite to true 
                    //return xls;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return ConversionStatus.Fail;
        }

        public void pdfConvertExcel(String path, String excelPath,PdfToExcelConverter converter)
        {
            try
            {
                    // Add files to convert. 
                    converter.AddSourceFile(path);

                    //This combines all tables onto one sheet 
                    converter.SingleTable = 0;

                    //This gets Non Table Content 
                    converter.TablesFromContent = false;

                    //convert the file, calling it the same name but with a different extention , setting overwrite to true 
                    converter.ConvertTo(excelPath, true);
                    Console.WriteLine("end............");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}
