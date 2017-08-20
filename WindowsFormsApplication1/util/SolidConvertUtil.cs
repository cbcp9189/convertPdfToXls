using SolidFramework.Converters;
using SolidFramework.Converters.Plumbing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class SolidConvertUtil
    {
        //pdf生成word
        public void pdfConvertWord(String pdfPath)
        {
            SolidFramework.License.Import(@"d:\User\license.xml");
            String wordPath = Path.ChangeExtension(pdfPath, ".docx");
            Console.WriteLine(wordPath);
            using (PdfToWordConverter converter = new PdfToWordConverter())
            {
                //Add the selected file 
                converter.AddSourceFile(pdfPath);

                //Continuous mode recovers text formatting, graphics and text flow 
                converter.ReconstructionMode = ReconstructionMode.Continuous;

                //Or Use Flowing Reconstruction Mode if you need to keep the look and feel of the PDF 
                converter.ReconstructionMode = ReconstructionMode.Flowing;

                // To catch conversion result and display it to the user use the following 
                ConversionStatus result = converter.ConvertTo(wordPath, true);
                //To just convert the file with no message use only 
                converter.ConvertTo(wordPath, true);
                Console.WriteLine("convert word end...");
            }
        }

        public void pdfConvertExcel(String path)
        {
            Console.WriteLine(".................start");
            SolidFramework.License.Import(@"d:\User\license.xml");
            String xlsFile = Path.ChangeExtension(path, "xlsx");
            Console.WriteLine(xlsFile);
            using (PdfToExcelConverter converter = new PdfToExcelConverter())
            {
                // Add files to convert. 
                converter.AddSourceFile(path);
                //Set the preferred conversion properties 

                //This combines all tables onto one sheet 
                converter.SingleTable = 0;

                //This gets Non Table Content 
                converter.TablesFromContent = false;

                //convert the file, calling it the same name but with a different extention , setting overwrite to true 
                converter.ConvertTo(xlsFile, true);
                Console.WriteLine(".................end");
            }

        }

    }
}
