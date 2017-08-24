using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using SolidFramework.Converters;
using SolidFramework.Converters.Plumbing;
using SolidFramework.Pdf;
using SolidFramework.Pdf.Transformers;
using SolidFramework.Plumbing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApplication1.util;

namespace WindowsFormsApplication1
{
    public partial class TestForm : Form
    {
        public static String remoteRoot = "/data/dearMrLei/data/subscriptions/";
        public static String localRoot = "D:\\test\\pdf\\";
        public static String chiPath = "D:\\tesseract\0823";
       
        public TestForm()
        {
            InitializeComponent();
           
        }


        private void button1_Click(object sender, EventArgs e)
        {
            SolidFramework.License.Import(@"d:\User\license.xml");


            OpenFileDialog OpFile = new OpenFileDialog();
            //show only PDF Files 
            OpFile.Filter = "PDF Files (*.pdf)|*.pdf";
            if (OpFile.ShowDialog() == DialogResult.OK)
            {
                String pdfPath = OpFile.FileName;
                Console.WriteLine(".................start");
                Console.WriteLine(pdfPath);
               
                String htmlFile = Path.ChangeExtension(pdfPath, ".html");
                Console.WriteLine(htmlFile);
                using (PdfToHtmlConverter converter = new PdfToHtmlConverter())
                {
                    //Add the selected file to the converter 
                    converter.AddSourceFile(pdfPath);

                    //Convert any graphics to Image Files 
                    converter.GraphicsAsImages = true;

                    //set the file type for the images 
                    converter.ImageType = ImageDocumentType.Jpeg;

                    //Automatically detect page orientation and rotate if required 
                    converter.AutoRotate = true;

                    //Convert the file 
                    converter.ConvertTo(htmlFile, true);
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SolidFramework.License.Import(@"d:\User\license.xml");
            //SolidFramework.Imaging.Ocr.SetTesseractDataDirectory(chiPath);
            OpenFileDialog OpFile = new OpenFileDialog();
            //show only PDF Files 
            OpFile.Filter = "PDF Files (*.pdf)|*.pdf";
            if (OpFile.ShowDialog() == DialogResult.OK)
            {
                //Define Two Strings to capture the selection and saving of your file 
                string pdfPath = OpFile.FileName;
                string searchablePdfPath = Path.ChangeExtension(pdfPath, "searchable.pdf");
                
                using (PdfDocument document = new PdfDocument(pdfPath))
                {
                    //Create a new OCRTransformer Object 
                    OcrTransformer transformer = new OcrTransformer();
                    
                    //Set the OcrType to Create a Searchable TextLayer
                    transformer.OcrType = OcrType.CreateSearchableTextLayer;

                    //Set the OCR Language to the Language in your PDF File - "en" for English, "es" for Spanish etc.
                    transformer.OcrLanguage = "zh";

                    //Preserve the Original PDF Files Image Compression
                    transformer.OcrImageCompression = SolidFramework.Imaging.Plumbing.ImageCompression.PreserveOriginal;

                    //Add the user selected PDF file to your transformer 
                    transformer.AddDocument(document);

                    //Transform the PDF File 
                    transformer.Transform();

                    //Save the new Searchable PDF file to the extension you specified earlier 
                    document.SaveAs(searchablePdfPath, OverwriteMode.ForceOverwrite);
                    Console.WriteLine(".................end");
                }
                DateTime d2 = System.DateTime.Now;
                Console.WriteLine(d2);

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Console.WriteLine(".................start");
            SolidFramework.License.Import(@"d:\User\license.xml");
            OpenFileDialog OpFile = new OpenFileDialog();
            //show only PDF Files 
            OpFile.Filter = "PDF Files (*.pdf)|*.pdf";
            if (OpFile.ShowDialog() == DialogResult.OK)
            {
                String pdfPath = OpFile.FileName;
                String xlsFile = Path.ChangeExtension(pdfPath, "xls");
                Console.WriteLine(xlsFile);
                using (PdfToExcelConverter converter = new PdfToExcelConverter())
                {
                    // Add files to convert. 
                    converter.AddSourceFile(pdfPath);
                    //Set the preferred conversion properties 
                    //converter.OutputType = ExcelDocumentType.Xls;

                    //This combines all tables onto one sheet 
                    converter.SingleTable = ExcelTablesOnSheet.PlaceEachTableOnOwnSheet;

                    //This gets Non Table Content 
                    converter.TablesFromContent = false;

                    //convert the file, calling it the same name but with a different extention , setting overwrite to true 
                    converter.ConvertTo(xlsFile, true);
                    Console.WriteLine(".................end");
                }
            }
        }
        //测试word文档
        private void button4_Click(object sender, EventArgs e)
        {
            Console.WriteLine(".................start");
            SolidFramework.License.Import(@"d:\User\license.xml");
            OpenFileDialog OpFile = new OpenFileDialog();
            //show only PDF Files 
            OpFile.Filter = "PDF Files (*.pdf)|*.pdf";
            if (OpFile.ShowDialog() == DialogResult.OK)
            {
                String pdfPath = OpFile.FileName;
                String wordPath = Path.ChangeExtension(pdfPath, ".docx");
                Console.WriteLine(wordPath);
                String tessdataFolder = "D:\\tesseract";
                SolidFramework.Imaging.Ocr.SetTesseractDataDirectory(tessdataFolder);
                using (PdfToWordConverter converter = new PdfToWordConverter())
                {
                    converter.TextRecoveryLanguage = "zh";
                    //Add the selected file 
                    converter.AddSourceFile(pdfPath);

                    //Continuous mode recovers text formatting, graphics and text flow 
                    //converter.ReconstructionMode = ReconstructionMode.Continuous;

                    //Or Use Flowing Reconstruction Mode if you need to keep the look and feel of the PDF 
                    converter.ReconstructionMode = ReconstructionMode.Flowing;


                    //To just convert the file with no message use only 
                    converter.ConvertTo(wordPath, true);
                    Console.WriteLine("end...");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SshConnectionInfo objInfo = new SshConnectionInfo();
            objInfo.User = "root";
            objInfo.Host = "106.75.3.227";
            //objInfo.IdentityFile = "password"; //有2中认证，一种基于PrivateKey,一种基于password
            objInfo.Pass = "Hu20160802Ben"; //基于密码
            SFTPHelper objSFTPHelper = new SFTPHelper(objInfo);
            DataBaseConnect dc = new DataBaseConnect();
            int index = 200;
            int limit = 2;
            String sql = "SELECT * from article where pdf_path != '' limit ";
            sql += index;
            sql += ",";
            sql += limit;
            Console.WriteLine(sql);
            MySqlDataReader reader = dc.getmysqlread(sql);
            //inv.Load(reader);
            while (reader.Read())
            {
                String pdf = (string)reader["pdf_path"];
                String real_pdf = remoteRoot + pdf;
                Console.WriteLine(real_pdf);
                int pdfIndex = pdf.IndexOf("/"); 
                String downloadPath = localRoot + pdf.Substring(pdfIndex+1);
                Console.WriteLine(downloadPath);
                objSFTPHelper.Download(real_pdf, downloadPath);
                pdfConvertWord(downloadPath);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SshConnectionInfo objInfo = new SshConnectionInfo();
            objInfo.User = "root";
            objInfo.Host = "106.75.3.227";
            //objInfo.IdentityFile = "password"; //有2中认证，一种基于PrivateKey,一种基于password
            objInfo.Pass = "Hu20160802Ben"; //基于密码
            SFTPHelper objSFTPHelper = new SFTPHelper(objInfo);
            //ArrayList list = objSFTPHelper.GetFileList("/data/dearMrLei/data/rich/2017/07/01");
            //string remotePath, string localPath
            String remotePath = "/data/dearMrLei/data/rich/2017/07/01/FZLzzR9H8uJCtGcqg.pdf";
            String localPath = "D:\\test\\pdf\\FZLzzR9H8uJCtGcqg.pdf";
            objSFTPHelper.Download(remotePath,localPath);
            //foreach (String str in list)
            //{
            //    Console.WriteLine(str);
            //}
           Console.WriteLine("download...");
           pdfConvertWord(localPath);

        }

        public void pdfConvertWord(String pdfPath) {
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
                SolidFramework.Forms.SolidMessageBox addDialog = new SolidFramework.Forms.SolidMessageBox(this);
                addDialog.Content = result.ToString();
                addDialog.Text = "Conversion Result";
                addDialog.ShowIcon = true;
                addDialog.MessageIcon = MessageBoxIcon.Information;
                //addDialog.Execute();

                //To just convert the file with no message use only 
                converter.ConvertTo(wordPath, true);
                Console.WriteLine("convert end...");
            }
        }

        private void upLoadbutton_Click(object sender, EventArgs e)
        {
            //将本地数据上传到服务器,然后如果存在的话返回false
            SshConnectionInfo objInfo = new SshConnectionInfo();
            objInfo.User = "root";
            objInfo.Host = "106.75.3.227";
            //objInfo.IdentityFile = "password"; //有2中认证，一种基于PrivateKey,一种基于password
            objInfo.Pass = "Hu20160802Ben"; //基于密码
            SFTPHelper objSFTPHelper = new SFTPHelper(objInfo);
            String local = localRoot + "2650939662_1.docx";
            String remote = "/data/dearMrLei/data/rich";
            String remoteFile = "/data/dearMrLei/data/rich/2650939662_1.docx";
            ArrayList list = objSFTPHelper.GetFileList(remoteFile);
            
            if (list != null && list.Count > 0)
            {
                Console.WriteLine("不是 0 ");
            }
            else
            {
                Console.WriteLine("是 0");
                objSFTPHelper.Upload(local, remote);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //测试多线程
            //将工作项加入到线程池队列中，这里可以传递一个线程参数
            //ThreadPool.QueueUserWorkItem(TestMethod, "1");
           // ThreadPool.QueueUserWorkItem(TestMethod, "2100");
            //ThreadPool.QueueUserWorkItem(TestMethod, "2000");
            ThreadPool.QueueUserWorkItem(TestMethod, "3000");
            ThreadPool.QueueUserWorkItem(TestMethod, "4000");
            //ThreadPool.QueueUserWorkItem(TestMethod, "5000");
            //ThreadPool.QueueUserWorkItem(TestMethod, "6000");
            //Console.Read();
            //ThreadStart childref = new ThreadStart(TestMethod);
            

        }

        public void TestMethod(object data)
        {
            string datastr = data as string;
            int index = int.Parse(datastr);
            Console.WriteLine(datastr);
            createdoc(index);
           
        }

        public void TestMethod2(object data)
        {
            string datastr = data as string;
            int index = int.Parse(datastr);
            Console.WriteLine(datastr);
            createdoc(index);

        }

        public static void TestMethod1(object data)
        {
            string datastr = data as string;
            Console.WriteLine(datastr);
            for (int i = 0; i < 1000; i++)
            {
                Console.WriteLine("method"+i);
                Thread.Sleep(200);
            }
            
            
        }

        private void createdoc(int index)
        {

            int limit = 100;
            SshConnectionInfo objInfo = new SshConnectionInfo();
            objInfo.User = "root";
            objInfo.Host = "106.75.3.227";
            //objInfo.IdentityFile = "password"; //有2中认证，一种基于PrivateKey,一种基于password
            objInfo.Pass = "Hu20160802Ben"; //基于密码
            SFTPHelper objSFTPHelper = new SFTPHelper(objInfo);
            DataBaseConnect dc = new DataBaseConnect();
            String sql = "SELECT * from article where pdf_path != '' limit ";
            sql += index;
            sql += ",";
            sql += limit;
            Console.WriteLine(sql);
            MySqlDataReader reader = dc.getmysqlread(sql);
            //inv.Load(reader);
            ArrayList list = new ArrayList();
            while (reader.Read())
            {
                list.Add((string)reader["pdf_path"]);
            }
            foreach (String pdf in list)
            {
                String real_pdf = remoteRoot + pdf;
                Console.WriteLine(real_pdf);
                int pdfIndex = pdf.IndexOf("/");
                String downloadPath = localRoot + pdf.Substring(pdfIndex + 1);
                Console.WriteLine(downloadPath);
                Boolean isSuccess = objSFTPHelper.Download(real_pdf, downloadPath);
                if (isSuccess)
                {
                    SolidConvertUtil solid = new SolidConvertUtil();
                    try
                    {
                        solid.pdfConvertWord(downloadPath);
                        solid.pdfConvertExcel(downloadPath);
                    }
                    catch
                    {
                        continue;
                    }

                }
            }


        }

        private void button8_Click(object sender, EventArgs e)
        {
            //Dao d = new Dao();
            //d.testspecilStr();
            Console.WriteLine("end.....");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            SolidFramework.License.Import(@"d:\User\license.xml");
            OpenFileDialog OpFile = new OpenFileDialog();
            //show only PDF Files 
            OpFile.Filter = "PDF Files (*.pdf)|*.pdf";

            if (OpFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                String pdfPath = OpFile.FileName;
                Console.WriteLine("{0}", pdfPath);
                String htmlPath = Path.ChangeExtension(pdfPath, ".html");
                String txtPath = Path.ChangeExtension(pdfPath, ".txt");
                DateTime d1 = System.DateTime.Now;
                Console.WriteLine(d1);
                TestTxt.SolidModelLayout(pdfPath, txtPath);
                DateTime d2 = System.DateTime.Now;
                Console.WriteLine(d2);
                Console.WriteLine(d1.Second - d2.Second);
            }
        }

        private void excelbutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpFile = new OpenFileDialog();
            //show only PDF Files 
            //OpFile.Filter = "PDF Files (*.xls)|*.xlsx";

            if (OpFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                String excelPath = OpFile.FileName;
                Console.WriteLine("{0}", excelPath);
               
                DateTime d1 = System.DateTime.Now;
                Console.WriteLine(d1);

                string convertxlsPath = Path.ChangeExtension(excelPath, "searchable.xls");
                //ExcelUtil.createExcel2(excelPath);

                //ExcelUtil.WriteExcel(dt,convertxlsPath);

                DateTime d2 = System.DateTime.Now;
                Console.WriteLine(d2);
                Console.WriteLine(d1.Second - d2.Second);
            }


        }

        private void toTxtButton_Click(object sender, EventArgs e)
        {
            LogHelper.WriteLog(typeof(TestForm), "测试Log4Net日志是否写入..............");
            OpenFileDialog OpFile = new OpenFileDialog();
            //show only PDF Files 
            //OpFile.Filter = "PDF Files (*.xls)|*.xlsx";
            if (OpFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                String path = OpFile.FileName;
                XLWorkbook workBook = new XLWorkbook(path);
                IXLWorksheet workSheet = workBook.Worksheet(1);
                var rows = workSheet.RowsUsed();
                StringBuilder text = new StringBuilder("");
                foreach (var row in rows)
                {
                    //遍历所有的Cells
                    foreach (var cell in row.Cells())
                    {
                        text.Append(cell.RichText.ToString()+" ");
                    }
                }
                String result = Regex.Replace(text.ToString(), "\\s+", " ");
                Console.WriteLine(result);
                Dao dao = new Dao();
                dao.testspecilStr(result);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Console.WriteLine(DateTimeUtil.GetTimeStamp());

        }
    }
}
