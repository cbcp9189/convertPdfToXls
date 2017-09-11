using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using SolidFramework.Converters;
using SolidFramework.Converters.Plumbing;
using SolidFramework.Pdf;
using SolidFramework.Pdf.Transformers;
using SolidFramework.Plumbing;
using SolidFramework.Services;
using SolidFramework.Services.Plumbing;
using Spire.Xls;
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
using WindowsFormsApplication1.constant;
using WindowsFormsApplication1.entity;
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
                using (JobProcessor processor = new JobProcessor())
                {
                        processor.KeepJobs = true;
                        PdfToExcelJobEnvelope jobEnvelope = new PdfToExcelJobEnvelope();
                        jobEnvelope.SourcePath = pdfPath;
                        jobEnvelope.SingleTable = ExcelTablesOnSheet.PlaceEachTableOnOwnSheet;
                        jobEnvelope.TablesFromContent = false;
                        processor.SubmitJob(jobEnvelope);
                        processor.WaitTillComplete();
                        Thread.Sleep(500);

                    foreach (JobEnvelope processedJob in processor.ProcessedJobs)
                    {
                        if ((processedJob.Status != SolidFramework.Services.Plumbing.JobStatus.Success) || (processedJob.OutputPaths.Count != 1))
                        {
                            //Console.WriteLine(Path.GetFileName(processedJob.SourcePath) + " failed because " + processedJob.Message);
                            //listBoxFiles.Items.Add(Path.GetFileName(processedJob.SourcePath) + " failed because " + processedJob.Message);
                            LogHelper.WriteLog(typeof(PdfConvertExcelForm), Path.GetFileName(processedJob.SourcePath) + " failed because " + processedJob.Message);

                        }
                        else
                        {  //生成成功
                            String wordTemporaryPath = processedJob.OutputPaths[0];
                            String outputExtension = Path.GetExtension(wordTemporaryPath);
                            Console.WriteLine(wordTemporaryPath);
                            String excelpath = Path.ChangeExtension(processedJob.SourcePath, outputExtension);
                            LogHelper.WriteLog(typeof(PdfConvertExcelForm), "temp path " + excelpath);
                            // listBoxFiles.Items.Add(d1+" start convert..." + excelpath);
                            if (File.Exists(wordTemporaryPath))
                            {
                                FileUtil.createDir(Path.GetDirectoryName(excelpath));
                                File.Copy(wordTemporaryPath, excelpath, true);
                            }
                            Thread.Sleep(50);
                        }
                    }
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
                        //solid.pdfConvertExcel(downloadPath);
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
                //String htmlPath = Path.ChangeExtension(pdfPath, ".html");
                String txtPath = Path.ChangeExtension(pdfPath, ".txt");
                DateTime d1 = System.DateTime.Now;
                Console.WriteLine(d1);
                List<TableEntity> tbPostionList = TestTxt.SolidModelLayout(pdfPath, txtPath);
                
                //对txt中的table进行合并
                List<TableEntity> mergeTableList = mergeTable(tbPostionList);
                //生成excel
                String excelPath = SolidConvertUtil.pdfConvertExcel(pdfPath);
                //获取多个sheet的文本
                ExcelUtil eu = new ExcelUtil();
                List<IXLWorksheet> sheetList = eu.getExcelSheetList(excelPath);
                int index = 0;
                //List<TableEntity> resultList = new List<TableEntity>();
                foreach (TableEntity tableEntity in mergeTableList) 
                {
                    int txtLen = tableEntity.content.Replace(" ", "").Replace("\n","").Length; //txt生成的表格文本长度
                    String excelTxt = eu.getExcelSheetText(sheetList[index]);
                    int excelTxtLen = excelTxt.Replace(" ", "").Replace("\n", "").Length;  //excel生成的文本长度
                    
                    double rate = (double)txtLen / excelTxtLen;
                    rate = Math.Round(rate, 3);
                    string singleExcelPath = Path.ChangeExtension(excelPath, index + ".xlsx");
                    if (rate < (1 + SysConstant.RANGE) && rate > (1 - SysConstant.RANGE)) //文本长度比例在 95%和105%之间
                    {
                        eu.createExcelBySheet(sheetList[index], singleExcelPath);
                        tableEntity.excelPath = singleExcelPath;
                        tableEntity.flag = SysConstant.SUCCESS;
                        
                    }
                    else if (rate <= (1 - SysConstant.RANGE))  //文本长度比例小于等于95%
                    {
                        //报错
                        tableEntity.flag = SysConstant.ERROR;
                        Console.WriteLine("error....................");
                        continue;
                    }
                    else if (rate >= (1 + SysConstant.RANGE))  //文本长度比例大于等于105%
                    {
                        int totalLen = excelTxtLen;
                        int initIndex = index;
                        Boolean isError = false;
                        while (true) {
                            //合并当前sheet跟下一个sheet
                            index++;
                            int secondSheetTxtLen = eu.getExcelSheetText(sheetList[index]).Replace(" ", "").Replace("\n", "").Length;
                            totalLen += secondSheetTxtLen;
                            double rate1 = (double)txtLen / totalLen;
                            rate1 = Math.Round(rate1, 3);
                            if (rate1 < (1 + SysConstant.RANGE) && rate1 > (1 - SysConstant.RANGE)) //文本长度比例在 95%和105%之间
                            {
                                //合并excel
                                eu.createExcelBySheetList(sheetList, singleExcelPath,initIndex,index);
                                tableEntity.flag = SysConstant.SUCCESS;
                                break;
                            }
                            else if (rate1 <= (1 - SysConstant.RANGE))  //文本长度比例小于等于95%
                            {
                                //报错
                                tableEntity.flag = SysConstant.ERROR;
                                Console.WriteLine("error....................");
                                isError = true;
                                break;
                            }
                        }
                        if (isError) 
                        {
                            break;
                        }
                    }
                    index++;
                }
                DateTime d2 = System.DateTime.Now;
                Console.WriteLine(d2);
                Console.WriteLine(d1.Second - d2.Second);
            }
        }

        //分析表格是否需要合并
        public List<TableEntity> mergeTable(List<TableEntity> tbPostionList)
        {
            List<TableEntity> mergeTableList = new List<TableEntity>();
            TableEntity tableObj = null;
            int currentPage = 1;
            foreach (TableEntity te in tbPostionList)
            {
                if (te.pageNumber == currentPage)
                {
                    if (te.content_type == 2 )  //类型等于table时 
                    {
                        if (tableObj == null)
                        {
                            //currentPage = te.pageNumber;
                            tableObj = te;
                            continue;
                        }
                        else 
                        {
                            mergeTableList.Add(tableObj);
                            tableObj = te;
                            continue;
                        }

                    }
                    else if (te.content_type == 6)  //等于段落
                    {
                        if (tableObj == null)  //忽略
                        {
                            continue;
                        }
                        else if (tableObj.bottom > te.bottom)
                        {
                            tableObj.content += te.content;
                            continue;
                        }
                        else if (tableObj.bottom < te.bottom)
                        {
                            mergeTableList.Add(tableObj);
                            tableObj = null;
                            continue;
                        }
                    }
                    else{
                        continue;
                    }
                }
                else   //不等于当前页
                {
                    currentPage = te.pageNumber;

                    if (te.content_type == 2) 
                    {
                       
                        if (tableObj == null) //说明是新表
                        {
                           
                            tableObj = te;
                            continue;
                        }
                        else
                        {
                            //合并的逻辑
                            tableObj.bottom = te.bottom;
                            tableObj.right = te.right;
                            tableObj.pages++;
                            continue;
                        }
                        
                    }
                    else if (te.content_type == 6)
                    {

                        if (tableObj == null)  //忽略
                        {
                            continue;
                        }else
                        {
                            mergeTableList.Add(tableObj);
                            tableObj = null;
                            continue;
                        }

                    }
                    else
                    {
                        continue;
                    }
                }
            }
            if (tableObj != null) 
            {
                mergeTableList.Add(tableObj);
            }
            foreach (TableEntity t in mergeTableList) {
                Console.WriteLine(t.content);
                Console.WriteLine(t.pageNumber+"--"+t.pages);
                Console.WriteLine("-----------------");
            }
            Console.WriteLine("end..........");
            return mergeTableList;
        }

        public void recursMergeTable(TableEntity tableObj, List<TableEntity> tbPostionList)
        {
            float botton = tableObj.bottom;  //表格右下角的位置
            int contentId = tableObj.content_id;
            Boolean isFirstTable = false;
            int index = 0;
            foreach (TableEntity subTe in tbPostionList)
            {
                if (subTe.content_type == 6 && subTe.content_id > contentId)  //等于段落
                {
                    if (subTe.bottom <= tableObj.bottom && subTe.pageNumber == tableObj.pageNumber)
                    {
                        tableObj.content += subTe.content;
                    }
                    else
                    {
                        break;
                    }
                }
                else if (subTe.pageNumber > tableObj.pageNumber && subTe.content_type == 2 && index == 0)   //下一页的第一个元素是table
                {
                    isFirstTable = true;
                    index++;
                }
                else if (subTe.pageNumber > tableObj.pageNumber && subTe.content_type == 6 && index == 0)   //下一页的第一个元素是Paragraph
                {
                    isFirstTable = false;
                    index++;
                    break;
                }
                if (isFirstTable)
                {//进行合并
                    tableObj.content_id = subTe.content_id;
                    tableObj.pageNumber = subTe.pageNumber;
                    tableObj.bottom = subTe.bottom;
                    recursMergeTable(tableObj, tbPostionList);
                    break;
                }
            }
        }

        private void excelbutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpFile = new OpenFileDialog();
            //show only PDF Files 
            //OpFile.Filter = "PDF Files (*.xls)|*.xlsx";

            if (OpFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string pdfPath = OpFile.FileName;
                var wb = new XLWorkbook(pdfPath);
                var ws = wb.Worksheet(1);
               
                // Define a range with the data
                var firstTableCell = ws.FirstCellUsed();
                var lastTableCell = ws.LastCellUsed();
                var rngData = ws.Range(firstTableCell.Address, lastTableCell.Address);
                var lastCellAddress = ws.LastCellUsed().Address;
                
                // Copy the table to another worksheet
                //var wsCopy = wb.Worksheets.Add("Contacts Copy2");
                var wsCopy = wb.Worksheet(2);
                wsCopy.Style.Alignment = ws.Style.Alignment;
                wsCopy.Style.Border = ws.Style.Border;
                wsCopy.Style.Font = ws.Style.Font;
                wsCopy.Style.Fill = ws.Style.Fill;
                wsCopy.Style.NumberFormat = ws.Style.NumberFormat;
                wsCopy.Cell(7, 1).Value = rngData;
                //wsCopy.rows
                wb.SaveAs("Copying123Ranges.xlsx", true);
            }


        }

        private void toTxtButton_Click(object sender, EventArgs e)
        {
            //Excel的文档结构是 Workbook->Worksheet（1个book可以包含多个sheet）
            Workbook workbook = new Workbook();

            //获取第一个sheet，进行操作，下标是从0开始
            Worksheet sheet = workbook.Worksheets[0];


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
                        String str1 = cell.GetFormattedString();
                        if (cell.DataType == XLCellValues.DateTime)
                        {
                            if (cell.RichText.ToString().Contains("@"))
                            {
                                String val = cell.RichText.ToString().Replace("@", "");
                                text.Append(val+" ");
                            }
                        }
                        else 
                        {
                            text.Append(cell.RichText.ToString() + " ");
                        }
                        
                       
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
            //double rate = Convert.ToDouble(100) / Convert.ToDouble(105);
            OpenFileDialog OpFile = new OpenFileDialog();
            if (OpFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                String path = OpFile.FileName;
                ExcelUtil.TestExcel(path);
            }
            //Console.WriteLine(rate);

        }

        private void button11_Click(object sender, EventArgs e)
        {
            startThreadAddItem("hello world");

        }

        public void startThreadAddItem(String context)
        {
            Thread t2 = new Thread(new ParameterizedThreadStart(addBoxItem));
            t2.Start(context);
        }

        public void addBoxItem(Object item) {
            ThreadStart ts = delegate
            {
                listBox1.Items.Add(item);
            };
            this.BeginInvoke(ts);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpFile = new OpenFileDialog();

            if (OpFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string excelPath = OpFile.FileName;
                ExcelUtil eu = new ExcelUtil();
                List<IXLWorksheet> sheetList = eu.getExcelSheetList(excelPath);
                string tempPath = Path.ChangeExtension(excelPath, "hello.xlsx");
                eu.createExcelBySheetList(sheetList,tempPath, 1, 3);
            }
            

        }
    }
}
