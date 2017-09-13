using ClosedXML.Excel;
using SolidFramework.Converters;
using SolidFramework.Converters.Plumbing;
using SolidFramework.Model.Layout;
using SolidFramework.Services;
using SolidFramework.Services.Plumbing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApplication1.constant;
using WindowsFormsApplication1.entity;
using WindowsFormsApplication1.util;

namespace WindowsFormsApplication1
{
    public partial class PdfConvertExcelForm : Form
    {
        //set your source and output folders 
        //static String sourceFolder = @"D:\process\source";
        //static String outputFolder = @"D:\process\output";
        //static String errorFolder = @"D:\process\error";
        static String sourceFolder = @"X:/juyuan_data/";
        static String outputFolder = @"X:\excel\";
        static Boolean startFlag = true;
        static int LIMIT = 50;
        public Dao dao = new Dao();
        public PdfConvertExcelForm()
        {
            InitializeComponent();

        }
        //开始转换
        private void buttonChoose_Click(object sender, EventArgs e)
        {
            SolidFramework.License.Import(@"d:\User\license.xml");
            buttonStart.Enabled = false;
            buttonStop.Enabled = true;
            //获取数据
            long minId = dao.getMinId();
            long maxId = dao.getMaxId();
            long jianju = (maxId - minId) / 15;
            for (int a = 0; a < 15; a++)
            {
                long param1 = minId + (a * jianju);
                long param2 = minId + (a + 1) * jianju;
                ThreadPool.QueueUserWorkItem(handlePdf, param1 + "-" + param2);
                startThreadAddItem(param1 + "-" + param2);
            }
        }

        public void handlePdf(Object str)
        {
            String[] param = str.ToString().Split('-');
            long minid = int.Parse(param[0]);
            long maxid = int.Parse(param[1]);
            List<AnnouncementEntity> articleList = dao.getPdfStreamList(minid, maxid, LIMIT);
            while (articleList != null && articleList.Count>0)
            {
                int pageSize = 2;
                int tatolPage = 1;
                if (articleList.Count % 2 == 0)
                {
                    tatolPage = articleList.Count / 2;
                }
                else
                {
                    tatolPage = (articleList.Count / 2) + 1;
                }
                for (int i = 1; i <= tatolPage; i++)
                {
                    List<AnnouncementEntity> result = articleList.Skip(pageSize * (i - 1)).Take(pageSize).ToList();
                    dealPdfConvertExcel(result);
                }
                articleList = dao.getPdfStreamList(minid, maxid, LIMIT);
                //startThreadAddItem(minid + "-" + (LIMIT));
            }
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            startThreadremoveItem();
            //listBoxFiles.Focus();
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            startFlag = false;
            buttonStart.Enabled = true;
            buttonStop.Enabled = false;
        }

        private void dealPdfConvertExcel(List<AnnouncementEntity> articleList)
        {
            using (JobProcessor processor = new JobProcessor())
            {
                try
                {
                    processor.KeepJobs = true;
                    foreach (AnnouncementEntity ae in articleList)
                    {
                        LogHelper.WriteLog(typeof(PdfConvertExcelForm), "start convert....");
                        String pdfPath = sourceFolder + ae.pdfPath.Replace("GSGGFWB/", "");
                        LogHelper.WriteLog(typeof(PdfConvertExcelForm), pdfPath);
                        //listBoxFiles.Items.Add(pdfPath);
                        if (!File.Exists(pdfPath))
                        {
                            dao.updatePdfStreamInfo(ae, -17);
                            continue;
                        }
                        startThreadAddItem(pdfPath);
                        String excelpath = pdfPath.Replace("juyuan_data", "excel/GSGGFWB");
                        excelpath = Path.ChangeExtension(excelpath, "xlsx");
                        
                        PdfToExcelJobEnvelope jobEnvelope = new PdfToExcelJobEnvelope();
                        //Set the Source Path 
                        jobEnvelope.SourcePath = pdfPath;
                        jobEnvelope.CustomData = ae;
                        jobEnvelope.SingleTable = 0;
                        jobEnvelope.TablesFromContent = false;

                        //Submit the Job 
                        processor.SubmitJob(jobEnvelope);
                    }
                    // wait until the queue is empty and all jobs are processed 
                    processor.WaitTillComplete();

                    foreach (JobEnvelope processedJob in processor.ProcessedJobs)
                    {
                        Boolean issaveSuccess = true;
                        DateTime d1 = System.DateTime.Now;
                        AnnouncementEntity announcement = (AnnouncementEntity)processedJob.CustomData;
                        if ((processedJob.Status != SolidFramework.Services.Plumbing.JobStatus.Success) || (processedJob.OutputPaths.Count != 1))
                        {
                            startThreadAddItem(d1 + ":" + Path.GetFileName(processedJob.SourcePath) + " failed because " + processedJob.Message);
                            //将pdf_stream表excel_flag标识改为 -1
                            int result = (int)processedJob.Status;
                            dao.updatePdfStreamInfo(announcement, -result);

                        }
                        else
                        {   //生成成功
                            String excelpath = processedJob.SourcePath.Replace("juyuan_data", "excel/GSGGFWB");
                            String wordTemporaryPath = processedJob.OutputPaths[0];
                            startThreadAddItem("temp file" + wordTemporaryPath);
                            String outputExtension = Path.GetExtension(wordTemporaryPath);
                            excelpath = Path.ChangeExtension(excelpath, outputExtension);
                            LogHelper.WriteLog(typeof(PdfConvertExcelForm), "create temp path " + excelpath);
                            startThreadAddItem("create temp path " + excelpath);
                            if (File.Exists(wordTemporaryPath))
                            {
                                startThreadAddItem("enter convert function ");
                                FileUtil.createDir(Path.GetDirectoryName(excelpath));
                                File.Copy(wordTemporaryPath, excelpath, true);

                                String txtPath = Path.ChangeExtension(excelpath, ".txt");
                                List<TableEntity> tbPostionList = TestTxt.SolidModelLayout(processedJob.SourcePath, txtPath);
                                if (tbPostionList == null || tbPostionList.Count == 0) 
                                {
                                    startThreadAddItem(d1 + " update excel_flag :-10 " + announcement.doc_id);
                                    dao.updatePdfStreamInfo(announcement, -10);
                                    break;
                                }
                                startThreadAddItem("tbPostionList:" + tbPostionList.Count);
                                //对txt中的table进行合并
                                List<TableEntity> mergeTableList = mergeTable(tbPostionList);
                                //获取多个sheet的文本
                                ExcelUtil eu = new ExcelUtil();
                                List<IXLWorksheet> sheetList = eu.getExcelSheetList(excelpath);
                                int index = 0;
                                foreach (TableEntity tableEntity in mergeTableList)
                                {
                                    int txtLen = tableEntity.content.Replace(" ", "").Replace("\n", "").Length; //txt生成的表格文本长度
                                    String excelTxt = eu.getExcelSheetText(sheetList[index]);
                                    int excelTxtLen = excelTxt.Replace(" ", "").Replace("\n", "").Length;  //excel生成的文本长度
                                    String excelContent = excelTxt.Replace("\n", "");
                                    double rate = (double)txtLen / excelTxtLen;
                                    rate = Math.Round(rate, 3);
                                    string singleExcelPath = Path.ChangeExtension(excelpath, index + ".xlsx");
                                    if (rate < (1 + SysConstant.RANGE) && rate > (1 - SysConstant.RANGE)) //文本长度比例在 95%和105%之间
                                    {
                                        //startThreadAddItem("succes:" + excelContent);
                                        eu.createExcelBySheet(sheetList[index], singleExcelPath);
                                        tableEntity.excelPath = singleExcelPath;
                                        tableEntity.flag = SysConstant.SUCCESS;
                                        tableEntity.content = excelContent;
                                        //resultList.Add(tableEntity);
                                    }
                                    else if (rate <= (1 - SysConstant.RANGE))  //文本长度比例小于等于95%
                                    {
                                        //报错
                                        tableEntity.flag = SysConstant.ERROR;
                                        tableEntity.excelPath = "";
                                        break;
                                    }
                                    else if (rate >= (1 + SysConstant.RANGE))  //文本长度比例大于等于105%
                                    {
                                        int totalLen = excelTxtLen;
                                        int initIndex = index;
                                        Boolean isError = false;
                                        while (true)
                                        {
                                            //合并当前sheet跟下一个sheet
                                            index++;
                                            int secondSheetTxtLen = eu.getExcelSheetText(sheetList[index]).Replace(" ", "").Replace("\n", "").Length;
                                            excelContent += eu.getExcelSheetText(sheetList[index]).Replace("\n", "");
                                            totalLen += secondSheetTxtLen;
                                            double rate1 = (double)txtLen / totalLen;
                                            rate1 = Math.Round(rate1, 3);
                                            if (rate1 < (1 + SysConstant.RANGE) && rate1 > (1 - SysConstant.RANGE)) //文本长度比例在 95%和105%之间
                                            {
                                                //合并excel
                                                eu.createExcelBySheetList(sheetList, singleExcelPath, initIndex, index);
                                                tableEntity.flag = SysConstant.SUCCESS;
                                                tableEntity.excelPath = singleExcelPath;
                                                tableEntity.content = excelContent;
                                                break;
                                            }
                                            else if (rate1 <= (1 - SysConstant.RANGE))  //文本长度比例小于等于95%
                                            {
                                                //报错
                                                tableEntity.flag = SysConstant.ERROR;
                                                tableEntity.excelPath = "";
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
                                foreach (TableEntity tableEntity in mergeTableList)
                                {
                                    if (tableEntity.flag == SysConstant.ERROR)
                                    {
                                        issaveSuccess = false;
                                    }
                                    dao.savePdfToExcelInfo(announcement, tableEntity);
                                }
                                if (issaveSuccess)
                                {
                                    startThreadAddItem(d1 + " update excel_flag :1 " + announcement.doc_id);
                                    dao.updatePdfStreamInfo(announcement, 1);
                                }
                                else
                                {
                                    startThreadAddItem(d1 + " update excel_flag :-10 " + announcement.doc_id);
                                    dao.updatePdfStreamInfo(announcement, -10);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    startThreadAddItem(ex.Message);
                    LogHelper.WriteLog(typeof(PdfConvertExcelForm), ex);
                }
                //结束生成
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
                    else if (te.content_type != 2)  //等于段落
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
                    else if (te.content_type != 2)
                    {

                        if (tableObj == null)  //忽略
                        {
                            continue;
                        }
                        else
                        {
                            mergeTableList.Add(tableObj);
                            tableObj = null;
                            continue;
                        }
                    }
                }
            }
            if (tableObj != null) 
            {
                mergeTableList.Add(tableObj);
            }
            
            return mergeTableList;
        }
        
        //分析表格是否需要合并
        public void mergeTabledemo(List<TableEntity> tbPostionList) 
        {
            List<TableEntity> mergeTableList = new List<TableEntity>();
            TableEntity tableObj = null;
            foreach (TableEntity te in tbPostionList) 
            {
                StringBuilder tableContent = new StringBuilder("");
                if (te.content_type == 2)  //类型等于table时 
                {
                    tableObj = te;
                    float botton = tableObj.bottom;  //表格右下角的位置
                    int contentId = tableObj.content_id;
                    //使用递归
                    recursMergeTable(tableObj, tbPostionList);
                    mergeTableList.Add(tableObj);
                }       
            }
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
                    recursMergeTable(tableObj, tbPostionList);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String path = "W:\\juyuan_data\\2017\\02\\hello.pdf";
            String savePath = path.Substring(path.IndexOf("02"));
            Console.WriteLine(savePath);
            listBoxFiles.Items.Add(savePath);
            String sd = DateTime.Now.Date.ToShortDateString();
            listBoxFiles.Items.Add(DateTime.Now.ToString());
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            TestForm myform = new TestForm();   //调用带参的构造函数
            myform.Show();
        }

        public void startThreadAddItem(String context)
        {
            Thread t2 = new Thread(new ParameterizedThreadStart(addBoxItem));
            t2.Start(context);
        }

        public void addBoxItem(Object item)
        {
            ThreadStart ts = delegate
            {
                listBoxFiles.Items.Add(item);
            };
            this.BeginInvoke(ts);
        }

        public void startThreadremoveItem()
        {
            Thread t3 = new Thread(removeBoxItem);
            t3.Start();
        }

        public void removeBoxItem()
        {
            ThreadStart ts = delegate
            {
                listBoxFiles.Items.Clear();
            };
            this.BeginInvoke(ts);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SolidFramework.License.Import(@"d:\User\license.xml");
            OpenFileDialog OpFile = new OpenFileDialog();

            if (OpFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                String pdfPath = OpFile.FileName;
                List<AnnouncementEntity> announceList = new List<AnnouncementEntity>();
                AnnouncementEntity ae = new AnnouncementEntity();
                ae.pdfPath = pdfPath;
                announceList.Add(ae);
                dealPdfConvertExcel(announceList);

            }


        }
    }
}
