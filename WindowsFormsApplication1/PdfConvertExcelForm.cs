﻿using ClosedXML.Excel;
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
        private List<JobOrder> orders = new List<JobOrder>();

        private List<JobOrder> secondOrders = new List<JobOrder>();

        private object locker = new object();

        private int converterType = 5;
        private ReconstructionMode reconstructionMode;

        private SolidFramework.Services.JobProcessor processor;

        private int processedCount;
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
            
            buttonStop.Enabled = true;
            //获取数据
            long minId = dao.getMinId();
            long maxId = dao.getMaxId();
            long jianju = (maxId - minId) / 5;
            for (int a = 0; a < 5; a++)
            {
                long param1 = minId + (a * jianju);
                long param2 = minId + (a + 1) * jianju;
                Task.Factory.StartNew(() =>
                {
                    startThreadAddItem(param1 + "-" + param2);
                    handlePdf(param1 + "-" + param2);
                   
                });
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
                foreach (AnnouncementEntity ae in articleList) 
                {
                    //dealPdfConvertExcel(ae);
                }
                articleList = dao.getPdfStreamList(minid, maxid, LIMIT);
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
            buttonStop.Enabled = false;
        }

        private void dealPdfConvertExcel(String pdfPath)
        {
                try
                {
                    DateTime d1 = System.DateTime.Now;
                    Console.WriteLine("d1-" + d1);
                    String xlsFile = Path.ChangeExtension(pdfPath, "xlsx");
                    SolidConvertUtil solidConvertUtil = new SolidConvertUtil();
                    ConversionStatus result = solidConvertUtil.pdfConvertExcel2(pdfPath,xlsFile);
                    if (result == ConversionStatus.Success)
                    {
                        Boolean issaveSuccess = true;
                        //生成成功
                        if (File.Exists(xlsFile))
                        {
                            FileUtil.createDir(Path.GetDirectoryName(xlsFile));
                            String txtPath = Path.ChangeExtension(xlsFile, ".txt");
                            List<TableEntity> tbPostionList = TestTxt.SolidModelLayout(pdfPath, txtPath);
                            if (tbPostionList == null || tbPostionList.Count == 0)
                            {
                                return;
                            }
                            //对txt中的table进行合并
                            List<TableEntity> mergeTableList = mergeTable(tbPostionList);
                            
                            //获取多个sheet的文本
                            ExcelUtil eu = new ExcelUtil();
                            List<IXLWorksheet> sheetList = eu.getExcelSheetList(xlsFile);
                            int index = 0;
                            DateTime d2 = System.DateTime.Now;
                            Console.WriteLine("d2-" + d2);
                            for (int i=0;i<mergeTableList.Count;i++)
                            {
                                TableEntity tableEntity = mergeTableList[i];
                                if (tableEntity == null) 
                                {
                                    continue;
                                }
                                int txtLen = tableEntity.content.Replace(" ", "").Replace("\n", "").Length; //txt生成的表格文本长度
                                String excelTxt = eu.getExcelSheetText(sheetList[index]);
                                int excelTxtLen = excelTxt.Replace(" ", "").Replace("\n", "").Length;  //excel生成的文本长度
                                String excelContent = excelTxt.Replace("\n", "");
                                double rate = (double)txtLen / excelTxtLen;
                                rate = Math.Round(rate, 3);
                                string singleExcelPath = Path.ChangeExtension(xlsFile, index + ".xlsx");
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
                            DateTime d3 = System.DateTime.Now;
                            Console.WriteLine("d3-"+d3);
                            foreach (TableEntity tableEntity in mergeTableList)
                            {
                                if (tableEntity.flag == SysConstant.ERROR)
                                {
                                    issaveSuccess = false;
                                }
                                //dao.savePdfToExcelInfo(ae, tableEntity);
                            }
                            if (issaveSuccess)
                            {
                                
                                //startThreadAddItem(d1 + " update excel_flag :1 " + ae.doc_id);
                                //dao.updatePdfStreamInfo(ae, 1);
                                //将txt中的列提取出来
                                List<TxtEntity> paragraphList = getParagraph(tbPostionList, 11, 22);
                                foreach (TxtEntity txt in paragraphList) {
                                    Console.WriteLine(txt.content);
                                }
                            }
                            else
                            {
                                // startThreadAddItem(d1 + " update excel_flag :-10 " + ae.doc_id);
                               // dao.updatePdfStreamInfo(ae, -10);
                            }
                        }
                        Console.WriteLine("end............");
                    }
                    else 
                    {
                        //startThreadAddItem(Path.GetFileName(pdfPath) + " failed ");
                        //将pdf_stream表excel_flag标识改为 -1
                        //dao.updatePdfStreamInfo(ae, -(int)result);
                    }
                   
                }
                catch (Exception ex)
                {
                    //startThreadAddItem(ex.GetBaseException().Message);
                    //LogHelper.WriteLog(typeof(PdfConvertExcelForm), ex);
                }
                //结束生成
        }

        //分析表格是否需要合并
        public List<TxtEntity> getParagraph(List<TableEntity> tbPostionList, long docid, int doctype)
        {
            List<TxtEntity> paragraphList = new List<TxtEntity>();
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
                            paragraphList.Add(getTxtEntity(te, docid,doctype,false,2));
                            tableObj = te;
                            continue;
                        }
                        else 
                        {
                            paragraphList.Add(getTxtEntity(te, docid, doctype, false, 2));
                            tableObj = te;
                            continue;
                        }

                    }
                    else if (te.content_type != 2)  //等于段落
                    {
                        if (tableObj == null)  //忽略
                        {
                            paragraphList.Add(getTxtEntity(te, docid, doctype, false, 1));
                            continue;
                        }
                        else if (tableObj.bottom > te.bottom)
                        {
                            
                            paragraphList.Add(getTxtEntity(te, docid, doctype, true, 1));
                            tableObj.content += te.content;
                            continue;
                        }
                        else if (tableObj.bottom < te.bottom)
                        {
                            paragraphList.Add(getTxtEntity(te, docid, doctype, true, 1));
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
                            paragraphList.Add(getTxtEntity(te, docid, doctype, false, 2));
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
                            paragraphList.Add(getTxtEntity(te, docid, doctype, false, 1));
                            continue;
                        }
                        else
                        {
                            paragraphList.Add(getTxtEntity(te, docid, doctype, true, 1));
                            tableObj = null;
                            continue;
                        }
                    }
                }
            }

            return paragraphList;
        }

        public TxtEntity getTxtEntity(TableEntity te, long docid, int doctype, Boolean is_tb_content,int type)
        { 
        TxtEntity txt = new TxtEntity();
            txt.docid = docid;
            txt.doctype = doctype;
            txt.content = te.content;
            txt.top = te.top;
            txt.left = te.left;
            txt.right = te.right;
            txt.bottom = te.bottom;
            txt.content_id = te.content_id;
            if (is_tb_content)
            {
                txt.is_tb_content = 1;
            }
            else {
                txt.is_tb_content = 0;
            }
            
            txt.pageNumber = te.pageNumber;
            txt.type = type;
            return txt;
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
                //listBoxFiles.Items.Add(item);
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
                //listBoxFiles.Items.Clear();
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
                dealPdfConvertExcel(pdfPath);

            }
        }
        protected void convertExcel()
        {
            try
            {
                // Create jobs for each file and hand them to the job processor.
                foreach (JobOrder order in orders)
                {
                    PdfToExcelJobEnvelope job5 = new PdfToExcelJobEnvelope();
                    job5.DoProgress = true;
                    job5.SourcePath = PathUtil.getAbsolutePdfPath(order.Filename,order.stream.doc_type);
                    job5.Password = order.Password;
                    job5.CustomData = order.stream;
                    ListViewItem item5 = listView1.Items.Add(new ListViewItem(Path.GetFileName(job5.SourcePath)));
                    item5.Name = Path.GetFileName(job5.SourcePath);
                    item5.SubItems.Add("0");
                    item5.SubItems.Add("Queued");
                    //item5.SubItems.Add = job5.SourcePath;
                    job5.TablesFromContent = false;
                    job5.SingleTable = 0;
                    processor.SubmitJob(job5);
                }
                
            }
            catch (System.Exception ex)
            {
                string strMessage = string.Empty;
                if (ex.Message.Contains("correct license"))
                {
                    strMessage = "Please enter valid unlock information in About.";
                }
                else
                {
                    strMessage = ex.Message;
                }

                SolidFramework.Forms.SolidMessageBox messageDialog = new SolidFramework.Forms.SolidMessageBox(this);
                messageDialog.Content = strMessage;
                messageDialog.Text = "Error";
                messageDialog.MessageIcon = MessageBoxIcon.Error;
                messageDialog.Buttons = MessageBoxButtons.OK;
                messageDialog.ShowIcon = true;
                messageDialog.Execute();

                this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                processor.JobProgressEvent -= processor_JobProgressEvent;
                processor.JobCompletedEvent -= processor_JobCompletedEvent;
                processor.Close();
                processor.Dispose();
                this.Close();
            }
        }

        void processor_JobCompletedEvent(object sender, SolidFramework.Services.JobCompletedEventArgs e)
        {
            Invoke(new Action(() => this.UpdateCompletedStatus(e)));

            processedCount++;
            PdfStream steam = (PdfStream)e.JobEnvelope.CustomData;
            if (e.JobEnvelope.Status == JobStatus.Success)
            {
                try
                {
                        string convertedTemp = e.JobEnvelope.OutputPaths[0];
                        string pdfPath = e.JobEnvelope.SourcePath;
                        string ext = Path.GetExtension(convertedTemp);
                        string xlsFile = PathUtil.getAbsolutExcelPath(steam.pdf_path, steam.doc_type);
                        LogHelper.WriteLog(typeof(PdfConvertExcelForm), "xlsFile:" + xlsFile);
                        SolidFramework.Plumbing.Utilities.FileCopy(convertedTemp, xlsFile, true);
                        Boolean issaveSuccess = true;
                        //生成成功
                        if (File.Exists(xlsFile))
                        {
                            FileUtil.createDir(Path.GetDirectoryName(xlsFile));
                            String txtPath = Path.ChangeExtension(xlsFile, ".txt");
                            List<TableEntity> tbPostionList = TestTxt.SolidModelLayout(pdfPath, txtPath);
                            if (tbPostionList == null || tbPostionList.Count == 0)
                            {
                                steam.excel_flag = -10;
                                updatePdfStream(steam);
                                return;
                            }
                            //对txt中的table进行合并
                            List<TableEntity> mergeTableList = mergeTable(tbPostionList);
                            
                            //获取多个sheet的文本
                            ExcelUtil eu = new ExcelUtil();
                            List<IXLWorksheet> sheetList = eu.getExcelSheetList(xlsFile);
                            int index = 0;
                            for (int i = 0; i < mergeTableList.Count; i++)
                            {
                                TableEntity tableEntity = mergeTableList[i];
                                if (tableEntity == null)
                                {
                                    continue;
                                }
                                int txtLen = tableEntity.content.Replace(" ", "").Replace("\n", "").Length; //txt生成的表格文本长度
                                String excelTxt = "";
                                if (index < sheetList.Count)
                                {
                                    excelTxt = eu.getExcelSheetText(sheetList[index]);
                                }
                                else 
                                {
                                    //报错
                                    tableEntity.flag = SysConstant.ERROR;
                                    tableEntity.excelPath = "";
                                    break;
                                }
                                int excelTxtLen = excelTxt.Replace(" ", "").Replace("\n", "").Length;  //excel生成的文本长度
                                String excelContent = excelTxt.Replace("\n", "");
                                double rate = (double)txtLen / excelTxtLen;
                                rate = Math.Round(rate, 3);
                                string singleExcelPath = Path.ChangeExtension(xlsFile, index + ".xlsx");
                                if (rate < (1 + SysConstant.RANGE) && rate > (1 - SysConstant.RANGE)) //文本长度比例在 95%和105%之间
                                {
                                    eu.createExcelBySheet(sheetList[index], singleExcelPath);
                                    tableEntity.excelPath = singleExcelPath;
                                    tableEntity.flag = SysConstant.SUCCESS;
                                    tableEntity.content = excelContent;
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
                                        if (index < sheetList.Count)
                                        {
                                            excelTxt = eu.getExcelSheetText(sheetList[index]);
                                        }
                                        else
                                        {
                                            //报错
                                            tableEntity.flag = SysConstant.ERROR;
                                            tableEntity.excelPath = "";
                                            isError = true;
                                            break;
                                        }
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
                                    break;
                                }
                            }
                            if (issaveSuccess)
                            {
                                foreach (TableEntity tableEntity in mergeTableList)
                                {
                                    dao.savePdfToExcelInfo(steam, tableEntity);
                                }
                                steam.excel_flag = 1;
                                updatePdfStream(steam);
                                //将txt中的列提取出来
                                List<TxtEntity> paragraphList = getParagraph(tbPostionList, steam.doc_id, steam.doc_type);
                                foreach (TxtEntity txt in paragraphList)
                                {
                                    dao.savePdfTxtInfo(txt);
                                }

                            }
                            else
                            {
                                //dao.updatePdfStreamInfo(ae, -10);
                                steam.excel_flag = -10;  //对比错误
                                updatePdfStream(steam);
                            }
                        }
                        else
                        {
                            steam.excel_flag = -11;  //文件未找到
                            updatePdfStream(steam);
                        }
                       
                    
                }
                catch (Exception ex) {
                    LogHelper.WriteLog(typeof(PdfConvertExcelForm), "error-"+ex.GetBaseException().Message);
                    LogHelper.WriteLog(typeof(PdfConvertExcelForm), ex);
                }
            }
            else { //解析失败
                //将pdf_stream表excel_flag标识改为 -1
                steam.excel_flag = -(int)e.JobEnvelope.Status;
                updatePdfStream(steam);
            }

            //当完成时,在列队里加入一个任务
            addConvertJob();

        }
        public void addConvertJob() {
            PdfData pdfdata = HttpUtil.getPdfStreamDataByLimit(1);
            if (pdfdata != null && pdfdata.data != null &&pdfdata.data.Count > 0)
            {
                PdfToExcelJobEnvelope job5 = new PdfToExcelJobEnvelope();
                job5.DoProgress = true;
                LogHelper.WriteLog(typeof(PdfConvertExcelForm), "add convertJob..." + pdfdata.data[0].pdf_path);
                job5.SourcePath = PathUtil.getAbsolutePdfPath(pdfdata.data[0].pdf_path, pdfdata.data[0].doc_type);
                job5.Password = "";
                job5.CustomData = pdfdata.data[0];
                ListViewItem item5 = listView1.Items.Add(new ListViewItem(Path.GetFileName(job5.SourcePath)));
                item5.Name = Path.GetFileName(job5.SourcePath);
                item5.SubItems.Add("0");
                item5.SubItems.Add("Queued...");
                job5.TablesFromContent = false;
                job5.SingleTable = 0;
                processor.SubmitJob(job5);
            }
            
        }

        void processor_JobProgressEvent(object sender, SolidFramework.Services.JobProgressEventArgs e)
        {
            Invoke(new Action(() => this.UpdateProgress(e)));
        }

        private void UpdateProgress(SolidFramework.Services.JobProgressEventArgs e)
        {
            
            ListViewItem[] items = listView1.Items.Find(Path.GetFileName(e.JobEnvelope.SourcePath), false);
            if (items.Count() > 0)
            {
                int row = items[0].Index;

                //ProgressBar pb = listView1.GetEmbeddedControl(1, row) as ProgressBar;
                //pb.Value = e.Position;
                
                ListViewItem item = listView1.Items[row];
                item.SubItems[1].Text = e.Position+"%";
                if (item.SubItems[2].Text != "Processing")
                {
                    item.SubItems[2].Text = "Processing";
                }
            }
        }

        private void UpdateCompletedStatus(SolidFramework.Services.JobCompletedEventArgs e)
        {
            ListViewItem[] items = listView1.Items.Find(Path.GetFileName(e.JobEnvelope.SourcePath), false);
            if (items.Count() > 0)
            {
                int row = items[0].Index;

                ListViewItem item = listView1.Items[row];
                //ProgressBar pb = listView1.GetEmbeddedControl(1, row) as ProgressBar;
                if (e.JobEnvelope.Status != JobStatus.Started && e.JobEnvelope.Status != JobStatus.Success)
                {
                    item.SubItems[1].Text = "100%";
                }

                item.SubItems[2].Text = GetStatusString(e.JobEnvelope.Status);
            }
        }

        private string GetStatusString(JobStatus status)
        {
            string message = string.Empty;
            switch (status)
            {
                case JobStatus.BadData:
                case JobStatus.BadDataFailure:
                    message = "Bad Data";
                    break;
                case JobStatus.Cancelled:
                    message = "Cancelled";
                    break;
                case JobStatus.Created:
                    message = "Created";
                    break;
                case JobStatus.Failure:
                    message = "Failed";
                    break;
                case JobStatus.InternalErrorFailure:
                case JobStatus.InternalError:
                    message = "Internal Error";
                    break;
                case JobStatus.InvalidPassword:
                case JobStatus.InvalidPasswordFailure:
                    message = "Password Failure";
                    break;
                case JobStatus.NoImages:
                case JobStatus.NoImagesFailure:
                    message = "No Images";
                    break;
                case JobStatus.NoTables:
                case JobStatus.NoTablesFailure:
                    message = "No Tables";
                    break;
                case JobStatus.NoTagged:
                case JobStatus.NoTaggedFailure:
                    message = "No Tagging";
                    break;
                case JobStatus.NotPdfA:
                case JobStatus.NotPdfAFailure:
                    message = "Not PDF/A";
                    break;
                case JobStatus.TimedOut:
                case JobStatus.TimedOutFailure:
                    message = "Timed Out";
                    break;
                case JobStatus.Success:
                    message = "Success";
                    break;
                case JobStatus.Started:
                    message = "Started";
                    break;
                default:
                    message = "Unknown Error";
                    break;
            }

            return message;
        }

        public void initData(List<PdfStream> pdfStreamList)
        {
            foreach(PdfStream stream in pdfStreamList){
                JobOrder order = new JobOrder();
                order.Filename = stream.pdf_path;
                order.stream = stream;
                orders.Add(order);
            }
            reconstructionMode = ReconstructionMode.Flowing;

            // Setup the Solid Framework Job Processor.
            processor = new SolidFramework.Services.JobProcessor();

            // We will allow the job processor process to run as 64 bit on X64 OS
            // even it application is 32 bit. This allows faster processing, and we
            // have access to more memory.
            processor.Allow64on32 = false;

            processor.JobProgressEvent += processor_JobProgressEvent;
            processor.JobCompletedEvent += processor_JobCompletedEvent;

            processedCount = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SolidFramework.License.Import(@"c:\User\license.xml");
            while (true)
            {
                try
                {

                    PdfData pdfData = HttpUtil.getPdfStreamDataByLimit(30);
                    if (pdfData != null && pdfData.data != null && pdfData.data.Count > 0)
                    {
                        LogHelper.WriteLog(typeof(PdfConvertExcelForm), "start next convert");
                        initData(pdfData.data);
                        convertExcel();
                        processor.WaitTillComplete();
                        processor.JobProgressEvent -= processor_JobProgressEvent;
                        processor.JobCompletedEvent -= processor_JobCompletedEvent;
                        processor.Close();
                        processor.Dispose();
                        listView1.Items.Clear();
                        orders.Clear();
                        LogHelper.WriteLog(typeof(PdfConvertExcelForm), "end convert");
                    }
                    else
                    {
                        LogHelper.WriteLog(typeof(PdfConvertExcelForm), "sleep 3 m...");
                        Thread.Sleep(3000);
                    }
                }
                catch (Exception ex) {
                    LogHelper.WriteLog(typeof(PdfConvertExcelForm), "error...");
                    LogHelper.WriteLog(typeof(PdfConvertExcelForm),ex);
                }
            }
        }

        public void updatePdfStream(PdfStream stream) 
        {
            List<PdfStream> pdfdata = new List<PdfStream>();
            pdfdata.Add(stream);
            HttpUtil.updatePdfStreamData(pdfdata);
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            SolidFramework.License.Import(@"d:\User\license.xml");
            //HttpUtil.getPdfStreamDataByLimit(1);
            String pdfPath = @"C:\Users\Administrator\Desktop\9\16\506190882934.PDF";
            String txtPath = @"C:\Users\Administrator\Desktop\9\16\506190882934.txt";
            List<TableEntity> tbPostionList = TestTxt.SolidModelLayout(pdfPath, txtPath);
            //将txt中的列提取出来
            Console.WriteLine("..............");
            List<TxtEntity> paragraphList = getParagraph(tbPostionList,1000,2);

            foreach (TxtEntity txt in paragraphList) {
                dao.savePdfTxtInfo(txt);
            }
        }
        
    }
}
