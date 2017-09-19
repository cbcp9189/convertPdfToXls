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
                    dealPdfConvertExcel(ae);
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

        private void dealPdfConvertExcel(AnnouncementEntity ae)
        {
                try
                {
                    DateTime d1 = System.DateTime.Now;
                    LogHelper.WriteLog(typeof(PdfConvertExcelForm), "start convert....");
                    String pdfPath = sourceFolder + ae.pdfPath.Replace("GSGGFWB/", "");
                    //String pdfPath = ae.pdfPath.Replace("GSGGFWB/", "");
                    LogHelper.WriteLog(typeof(PdfConvertExcelForm), pdfPath);
                    if (!File.Exists(pdfPath))
                    {
                        dao.updatePdfStreamInfo(ae, -17);
                    }
                    startThreadAddItem(pdfPath);
                    String xlsFile = Path.ChangeExtension(pdfPath.Replace("juyuan_data", "excel/GSGGFWB"), "xlsx");
                    SolidConvertUtil solidConvertUtil = new SolidConvertUtil();
                    ConversionStatus result = solidConvertUtil.pdfConvertExcel(pdfPath, xlsFile);
                    if (result == ConversionStatus.Success)
                    {
                        Boolean issaveSuccess = true;
                        //生成成功
                        if (File.Exists(xlsFile))
                        {
                            startThreadAddItem("enter convert function...");
                            FileUtil.createDir(Path.GetDirectoryName(xlsFile));
                            String txtPath = Path.ChangeExtension(xlsFile, ".txt");
                            List<TableEntity> tbPostionList = TestTxt.SolidModelLayout(pdfPath, txtPath);
                            if (tbPostionList == null || tbPostionList.Count == 0)
                            {
                                startThreadAddItem(d1 + ": update excel_flag :-10 " + ae.doc_id);
                                dao.updatePdfStreamInfo(ae, -10);
                                return;
                            }
                            startThreadAddItem("tbPostionList:" + tbPostionList.Count);
                            //对txt中的table进行合并
                            List<TableEntity> mergeTableList = mergeTable(tbPostionList);
                            //获取多个sheet的文本
                            ExcelUtil eu = new ExcelUtil();
                            List<IXLWorksheet> sheetList = eu.getExcelSheetList(xlsFile);
                            int index = 0;
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
                                startThreadAddItem(d1 + " update excel_flag :1 " + ae.doc_id);
                                dao.updatePdfStreamInfo(ae, 1);
                            }
                            else
                            {
                                startThreadAddItem(d1 + " update excel_flag :-10 " + ae.doc_id);
                                dao.updatePdfStreamInfo(ae, -10);
                            }
                        }
                    }
                    else 
                    {
                        startThreadAddItem(Path.GetFileName(pdfPath) + " failed ");
                        //将pdf_stream表excel_flag标识改为 -1
                        dao.updatePdfStreamInfo(ae, -(int)result);
                    }
                   
                }
                catch (Exception ex)
                {
                    startThreadAddItem(ex.GetBaseException().Message);
                    LogHelper.WriteLog(typeof(PdfConvertExcelForm), ex);
                }
                //结束生成
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
                AnnouncementEntity ae = new AnnouncementEntity();
                ae.pdfPath = pdfPath;
                dealPdfConvertExcel(ae);

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

                    if (e.JobEnvelope.GetType() == typeof(PdfToExcelJobEnvelope))
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
                            }
                            else
                            {
                                //dao.updatePdfStreamInfo(ae, -10);
                                steam.excel_flag = -10;
                                updatePdfStream(steam);
                            }
                        }
                    }
                    else
                    {
                        //将pdf_stream表excel_flag标识改为 -1
                        //dao.updatePdfStreamInfo(ae, -(int)result);
                        steam.excel_flag = -1;
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
            //addConvertJob();

        }
        public void addConvertJob() {
            lock (locker) {
                if (secondOrders == null || secondOrders.Count() == 0)
                {
                    //获取pdfstream信息
                    PdfData pdfData = HttpUtil.getPdfStreamData();
                    if (pdfData != null)
                    {
                        foreach (PdfStream stream in pdfData.data)
                        {
                            JobOrder order = new JobOrder();
                            order.Filename = stream.pdf_path;
                            order.stream = stream;
                            secondOrders.Add(order);
                        }
                        if (secondOrders.Count > 0)
                        {
                            JobOrder tempOrder = secondOrders[0];
                            PdfToExcelJobEnvelope job5 = new PdfToExcelJobEnvelope();
                            job5.DoProgress = true;
                            job5.SourcePath = PathUtil.getAbsolutePdfPath(tempOrder.Filename, tempOrder.stream.doc_type);
                            job5.Password = tempOrder.Password;
                            job5.CustomData = tempOrder.stream;
                            ListViewItem item5 = listView1.Items.Add(new ListViewItem(Path.GetFileName(job5.SourcePath)));
                            item5.Name = Path.GetFileName(job5.SourcePath);
                            item5.SubItems.Add("0");
                            item5.SubItems.Add("Queued...");
                            //item5.SubItems.Add = job5.SourcePath;
                            job5.TablesFromContent = false;
                            job5.SingleTable = 0;
                            processor.SubmitJob(job5);
                            secondOrders.RemoveAt(0);
                        }
                    }
                }
                else
                {
                    JobOrder order = secondOrders[0];
                    PdfToExcelJobEnvelope job5 = new PdfToExcelJobEnvelope();
                    job5.DoProgress = true;
                    job5.SourcePath = PathUtil.getAbsolutePdfPath(order.Filename, order.stream.doc_type);
                    job5.Password = order.Password;
                    job5.CustomData = order.stream;
                    ListViewItem item5 = listView1.Items.Add(new ListViewItem(Path.GetFileName(job5.SourcePath)));
                    item5.Name = Path.GetFileName(job5.SourcePath);
                    item5.SubItems.Add("0");
                    item5.SubItems.Add("Queued...");
                    //item5.SubItems.Add = job5.SourcePath;
                    job5.TablesFromContent = false;
                    job5.SingleTable = 0;
                    processor.SubmitJob(job5);
                    secondOrders.RemoveAt(0);
                }
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

        public void initData(int convertType,List<PdfStream> pdfStreamList)
        {
            foreach(PdfStream stream in pdfStreamList){
                JobOrder order = new JobOrder();
                order.Filename = stream.pdf_path;
                order.stream = stream;
                orders.Add(order);
            }
            converterType = convertType;
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
            SolidFramework.License.Import(@"d:\User\license.xml");
            int type = 5;
            
            while (true)
            {
                PdfData pdfData = HttpUtil.getPdfStreamData();
                if (pdfData != null)
                {
                    initData(type, pdfData.data);
                    convertExcel();
                    processor.WaitTillComplete();
                    processor.JobProgressEvent -= processor_JobProgressEvent;
                    processor.JobCompletedEvent -= processor_JobCompletedEvent;
                    processor.Close();
                    processor.Dispose();
                    listView1.Items.Clear();
                    orders.Clear();
                }
                else {
                    Thread.Sleep(3000);
                }
            }
        }

        public void updatePdfStream(PdfStream stream) 
        {
            List<PdfStream> pdfdata = new List<PdfStream>();
            pdfdata.Add(stream);
            HttpUtil.updatePdfStreamData(pdfdata);
        }
        
    }
}
