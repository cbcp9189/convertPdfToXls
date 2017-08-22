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
        static String sourceFolder = @"W:/juyuan_data/";
        static String outputFolder = @"W:\excel\";
        static Boolean startFlag = true;
        static int LIMIT = 50;
        public  Dao dao = new Dao();
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
            Console.WriteLine(minId+"-"+maxId);
            long jianju = (maxId - minId) / 3;
            for (int a = 0; a < 3; a++)
            {
                long param1 = minId + (a * jianju);
                long param2 = minId + (a + 1) * jianju;
                
                ThreadPool.QueueUserWorkItem(handlePdf, param1 + "-" + param2);
                //handlePdf(param1 + "-" + param2);
            }
        }


        public void handlePdf(Object str) {
            String[] param  = str.ToString().Split('-');
            long minid = int.Parse(param[0]);
            long maxid = int.Parse(param[1]);
           
            while (startFlag && minid <= maxid)
            {
                listBoxFiles.Items.Add(minid + "--" + LIMIT);
                List<AnnouncementEntity> articleList = dao.getAnnouncementList(minid, minid+LIMIT);
                dealPdfConvertExcel(articleList);
                minid += LIMIT;
            }

        }
        private void buttonClear_Click(object sender, EventArgs e)
        {
            listBoxFiles.Items.Clear();
            
            listBoxFiles.Focus();
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
                processor.KeepJobs = true;
                foreach (AnnouncementEntity ae in articleList)
                {
                    String pdfPath = sourceFolder + ae.pdfPath.Replace("GSGGFWB/", "");
               
                    listBoxFiles.Items.Add(pdfPath);
                    String excelpath = pdfPath.Replace("juyuan_data", "excel/GSGGFWB");
                    excelpath = Path.ChangeExtension(excelpath, "xlsx");
                    listBoxFiles.Items.Add("ex-" + excelpath);
                    if (File.Exists(excelpath)) {
                        listBoxFiles.Items.Add(pdfPath+" excel File exist");
                        continue;
                    }
                    PdfToExcelJobEnvelope jobEnvelope = new PdfToExcelJobEnvelope();
                    //Set the Source Path 
                    jobEnvelope.SourcePath = pdfPath;
                    jobEnvelope.SingleTable = 0;
                    jobEnvelope.TablesFromContent = false;

                    //Submit the Job 
                    processor.SubmitJob(jobEnvelope);
                }
                // wait until the queue is empty and all jobs are processed 
                Console.WriteLine("before");
                processor.WaitTillComplete();
                Console.WriteLine("after");
                Thread.Sleep(1000);

                foreach (JobEnvelope processedJob in processor.ProcessedJobs)
                {
                    
                    if ((processedJob.Status != SolidFramework.Services.Plumbing.JobStatus.Success) || (processedJob.OutputPaths.Count != 1))
                    {
                        // report errors to the console window 
                        Console.WriteLine(Path.GetFileName(processedJob.SourcePath) + " failed because " + processedJob.Message);
                        listBoxFiles.Items.Add(Path.GetFileName(processedJob.SourcePath) + " failed because " + processedJob.Message);
                        //String eFolder = Path.Combine(errorFolder, Path.GetFileName(processedJob.SourcePath));
                        //String ePath = processedJob.SourcePath;
                        //File.Move(ePath, eFolder);
                        savePdfToExcelInfo(articleList, Path.GetFileName(processedJob.SourcePath), "");
                    }
                    else
                    {  //生成成功
                        String excelpath = processedJob.SourcePath.Replace("juyuan_data", "excel/GSGGFWB");
                        String wordTemporaryPath = processedJob.OutputPaths[0];
                        String outputExtension = Path.GetExtension(wordTemporaryPath);
                        excelpath = Path.ChangeExtension(excelpath, outputExtension);
                        listBoxFiles.Items.Add("start convert..." + excelpath);
                        if (File.Exists(wordTemporaryPath))
                        {
                            FileUtil.createDir(Path.GetDirectoryName(excelpath));
                            File.Copy(wordTemporaryPath, excelpath, true);
                            listBoxFiles.Items.Add(DateTime.Now.ToString()+" success: " + excelpath);
                            //添加对应关系
                            String savePath = excelpath.Substring(excelpath.IndexOf("GSGGFWB"));
                            listBoxFiles.Items.Add("save path: " + savePath);
                            savePdfToExcelInfo(articleList, Path.GetFileName(processedJob.SourcePath), savePath);
                        }
                    }
                }

                

                //结束生成
                listBoxFiles.Items.Add("convert end .....");
                //buttonStart.Enabled = true;
                //buttonStop.Enabled = false;
            }
        }

        public void savePdfToExcelInfo(List<AnnouncementEntity> articleList,String pdfName,String excelPath)
        {
            foreach (AnnouncementEntity ae in articleList)
            {
                if (ae.pdfPath.Contains(pdfName))
                {
                    if (!dao.getAnnouncementCount(ae.id))
                    {
                        if (excelPath.Equals(""))
                        {
                            dao.savePdfToExcelInfo(ae, excelPath,false);
                        }
                        else
                        {
                            dao.savePdfToExcelInfo(ae, excelPath, true);
                        }
                        break;
                        
                    }
                    break;
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
    }
}
