using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

using System.IO;
using SolidFramework;
using SolidFramework.Converters.Plumbing;

using SolidFramework.Model;
using SolidFramework.Model.Layout;
using SolidFramework.Model.Plumbing;

using System.Threading.Tasks;
using WindowsFormsApplication1.entity;

namespace WindowsFormsApplication1
{
    class TestTxt
    {
        internal static void SolidModelLayoutTest1(string pdfFile, string outSvgFile, string outTxtFile)
        {
            PdfOptions options = new PdfOptions();
            options.TextRecoveryEngine = TextRecoveryEngine.SolidOCR;
            options.TextRecovery = TextRecovery.Automatic;
            options.TextRecoveryEngineNse = TextRecoveryEngineNse.Automatic;
            options.TextRecoveryNSE = TextRecoveryNSE.Automatic;
            options.ConvertMode = ConvertMode.Document;
            options.ReconstructionMode = ReconstructionMode.Flowing;
            options.ExposeTargetDocumentPagination = true;

            using (CoreModel model = CoreModel.Create(pdfFile, options))
            {
                LayoutDocument layoutDoc = model.GetLayout();
                TraceToTxt1(layoutDoc, outTxtFile);
                model.Dispose();
            }
        }

        static List<TableEntity> TraceToTxt1(LayoutDocument layoutDoc, string outputFile)
        {
            List<TableEntity> tbList = new List<TableEntity>();
            using (StreamWriter file = new StreamWriter(outputFile))
            {
                Console.WriteLine("TOTAL PAGES: {0}\n", layoutDoc.Count);
                int pageIndex = 0;
                foreach (LayoutObject page in layoutDoc)
                {
                    
                    RectangleF pageBounds = page.Bounds;
                    Console.WriteLine("PAGE #{0} (Left={1} Right={2} Top={3} Bottom={4}):\n",
                        ++pageIndex, pageBounds.Left, pageBounds.Right, pageBounds.Top, pageBounds.Bottom);
                    
                    Action<StreamWriter, LayoutObject> dumpEntities = null;
                    dumpEntities = (StreamWriter stream, LayoutObject obj) =>
                    {
                        switch (obj.GetObjectType())
                        {
                            case LayoutObjectType.Page:
                                {
                                    LayoutPage coll = obj as LayoutPage;
                                    foreach (LayoutObject obj1 in coll)
                                    {
                                        dumpEntities(stream, obj1);
                                    }
                                }
                                break;
                            case LayoutObjectType.Table:
                                {
                                    TableEntity tb = new TableEntity();
                                    tb.totalPage = layoutDoc.Count;
                                    tb.pageNumber = pageIndex;
                                    LayoutTable coll = obj as LayoutTable;
                                    Console.WriteLine("Table [ID={0}] (left:{1},right:{2},top:{3},buttom:{4})", coll.GetID(), coll.Bounds.Left, coll.Bounds.Right
                                        , coll.Bounds.Top, coll.Bounds.Bottom);
                                    tb.left = coll.Bounds.Left;
                                    tb.right = coll.Bounds.Right;
                                    tb.top = coll.Bounds.Top;
                                    tb.bottom = coll.Bounds.Bottom;
                                    tbList.Add(tb);
                                    //Console.WriteLine(String.Empty);
                                    foreach (LayoutObject obj1 in coll)
                                    {
                                        dumpEntities(stream, obj1);
                                    }
                                }
                                break;
                            case LayoutObjectType.Group:
                                {
                                    LayoutGroup coll = obj as LayoutGroup;
                                    //Console.WriteLine("Group [ID={0}]", coll.GetID());
                                    Console.WriteLine(String.Empty);
                                    foreach (LayoutObject obj1 in coll)
                                    {
                                        dumpEntities(stream, obj1);
                                    }
                                }
                                break;
                            case LayoutObjectType.TextBox:
                                {
                                    LayoutTextBox coll = obj as LayoutTextBox;
                                   // Console.WriteLine("TextBox [ID={0}]", coll.GetID());
                                    Console.WriteLine(String.Empty);
                                    foreach (LayoutObject obj1 in coll)
                                    {
                                        dumpEntities(stream, obj1);
                                    }
                                }
                                break;
                            case LayoutObjectType.Paragraph:
                                {
                                    LayoutParagraph par = obj as LayoutParagraph;

                                    string parText = par.AllText;
                                    if (0 != parText.Length)
                                    {
                                        RectangleF bounds = par.Bounds;
                                        //Console.WriteLine("Paragraph [ID={4}] (Left={0} Right={1} Top={2} Bottom={3}):\n" + parText,
                                         //   bounds.Left, bounds.Right, bounds.Top, bounds.Bottom, par.GetID());
                                        //Console.WriteLine(String.Empty);
                                    }
                                }
                                break;
                            default:
                                break;
                        }
                    };

                    dumpEntities(file, page);
                }
                return tbList;
            }
        }
    }
}
