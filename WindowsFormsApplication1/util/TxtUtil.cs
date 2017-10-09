using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1.util
{
    class TxtUtil
    {
        static SpreadsheetDocument spreadSheetDocument = null;

        public static void createExcel2(string excelFileName)
        {
            spreadSheetDocument = SpreadsheetDocument.Open(excelFileName, true);

            WorkbookPart workBookPart = spreadSheetDocument.WorkbookPart;
            List<string> lstComments = new List<string>();
            foreach (WorksheetPart sheet in workBookPart.WorksheetParts)
            {

                foreach (WorksheetCommentsPart commentsPart in sheet.GetPartsOfType<WorksheetCommentsPart>())
                {
                    foreach (Comment comment in commentsPart.Comments.CommentList)
                    {
                        //lstComments.Add(comment.InnerText);
                        Console.WriteLine(comment.InnerText);
                    }
                }
            }

        }
    }
}
