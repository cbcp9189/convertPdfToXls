using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using WindowsFormsApplication1.entity;

namespace WindowsFormsApplication1.util
{
   public class FileUtil
    {
       public static List<PdfModel> mappingList = new List<PdfModel>();
        public static void createDir(String path) {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

       

    }
}
