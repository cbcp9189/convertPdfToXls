using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1.entity
{
    //txt实体类
    public class TableEntity
    {
        public int totalPage { get; set; } //总页数

        public int pageNumber { get; set; } //当前页
        public float left { get; set; }
        public float right { get; set; }
        public float top { get; set; }
        public float bottom { get; set; }
       
        public int content_type;  //内容类型
        public int content_id;   //内容id
        public String content = "";  //内容
        public long pdf_id;    //pdfid
        public int pages = 0;     //跨的页数
        public String excelPath;  //excel路径
        public List<String> sheetIds;  //需要合并的sheetid
        public int flag = 0;  //状态
    }
}
