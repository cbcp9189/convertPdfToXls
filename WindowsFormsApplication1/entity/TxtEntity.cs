using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1.entity
{
    //txt实体类
    public class TxtEntity
    {
        public long docid;    //pdfid

        public int doctype;
        public int pageNumber { get; set; } //当前页
        public float left { get; set; }
        public float right { get; set; }
        public float top { get; set; }
        public float bottom { get; set; }

        public int type;  //内容类型1段落  2表格  
        public int content_id;   //内容id
        public String content;  //内容
        public int is_tb_content = 0;     //是否在表格中  0 否   1是
    }
}
