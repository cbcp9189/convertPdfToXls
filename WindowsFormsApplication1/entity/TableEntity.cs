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



    }
}
