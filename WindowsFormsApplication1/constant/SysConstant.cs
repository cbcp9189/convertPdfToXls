using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1.constant
{
    public class SysConstant
    {
        public static byte REPORTS_TYPE = 2;

        public static byte ANNOUNCEMENT_TYPE = 13;

        public static byte NEW_OTC_ANNOUNCEMENT_TYPE = 17;

        public static byte HK_ANNOUNCEMENT_TYPE = 19;

        public static byte WECHAT_ARTICLE_TYPE = 127;

        public static String ANNOUNCEMENT_EXCEL = "excel";

        public static String subscriptions = "subscriptions";

        public static double RANGE = 0.05;

        public static int ERROR = -1;

        public static int SUCCESS = 1;

        public static int NO_DEAL = 14;

        public static int MAX_PROCESS = 1000;


        //处理标识
        public static int BIG_FILE_FLAG = -21;

        public static int ERROR_STR_FLAG = -13;  //解析的日期中包含  @;

        public static String DATE_ERROR_STR = ";@";

        public static String PROGRAMNAME = "excel";  //程序名称

        public static int DEL = 1;

    }
}
