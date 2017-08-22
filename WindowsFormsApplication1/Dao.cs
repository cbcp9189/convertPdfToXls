using FluentFTP;
using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WindowsFormsApplication1.constant;
using WindowsFormsApplication1.entity;

namespace WindowsFormsApplication1
{
    public class Dao
    {
        public MySqlConnection getmysqlcon()
        {
            string M_str_sqlcon = "server=106.75.3.227;user id=root;password=hoboom;database=scrapy"; //根据自己的设置
            //string M_str_sqlcon = "server=106.75.116.2;user id=root;password=hoboom;database=hooboom"; //根据自己的设置
            return new MySqlConnection(M_str_sqlcon);
        }
        public void updateArticle(String xls_path,long id)
        {
            DataBaseConnect dc = new DataBaseConnect();
            //String sql = "UPDATE sys_config SET set_by = 2 where variable = 'diagnostics.include_raw'";
            StringBuilder sql = new StringBuilder("UPDATE article SET xls_path = '");
            sql.Append(xls_path);
            sql.Append("' where id =");
            sql.Append(id);
            dc.getmysqlcom(sql.ToString());
            Console.WriteLine("update end.....");
        }

        public List<AnnouncementEntity> getArticleList(long index, int count)
        {
            DataBaseConnect dc = new DataBaseConnect();
            StringBuilder sql = new StringBuilder("SELECT id,pdf_path from article where pdf_path != '' AND del_flag = 0 ");
            sql.Append(" limit ");
            sql.Append(index);
            sql.Append(",");
            sql.Append(count);
            MySqlDataReader reader = dc.getmysqlread(sql.ToString());
            List<AnnouncementEntity> list = new List<AnnouncementEntity>();
            while (reader.Read())
            {
                AnnouncementEntity ae = new AnnouncementEntity();
                ae.id = (long)reader["id"];
                ae.pdfPath = (string)reader["pdf_path"];
                list.Add(ae);

            }
            Console.WriteLine("get list end.....");
            return list;
        }

        public List<AnnouncementEntity> getAnnouncementList(long index, int count)
        {
           
            StringBuilder sql = new StringBuilder("SELECT id,announcement_id,file_path from stock_announcement where file_path != '' AND del_flag = 0 and DATE_FORMAT(pub_date,'%Y')='2017'");
            sql.Append(" limit ");
            sql.Append(index);
            sql.Append(",");
            sql.Append(count);
            
            MySqlConnection con = getmysqlcon();
            con.Open();
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);
            MySqlDataReader reader = mysqlcom.ExecuteReader();
            List<AnnouncementEntity> list = new List<AnnouncementEntity>();
            while (reader.Read())
            {
                AnnouncementEntity ae = new AnnouncementEntity();
                ae.id = (long)reader["id"];
                ae.announcement_id = (long)reader["announcement_id"];
                ae.pdfPath = (string)reader["file_path"];
                list.Add(ae);

            }
            Console.WriteLine("get list end.....");
            con.Close();
            return list;
        }


        public void savePdfToExcelInfo(AnnouncementEntity aey,String excelPath,Boolean isSucces)
        {
            
            StringBuilder sql = new StringBuilder("INSERT INTO pdf_to_excel(docid,doctype,origin_id,state,excel_path) VALUES(");
            sql.Append(aey.id);
            sql.Append(", ");
            sql.Append(SysConstant.ANNOUNCEMENT_TYPE);
            sql.Append(", ");
            sql.Append(aey.announcement_id);
            sql.Append(", ");
            if (isSucces) {
                sql.Append(0);
            }
            else
            {
                sql.Append(1);
            }
            
            sql.Append(",'");
            sql.Append(excelPath);
            sql.Append("')");
            Console.WriteLine(sql.ToString());
            MySqlConnection con = getmysqlcon();
            con.Open();
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);
            mysqlcom.ExecuteNonQuery();
            con.Close();
            Console.WriteLine("insert end.....");
        }

        public Boolean getAnnouncementCount(long id)
        {
           
            StringBuilder sql = new StringBuilder("SELECT count(*) count from pdf_to_excel where docid = ");
            sql.Append(id);
            MySqlConnection con = getmysqlcon();
            con.Open();
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);
            MySqlDataReader reader = mysqlcom.ExecuteReader();
            if (reader.Read())
            {
                Object obj = (Object)reader["count"];
                con.Close();
                if (obj != null && int.Parse(obj.ToString()) != 0)
                {
                    return true;
                }
                return false;

            }
            Console.WriteLine("get getAnnouncementCount end.....");
            return false;
        }


        public long getMinId()
        {
            MySqlConnection con = getmysqlcon();
            con.Open();
            StringBuilder sql = new StringBuilder("SELECT min(id) min from stock_announcement where del_flag = 0");
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);

            MySqlDataReader reader = mysqlcom.ExecuteReader();
            if (reader.Read())
            {
                long obj = (long)reader["min"];
                con.Close();
               return obj;

            }
            con.Close();
            return 0;
        }

        public long getMaxId()
        {
            MySqlConnection con = getmysqlcon();
            con.Open();
            StringBuilder sql = new StringBuilder("SELECT max(id) max from stock_announcement where del_flag = 0");
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);
            MySqlDataReader reader = mysqlcom.ExecuteReader();
            if (reader.Read())
            {
                long obj = (long)reader["max"];
                con.Close();
                return obj;

            }
            con.Close();
            return 0;
        }

        public List<AnnouncementEntity> getAnnouncementList(long start, long end)
        {
            MySqlConnection con = getmysqlcon();
            con.Open();
            StringBuilder sql = new StringBuilder("SELECT id,announcement_id,file_path from stock_announcement where file_path != '' AND del_flag = 0 and DATE_FORMAT(pub_date,'%Y')='2017'");
            sql.Append(" and id >= ");
            sql.Append(start);
            sql.Append(" and id <= ");
            sql.Append(end);
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);
            MySqlDataReader reader = mysqlcom.ExecuteReader();
            List<AnnouncementEntity> list = new List<AnnouncementEntity>();
            while (reader.Read())
            {
                AnnouncementEntity ae = new AnnouncementEntity();
                ae.id = (long)reader["id"];
                ae.announcement_id = (long)reader["announcement_id"];
                ae.pdfPath = (string)reader["file_path"];
                list.Add(ae);

            }
            con.Close();
            Console.WriteLine("get list end.....");
            return list;
        }
    }
}
