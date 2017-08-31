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
using WindowsFormsApplication1.util;

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

        public MySqlConnection getmysqlconlocal()
        {
            string M_str_sqlcon = "server=127.0.0.1;user id=root;password=root;database=zs"; //根据自己的设置
            //string M_str_sqlcon = "server=106.75.116.2;user id=root;password=hoboom;database=hooboom"; //根据自己的设置
            return new MySqlConnection(M_str_sqlcon);
        }
       

        public List<AnnouncementEntity> getArticleList(long index, int count)
        {
            DataBaseConnect dc = new DataBaseConnect();
            StringBuilder sql = new StringBuilder("SELECT id,pdf_path from article where pdf_path != '' AND del_flag = 0  ");
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
            StringBuilder sql = new StringBuilder("SELECT id,doc_id,pdf_path,doc_type from pdf_stream where pdf_path != '' AND doc_type = 13 AND excel_flag = 0 and update_flag = 3 ");
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
                ae.doc_id = (long)reader["doc_id"];
                ae.pdfPath = (string)reader["pdf_path"];
                ae.doc_type = 13;
                list.Add(ae);
            }
            Console.WriteLine("get list end....." + sql);
            LogHelper.WriteLog(typeof(Dao),sql.ToString());
            con.Close();
            return list;
        }


        public void savePdfToExcelInfo(AnnouncementEntity aey, KeyValEntity kve,TableEntity tb)
        {

            StringBuilder sql = new StringBuilder("INSERT INTO pdf_to_excel(docid,doctype,pdf_stream_id,excel_path,page_number,total_page,left_x,top_y,right_x,bottom_y,create_time,version,content) VALUES(");
            sql.Append(aey.doc_id);
            sql.Append(", ");
            sql.Append(aey.doc_type);
            sql.Append(",");
            sql.Append(aey.id);
            sql.Append(",'");
            sql.Append(kve.desc);
            sql.Append("',");
            sql.Append(tb.pageNumber);
            sql.Append(",");
            sql.Append(tb.totalPage);
            sql.Append(",");
            sql.Append(tb.left);
            sql.Append(",");
            sql.Append(tb.top);
            sql.Append(",");
            sql.Append(tb.right);
            sql.Append(",");
            sql.Append(tb.bottom);
            sql.Append(",");
            sql.Append(DateTimeUtil.GetTimeStamp());
            sql.Append(",");
            sql.Append(DateTimeUtil.GetTimeStampWithMs());
             sql.Append(",");
             sql.Append("@content");
            sql.Append(")");
            //Console.WriteLine(sql.ToString());
            MySqlConnection con = getmysqlcon();
            con.Open();
            MySqlCommand mysqlcom = con.CreateCommand();
            mysqlcom.Parameters.AddWithValue("@content", kve.value);
            mysqlcom.CommandText = sql.ToString();
            mysqlcom.ExecuteNonQuery();
            con.Close();
            Console.WriteLine("insert and update end.....");
        }

        public void updatePdfStreamInfo(AnnouncementEntity aey,int status)
        {
            //更新pdf_stream表中的excel_flag和version字段
            StringBuilder updateSql = new StringBuilder("UPDATE pdf_stream SET excel_flag = ");
            updateSql.Append(status);
            updateSql.Append(" ,version =unix_timestamp(now())");
            updateSql.Append(" where id = ");
            updateSql.Append(aey.id);
            //updateSql.Append(" AND doc_type = ");
            //updateSql.Append(aey.doc_type);
            MySqlConnection con = getmysqlcon();
            con.Open();
            MySqlCommand mysqlcom1 = new MySqlCommand(updateSql.ToString(), con);
            mysqlcom1.ExecuteNonQuery();
            con.Close();
            LogHelper.WriteLog(typeof(Dao),updateSql.ToString());
            Console.WriteLine("update end.....");
        }

        public Boolean getAnnouncementCount(long id,int doc_type)
        {
           
            StringBuilder sql = new StringBuilder("SELECT count(*) count from pdf_to_excel where docid = ");
            sql.Append(id);
            sql.Append(" and doctype =");
            sql.Append(doc_type);
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

        public long getPdfStreamCount(int doc_type)
        {

            StringBuilder sql = new StringBuilder("SELECT count(*) count from pdf_stream where ");
            sql.Append("doc_type = ");
            sql.Append(doc_type);
            sql.Append(" and update_flag = 3 ");
            MySqlConnection con = getmysqlcon();
            con.Open();
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);
            MySqlDataReader reader = mysqlcom.ExecuteReader();
            if (reader.Read())
            {
                long count = (long)reader["count"];
                con.Close();
                return count;
            }
            return 0;
        }


        public long getMinId()
        {
            MySqlConnection con = getmysqlcon();
            con.Open();
            StringBuilder sql = new StringBuilder("SELECT min(id) min from pdf_stream where update_flag = 4 ");
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
            StringBuilder sql = new StringBuilder("SELECT max(id) max from pdf_stream where update_flag = 4 ");
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
            StringBuilder sql = new StringBuilder("SELECT id,doc_id,pdf_path,doc_type from pdf_stream where pdf_path != '' AND doc_type = 13  AND excel_flag = 0  and pdf_path like '%/2016/%' ");
            sql.Append(" and id >= ");
            sql.Append(start);
            sql.Append(" and id < ");
            sql.Append(end);
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);
            MySqlDataReader reader = mysqlcom.ExecuteReader();
            List<AnnouncementEntity> list = new List<AnnouncementEntity>();
            while (reader.Read())
            {
                AnnouncementEntity ae = new AnnouncementEntity();
                ae.id = (long)reader["id"];
                ae.doc_id = (long)reader["doc_id"];
                ae.pdfPath = (string)reader["pdf_path"];
                ae.doc_type = 13;
                list.Add(ae);
            }
            con.Close();
            Console.WriteLine("get list end....."+sql.ToString());
            return list;
        }

        public void testspecilStr(String txt) {
            MySqlConnection con = getmysqlconlocal();
            con.Open();
            String sql = "INSERT t_title(name,catalog_id,content) VALUES(@name,1,@content)";
            MySqlCommand mysqlcom = con.CreateCommand();
            
            mysqlcom.Parameters.AddWithValue("@name", "hello world");
            mysqlcom.Parameters.AddWithValue("@content", txt);
            mysqlcom.CommandText = sql;
            mysqlcom.ExecuteNonQuery();
            con.Close();
        }
    }
}
