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
            StringBuilder sql = new StringBuilder("SELECT id,doc_id,pdf_path,doc_type from pdf_stream where pdf_path != '' AND doc_type = 13 AND excel_flag = 0 and pdf_path like '%/2016/%' ");
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
            Console.WriteLine("get list end.....");
            con.Close();
            return list;
        }


        public void savePdfToExcelInfo(AnnouncementEntity aey, String excelPath, Boolean isSucces, TableEntity tb)
        {

            StringBuilder sql = new StringBuilder("INSERT INTO pdf_to_excel(docid,doctype,excel_path,page_number,total_page,left_x,top_y,right_x,bottom_y,create_time) VALUES(");
            sql.Append(aey.doc_id);
            sql.Append(", ");
            sql.Append(aey.doc_type);
            sql.Append(",'");
            sql.Append(excelPath);
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
            sql.Append(")");
            Console.WriteLine(sql.ToString());
            MySqlConnection con = getmysqlcon();
            con.Open();
            MySqlCommand mysqlcom = new MySqlCommand(sql.ToString(), con);
            mysqlcom.ExecuteNonQuery();
            //同时更新pdf_stream表中的excel_flag
            //UPDATE pdf_stream SET excel_flag = 1 where doc_id = 137922 AND doc_type = 2
            StringBuilder updateSql = new StringBuilder("UPDATE pdf_stream SET excel_flag = 1 where doc_id = ");
            updateSql.Append(aey.doc_id);
            updateSql.Append(" AND doc_type = ");
            updateSql.Append(aey.doc_type);
            MySqlCommand mysqlcom1 = new MySqlCommand(updateSql.ToString(), con);
            mysqlcom1.ExecuteNonQuery();
            con.Close();
            Console.WriteLine("insert and update end.....");
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

        public long getMinId()
        {
            MySqlConnection con = getmysqlcon();
            con.Open();
            StringBuilder sql = new StringBuilder("SELECT min(id) min from pdf_stream");
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
            StringBuilder sql = new StringBuilder("SELECT max(id) max from pdf_stream");
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
            StringBuilder sql = new StringBuilder("SELECT id,doc_id,pdf_path,doc_type from pdf_stream where pdf_path != '' AND doc_type = 13  AND excel_flag = 0 ");
            sql.Append(" and id >= ");
            sql.Append(start);
            sql.Append(" and id <");
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
                ae.doc_type = (byte)reader["doc_type"];
                list.Add(ae);

            }
            con.Close();
            Console.WriteLine("get list end.....");
            return list;
        }
    }
}
