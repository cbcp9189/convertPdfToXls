


using MySql.Data.MySqlClient;
using System.Collections.Generic;
using System.Data;
using WindowsFormsApplication1.entity;
namespace WindowsFormsApplication1.Test
{
    class TestDapper
    {
        public readonly string mysqlconnectionString =  
                 @"server=127.0.0.1;database=test;uid=renfb;pwd=123456;charset='gbk'"; 

        //获取MySql的连接数据库对象。MySqlConnection  
        public MySqlConnection OpenConnection()  
        {  
             MySqlConnection connection = new MySqlConnection(mysqlconnectionString);  
             connection.Open();  
             return connection;  
        }

       
    }
}
