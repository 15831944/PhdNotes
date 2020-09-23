using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;              //引用数据库就能用到这句话
using System.Data.SqlClient;    //连接数据库

namespace Phd
{
    class PhdSqlServer
    {
        private SqlConnection m_connect;    //
        private SqlCommand m_command;
        private SqlDataReader m_reader;
        private string m_strLastError;

        public PhdSqlServer()
        { }
        ~PhdSqlServer()
        { }

        //连接数据库
        public bool Connect(string strServerName,string strDbName,string strUserName,string strPassword)
        {
            try
            {
                string strConnect = string.Format("server={0};uid={1};pwd={2};database={3}",
                   strServerName, strUserName, strPassword, strDbName);
                m_connect = new SqlConnection(strConnect);
                m_command = m_connect.CreateCommand();
                m_connect.Open();
            }
            catch(Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return false;
            }
            return true;
        }

        //关闭数据库
        public void Close()
        {
            if (m_reader != null)
                m_reader.Dispose();
            if (m_command != null)
                m_command.Dispose();
            if (m_connect != null)
            {
                m_connect.Dispose();//释放资源
                m_connect.Close();   //关闭连接
            }
        }

        //增删改插入数据
        public int ExecSql(string strSql)
        {
            int nRow = 0;//对连接执行SQL语句并返回受影响的行数
            try
            {
                m_command.CommandText = strSql;
                nRow = m_command.ExecuteNonQuery();
            }
            catch(Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return 0;
            }
            return nRow;
        }

        //查询数据
        public bool Select(string strSql)
        {
            try
            {
                if (m_reader != null)
                {
                    m_reader.Dispose();
                    m_reader.Close();
                }
                m_command.CommandText = strSql;
                m_reader = m_command.ExecuteReader();
                if (!m_reader.HasRows)
                    return false;
            }
            catch (Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return false;
            }
            return true;
        }

        //遍历数据
        public bool MoveNext()
        {
            bool bRt = false;
            try
            {
                bRt = m_reader.Read();
            }
            catch(Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return false;
            }
            return bRt;
        }

        //通过索引得到数据
        public bool GetFieldByIndex(int nIndex, ref string strValue)
        {
            try
            {
                strValue = m_reader.GetString(nIndex);
            }
            catch(Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return false;
            }
            return true;
        }
        public bool GetFieldByIndex(int nIndex, ref int nValue)
        {
            try
            {
                nValue = m_reader.GetInt32(nIndex);
            }
            catch(Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return false;
            }
            return true;
        }
        public bool GetFieldByIndex(int nIndex, ref double dValue)
        {
            try
            {
                dValue = m_reader.GetDouble(nIndex);
            }
            catch (Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return false;
            }
            return true;
        }
        public bool GetFieldByIndex(int nIndex, ref DateTime value)
        {
            try
            {
                value = m_reader.GetDateTime(nIndex);
            }
            catch (Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return false;
            }
            return true;
        }
        public bool GetFieldByIndex(int nIndex, ref bool bValue)
        {
            try
            {
                bValue = m_reader.GetBoolean(nIndex);
            }
            catch (Exception e)
            {
                m_strLastError = Convert.ToString(e);
                return false;
            }
            return true;
        }

        //得到错误信息
        public string GetLastError()
        {
            return m_strLastError;
        }
        //得到列的数量
        public int GetFieldCount()
        {
            return m_reader.FieldCount;
        }
        //得到指定列的数据类型
        public string GetFieldTypeByIndex(int nIndex)
        {
            return m_reader.GetDataTypeName(nIndex);
        }

    }
}
