using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace ChoiceDealing
{
    public class DBWrapper
    {
        public static object objReturnObject;

        public static object ReturnObject
        {
            get { return objReturnObject; }
        }

        public DBWrapper()
        {

        }

        public static SqlConnection GetNewSqlConnection()
        {
            try
            {
                string CurrYear = "";

                if (HttpContext.Current.Session != null)
                {

                    if (HttpContext.Current.Session["Year"] != null)
                    {
                        CurrYear = System.Convert.ToString(HttpContext.Current.Session["Year"]);
                    }
                    else
                    {
                        CurrYear = "Conn2";
                    }
                }
                else if (HttpContext.Current.Session == null)
                {
                    CurrYear = "Conn2";
                }

                return new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings[CurrYear].ToString());
            }
            catch
            {
                throw new ApplicationException("SQL connection string is not valid");
            }
        }

        public static SqlConnection GetNewSqlConnection(string strConn)
        {
            try
            {
                return new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings[strConn].ToString());
            }
            catch
            {
                throw new ApplicationException("SQL connection string is not valid");
            }
        }

        public static DataSet ReturnDS(SqlParameter[] para, string ProcName)
        {
            SqlConnection conn = GetNewSqlConnection();
            SqlCommand cmd = new SqlCommand();
            DataSet dts = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(cmd);

            try
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;
                cmd.CommandText = ProcName;
                cmd.Parameters.AddRange(para);
                da.Fill(dts);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (conn.State != System.Data.ConnectionState.Closed)
                    conn.Close();
                conn.Dispose();
            }
            return dts;
        }

        public static DataSet ReturnDS(string ProcName)
        {

            SqlConnection conn = GetNewSqlConnection();
            SqlCommand cmd = new SqlCommand();
            DataSet dts = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(cmd);

            try
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;
                cmd.CommandText = ProcName;
                da.Fill(dts);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (conn.State != System.Data.ConnectionState.Closed)
                    conn.Close();
                conn.Dispose();
            }
            return dts;
        }
    }
}