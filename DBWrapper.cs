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
                        CurrYear = "ConnMiddleWare";
                    }
                }
                else if (HttpContext.Current.Session == null)
                {
                    CurrYear = "ConnMiddleWare  ";
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
            DataSet dts = new DataSet();

            // defensive: ensure ProcName provided
            if (string.IsNullOrWhiteSpace(ProcName))
                throw new ArgumentException("ProcName is required.", nameof(ProcName));

            using (SqlConnection conn = GetNewSqlConnection())
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcName;
                cmd.CommandTimeout = 0;

                if (para != null && para.Length > 0)
                    cmd.Parameters.AddRange(para);

                // Debug: log parameters (optional, remove in production)
                try
                {
                    for (int i = 0; i < cmd.Parameters.Count; i++)
                    {
                        var p = cmd.Parameters[i];
                        System.Diagnostics.Debug.WriteLine($"Param[{i}] {p.ParameterName} = {(p.Value ?? "NULL")} (DbType={p.DbType})");
                    }
                }
                catch { /* ignore logging errors */ }

                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    // Ensure connection is available; SqlDataAdapter.Fill will open if closed,
                    // but opening explicitly gives clearer error messages and easier debugging.
                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    da.Fill(dts);

                    // Debug: log how many tables/rows we got (optional)
                    System.Diagnostics.Debug.WriteLine($"ReturnDS: tables = {dts.Tables.Count}");
                    for (int t = 0; t < dts.Tables.Count; t++)
                    {
                        System.Diagnostics.Debug.WriteLine($"ReturnDS: table[{t}] rows = {dts.Tables[t].Rows.Count}");
                    }
                }
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