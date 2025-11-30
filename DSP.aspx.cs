using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using static ChoiceDealing.MotiAMCBasketFile;

namespace ChoiceDealing
{
    public partial class DSP : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnView_Click(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            try
            {
                if (fileUpload.HasFile)
                {
                    string fileName = Path.GetFileName(fileUpload.PostedFile.FileName);
                    string filePath = Server.MapPath("~/FileUploads/" + fileName);

                    // Save file temporarily on the server
                    fileUpload.SaveAs(filePath);

                    FileInfo existingFile = new FileInfo(filePath);
                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            string sheetName = worksheet.Name;

                            if (sheetName.StartsWith("Sheet") && int.TryParse(sheetName.Substring(5), out _))
                                continue;

                            if (worksheet.Dimension == null)
                                continue;

                            DataTable formattedTable = new DataTable();
                            formattedTable.Columns.Add("TabName", typeof(string));
                            int maxColumns = Math.Min(worksheet.Dimension.End.Column - 1, 9);
                            int rowCount = worksheet.Dimension.End.Row;

                            // Add columns from header row (assumed to be row 1)
                            for (int col = 2; col <= maxColumns + 1; col++)
                            {
                                var columnName = worksheet.Cells[1, col].Text.Trim();
                                formattedTable.Columns.Add(string.IsNullOrEmpty(columnName) ? $"Column{col}" : columnName);
                            }

                            for (int row = 2; row <= rowCount; row++)
                            {
                                if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text))
                                    break;

                                DataRow newRow = formattedTable.NewRow();
                                newRow["TabName"] = sheetName;

                                for (int col = 2; col <= maxColumns + 1; col++)
                                {
                                    newRow[col - 1] = worksheet.Cells[row, col].Text;
                                }

                                formattedTable.Rows.Add(newRow);
                            }

                            if (formattedTable.Rows.Count > 0)
                            {
                                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString))
                                {
                                    connection.Open();
                                    if (worksheet.Index == 0)
                                    {
                                        new SqlCommand("TRUNCATE TABLE tbl_DSPBASKETVALUES", connection).ExecuteNonQuery();
                                        new SqlCommand("TRUNCATE TABLE tbl_DSPSCHEMES", connection).ExecuteNonQuery();
                                    }
                                }

                                foreach (DataRow row in formattedTable.Rows)
                                {
                                    DSPBasketValues dSPBasket = new DSPBasketValues
                                    {
                                        BASKET_ID = sheetName,
                                        BASKET_DATE = formattedTable.Columns.Count > 1 ? row[1]?.ToString() : null,
                                        SYMBOL = formattedTable.Columns.Count >2? row[2]?.ToString() : null,
                                        ISIN = formattedTable.Columns.Count>3? row[3]?.ToString() : null,
                                        SECURITY = formattedTable.Columns.Count>4? row[4]?.ToString() : null,
                                        QUANTITY = formattedTable.Columns.Count>5? row[5]?.ToString() : null,
                                        PRICE = formattedTable.Columns.Count>6? row[6]?.ToString() : null,
                                        VALUE = formattedTable.Columns.Count>7? row[7]?.ToString() : null
                                    };
                                   
                                    DataTable datatable2 = InstiTabs(dSPBasket);
                                }

                                // Insert summary row
                                int lineCount = formattedTable.Rows.Count;
                                decimal totalUnits = 0;
                                foreach (DataRow row in formattedTable.Rows)
                                {
                                    if (formattedTable.Columns.Count > 6 && decimal.TryParse(row[6]?.ToString(), out decimal unit))
                                    {
                                        totalUnits += unit;
                                    }
                                }

                                DSPSchemes main = new DSPSchemes
                                {
                                    SchemeCode = sheetName,
                                    ClientCode = "",
                                    //Transactions = totalUnits.ToString(),
                                    Qty = totalUnits.ToString(),
                                    Linecount = lineCount.ToString(),
                                };
                                DataTable table1 = INSTIMain(main);
                            }
                        
                        }
                    }

                    File.Delete(filePath);
                    lblMessage.Text = "File processed successfully!";
                }
                else
                {
                    lblMessage.Text = "Please select a file to upload.";
                }
            }
            catch(Exception ex)
            {
                lblMessage.Text = "Error: " + ex.Message;
            }
        }

        public DataTable InstiTabs(DSPBasketValues Inputmodel)
        {
            SqlCommand command = new SqlCommand("proc_DSP");
            command.Parameters.Add("@Option", SqlDbType.VarChar).Value = "INSERTDSP";
            command.Parameters.Add("@BASKET_ID", SqlDbType.VarChar).Value = Inputmodel.BASKET_ID;
            command.Parameters.Add("@BASKET_DATE", SqlDbType.VarChar).Value = Inputmodel.BASKET_DATE;
            command.Parameters.Add("@SYMBOL", SqlDbType.VarChar).Value = Inputmodel.SYMBOL;
            command.Parameters.Add("@ISIN", SqlDbType.VarChar).Value = Inputmodel.ISIN;
            command.Parameters.Add("@SECURITY", SqlDbType.VarChar).Value = Inputmodel.SECURITY;
            command.Parameters.Add("@QUANTITY", SqlDbType.VarChar).Value = Inputmodel.QUANTITY;
            command.Parameters.Add("@PRICE", SqlDbType.VarChar).Value = Inputmodel.PRICE;
            command.Parameters.Add("@VALUE", SqlDbType.VarChar).Value = Inputmodel.VALUE;

            command.CommandType = CommandType.StoredProcedure;

            return ExecuteStoredProcedure(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString, command);
        }
        public DataTable INSTIMain(DSPSchemes Inputmodel)
        {
            SqlCommand command = new SqlCommand("proc_DSP");
            command.Parameters.Add("@Option", SqlDbType.VarChar).Value = "MAINPAGE";
            command.Parameters.Add("@SCHEMECODE", SqlDbType.VarChar).Value = Inputmodel.SchemeCode;
            command.Parameters.Add("@BUYSELL", SqlDbType.VarChar).Value = Inputmodel.BuySell;
            command.Parameters.Add("@CLIENTCODE", SqlDbType.VarChar).Value = Inputmodel.ClientCode;
            command.Parameters.Add("@TRADEINSTRUCTIONS", SqlDbType.VarChar).Value = Inputmodel.TRADEINSTRUCTIONS;
            command.Parameters.Add("@NOOFBASKETS", SqlDbType.VarChar).Value = Inputmodel.noofBaskets;
            command.Parameters.Add("@INITIALQUANTITY", SqlDbType.VarChar).Value = Inputmodel.INITIALQUANTITY;
            command.Parameters.Add("@LINECOUNTS", SqlDbType.VarChar).Value = Inputmodel.Linecount;
            command.Parameters.Add("@QUANTITY", SqlDbType.VarChar).Value = Inputmodel.Qty;

            command.CommandType = CommandType.StoredProcedure;

            return ExecuteStoredProcedure(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString, command);
        }
        protected void btnDownload_Click(object sender, EventArgs e)
        {

        }

        public DataTable ExecuteStoredProcedure(string SqlConnection, SqlCommand command)
        {
            IDataReader reader = null;
            DataTable table = new DataTable();
            SqlConnection Connection = new SqlConnection(SqlConnection);

            command.Connection = Connection;
            command.CommandType = CommandType.StoredProcedure;
            Connection.Open();
            try
            {
                try
                {
                    using (reader = command.ExecuteReader())
                    {
                        table.Load(reader);
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    // Always call Close when done reading.
                    reader.Close();
                }
            }
            finally
            {
                Connection.Close();
            }
            return table;
        }

        public class DSPBasketValues
        {
            public string BASKET_ID { get; set; }
            public string BASKET_DATE {  get; set; }
            public string SYMBOL {  get; set; }
            public string ISIN {  get; set; }
            public string SECURITY { get; set; }

            public string QUANTITY { get; set; }

            public string PRICE { get; set; }

            public string VALUE { get; set; }
        }

        public class DSPSchemes
        {
            public string SchemeCode {  get; set; }
            public string BuySell {  get; set; }
            public string ClientCode {  get; set; }
            public int noofBaskets { get; set; }
            public string TRADEINSTRUCTIONS { get; set; } 
            public string INITIALQUANTITY { get; set; }
            public string Qty { get; set; }
            public string Linecount { get; set; }
        }
    }
}