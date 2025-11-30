using Microsoft.Ajax.Utilities;
using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using FieldType = Microsoft.VisualBasic.FileIO.FieldType;

namespace ChoiceDealing
{
    public partial class MotiAMCBasketFile : System.Web.UI.Page
    {
        Label lblQuantity;
        Label lblNoOfBaskets;
        Label lblTabName;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        #region View
        protected void btnView_Click(object sender, EventArgs e)
        {
            ShowGrid();
        }
        #endregion
        #region Showgrid
        private void ShowGrid()
        {
            try
            {
                DataSet dts = new DataSet();

                SqlParameter[] para = new SqlParameter[1];
                para[0] = new SqlParameter("@Option", "MOTIVIEW");

                dts = DBWrapper.ReturnDS(para, "usp_InstiEXCEL");
                if (dts.Tables.Count > 0)
                {
                    Moti_AMCReport.DataSource = dts;
                    Moti_AMCReport.DataBind();
                    Moti_AMCReport.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
        protected void btnDownload_Click(object sender, EventArgs e)
        {
            try
            {
                DataSet dts = new DataSet();

                SqlParameter[] para = new SqlParameter[1];
                para[0] = new SqlParameter("@Option", "VIEW");

                dts = DBWrapper.ReturnDS(para, "usp_InstiEXCEL");
                if (dts.Tables.Count > 0 && dts.Tables[0].Rows.Count > 0)
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (var package = new OfficeOpenXml.ExcelPackage()) // EPPlus Library
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                        worksheet.Cells.LoadFromDataTable(dts.Tables[0], true); // Load data

                        Response.Clear();
                        //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        Response.ContentType = "text/csv";
                        Response.AddHeader("content-disposition", "attachment; filename=MotiAMCFile.csv");

                        using (var stream = new MemoryStream())
                        {
                            package.SaveAs(stream);
                            stream.WriteTo(Response.OutputStream);
                        }

                        Response.Flush();
                        Response.End();
                    }
                }
                else
                {
                    // Show message if no data found
                    ScriptManager.RegisterStartupScript(this, GetType(), "alertMessage", "alert('No data found to export!');", true);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
                                        new SqlCommand("TRUNCATE TABLE tbl_Tabnames", connection).ExecuteNonQuery();
                                        new SqlCommand("TRUNCATE TABLE tbl_ExcelMain", connection).ExecuteNonQuery();
                                    }
                                }

                                foreach (DataRow row in formattedTable.Rows)
                                {
                                    Instireminagtabs instireminagtabs = new Instireminagtabs
                                    {
                                        Option = "REMAININGTABS",
                                        Tabname = sheetName,
                                        Isin = formattedTable.Columns.Count > 1 ? row[1]?.ToString() : null,
                                        SecurityName = formattedTable.Columns.Count > 2 ? row[2]?.ToString() : null,
                                        Pricedate = formattedTable.Columns.Count > 3 ? row[3]?.ToString() : null,
                                        ClosingMarketPriceNSE = formattedTable.Columns.Count > 4 ? row[4]?.ToString() : null,
                                        AdjustedClosingMarketPriceNSE = formattedTable.Columns.Count > 5 ? row[5]?.ToString() : null,
                                        PurchaseableUnits = formattedTable.Columns.Count > 6 ? row[6]?.ToString() : null,
                                        Adjustedvalue = formattedTable.Columns.Count > 7 ? row[7]?.ToString() : null,
                                        PercentageinCreationUnit = formattedTable.Columns.Count > 8 ? row[8]?.ToString() : null,
                                        NiftyWeightage = formattedTable.Columns.Count > 9 ? row[9]?.ToString() : null
                                    };
                                    DataTable datatable2 = InstiTabs(instireminagtabs);
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

                                InstiMain main = new InstiMain
                                {
                                    TabName = sheetName,
                                    ClientCode = "",
                                    Transactions = totalUnits.ToString(),
                                    Qty = totalUnits.ToString(),
                                    Linecount = lineCount.ToString(),
                                };
                                DataTable table1 = INSTIMain(main);
                            }
                        }
                    }

                    // Clean up the temporary file
                    File.Delete(filePath);
                    lblMessage.Text = "File processed successfully!";
                }
                else
                {
                    lblMessage.Text = "Please select a file to upload.";
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = "Error: " + ex.Message;
            }
        }
        public DataTable InstiTabs(Instireminagtabs Inputmodel)
        {
            SqlCommand command = new SqlCommand("usp_InstiEXCEL");
            command.Parameters.Add("@Option", SqlDbType.VarChar).Value = "REMAININGTABS";
            command.Parameters.Add("@TabName", SqlDbType.VarChar).Value = Inputmodel.Tabname;
            command.Parameters.Add("@ISIN", SqlDbType.VarChar).Value = Inputmodel.Isin;
            command.Parameters.Add("@SecurityName", SqlDbType.VarChar).Value = Inputmodel.SecurityName;
            command.Parameters.Add("@Pricedate", SqlDbType.VarChar).Value = Inputmodel.Pricedate;
            command.Parameters.Add("@ClosingMarketPriceNSE", SqlDbType.VarChar).Value = Inputmodel.ClosingMarketPriceNSE;
            command.Parameters.Add("@AdjustedClosingMarketPriceNSE", SqlDbType.VarChar).Value = Inputmodel.AdjustedClosingMarketPriceNSE;
            command.Parameters.Add("@PurchaseableUnits", SqlDbType.VarChar).Value = Inputmodel.PurchaseableUnits;
            command.Parameters.Add("@Adjustedvalue", SqlDbType.VarChar).Value = Inputmodel.Adjustedvalue;
            command.Parameters.Add("@PercentageinCreationUnit", SqlDbType.VarChar).Value = Inputmodel.PercentageinCreationUnit;
            command.Parameters.Add("@NiftyWeightage", SqlDbType.VarChar).Value = Inputmodel.NiftyWeightage;

            command.CommandType = CommandType.StoredProcedure;

            return ExecuteStoredProcedure(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString, command);
        }
        public DataTable INSTIMain(InstiMain Inputmodel)
        {
            SqlCommand command = new SqlCommand("usp_InstiEXCEL");
            command.Parameters.Add("@Option", SqlDbType.VarChar).Value = "INSTIMAIN";
            command.Parameters.Add("@TabName", SqlDbType.VarChar).Value = Inputmodel.TabName;
            command.Parameters.Add("@SCHEMENAME", SqlDbType.VarChar).Value = Inputmodel.SCHEMENAME;
            command.Parameters.Add("@ClientCode", SqlDbType.VarChar).Value = Inputmodel.ClientCode;
            command.Parameters.Add("@Transactions", SqlDbType.VarChar).Value = Inputmodel.Transactions;
            command.Parameters.Add("@OrderNo", SqlDbType.VarChar).Value = Inputmodel.OrderNo;
            command.Parameters.Add("@#ofBaskets", SqlDbType.VarChar).Value = Inputmodel.noofBaskets;
            command.Parameters.Add("@TF", SqlDbType.VarChar).Value = Inputmodel.TF;
            command.Parameters.Add("@Linecount", SqlDbType.VarChar).Value = Inputmodel.Linecount;
            command.Parameters.Add("@Qty", SqlDbType.VarChar).Value = Inputmodel.Qty;

            command.CommandType = CommandType.StoredProcedure;

            return ExecuteStoredProcedure(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString, command);
        }
        public class Instireminagtabs
        {
            public string Option { get; set; }
            public string Tabname { get; set; }
            public string Isin { get; set; }
            public string SecurityName { get; set; }
            public string Pricedate { get; set; }
            public string ClosingMarketPriceNSE { get; set; }
            public string AdjustedClosingMarketPriceNSE { get; set; }
            public string PurchaseableUnits { get; set; }
            public string Adjustedvalue { get; set; }
            public string PercentageinCreationUnit { get; set; }
            public string NiftyWeightage { get; set; }
        }
        public class InstiMain
        {
            public string Option { get; set; }
            public string TabName { get; set; }
            public string SCHEMENAME { get; set; }
            public string ClientCode { get; set; }
            public string Transactions { get; set; }
            public string OrderNo { get; set; }
            public string noofBaskets { get; set; }
            public string TF { get; set; }
            public string Linecount { get; set; }
            public string Qty { get; set; }

        }
        protected void btnBhavUpload_Click(object sender, EventArgs e)
        {
            try
            {
                if (BhavfileUpload.HasFile)
                {
                    string fileName = Path.GetFileName(BhavfileUpload.PostedFile.FileName);
                    string filePath = Server.MapPath("~/FileUploads/" + fileName);

                    // Save file temporarily on the server
                    BhavfileUpload.SaveAs(filePath);

                    DataTable formattedTable = new DataTable();

                    using (var reader = new StreamReader(filePath, Encoding.UTF8))
                    using (var parser = new TextFieldParser(reader))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");
                        parser.HasFieldsEnclosedInQuotes = true;

                        bool isHeader = true;

                        while (!parser.EndOfData)
                        {
                            string[] fields = parser.ReadFields();

                            if (isHeader)
                            {
                                foreach (string column in fields)
                                {
                                    string columnName = column.Trim();
                                    formattedTable.Columns.Add(string.IsNullOrWhiteSpace(columnName) ? "Column" + formattedTable.Columns.Count : columnName);
                                }
                                isHeader = false;
                            }
                            else
                            {
                                formattedTable.Rows.Add(fields);
                            }
                        }
                    }
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString))
                    {
                        connection.Open();

                        // Assuming _Commonservice.ExecuteTSQL just runs a T-SQL command before insert
                        // Example: delete from table or truncate
                        string bhavSqlCommand = "TRUNCATE TABLE tblBhavCopy"; // If you store SQL in config
                        SqlCommand cmd = new SqlCommand(bhavSqlCommand, connection);
                        cmd.ExecuteNonQuery();

                        foreach (DataRow row in formattedTable.Rows)
                        {
                            BhavCopy instiMain = new BhavCopy
                            {
                                TradDt = row[0]?.ToString()?.Trim(),
                                BizDt = row[1]?.ToString()?.Trim(),
                                Sgmt = row[2]?.ToString()?.Trim(),
                                Src = row[3]?.ToString()?.Trim(),
                                FinInstrmTp = row[4]?.ToString()?.Trim(),
                                FinInstrmId = row[5]?.ToString()?.Trim(),
                                ISIN = row[6]?.ToString()?.Trim(),
                                TckrSymb = row[7]?.ToString()?.Trim(),
                                SctySrs = row[8]?.ToString()?.Trim(),
                                FinInstrmNm = row[9]?.ToString()?.Trim(),
                                OpnPric = row[10]?.ToString()?.Trim(),
                                HghPric = row[11]?.ToString()?.Trim(),
                                LwPric = row[12]?.ToString()?.Trim(),
                                ClsPric = row[13]?.ToString()?.Trim(),
                                LastPric = row[14]?.ToString()?.Trim(),
                                PrvsClsgPric = row[15]?.ToString()?.Trim(),
                                SttlmPric = row[16]?.ToString()?.Trim(),
                                TtlTradgVol = row[17]?.ToString()?.Trim(),
                                TtlTrfVal = row[18]?.ToString()?.Trim(),
                                TtlNbOfTxsExctd = row[19]?.ToString()?.Trim(),
                                SsnId = row[20]?.ToString()?.Trim(),
                                NewBrdLotQty = row[21]?.ToString()?.Trim(),
                            };

                            // Call your business layer method directly
                            // You should create a static or instance method to insert BhavCopy data
                            DataTable result = InsertBhavCopy(instiMain);
                        }
                    }
                    lblBhavMessage.Text = "File processed successfully!";
                    // Clean up the temporary file
                    File.Delete(filePath);
                }
                else
                {
                    lblBhavMessage.Text = "Please select a file to upload.";
                }
            }
            catch (Exception ex)
            {
                lblBhavMessage.Text = "Error: " + ex.Message;
            }
        }
        public DataTable InsertBhavCopy(BhavCopy bhavCopy)
        {
            SqlCommand command = new SqlCommand("usp_InstiEXCEL");
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.Add("@Option", SqlDbType.VarChar).Value = "BHAVCOPY";
            command.Parameters.Add("@TradeDate", SqlDbType.VarChar).Value = bhavCopy.TradDt;
            command.Parameters.Add("@BizDate", SqlDbType.VarChar).Value = bhavCopy.BizDt;
            command.Parameters.Add("@Segment", SqlDbType.VarChar).Value = bhavCopy.Sgmt;
            command.Parameters.Add("@Src", SqlDbType.VarChar).Value = bhavCopy.Src;
            command.Parameters.Add("@FinInstrmTp", SqlDbType.VarChar).Value = bhavCopy.FinInstrmTp;
            command.Parameters.Add("@FinInstrmId", SqlDbType.VarChar).Value = bhavCopy.FinInstrmId;
            command.Parameters.Add("@ISIN", SqlDbType.VarChar).Value = bhavCopy.ISIN;
            command.Parameters.Add("@TckrSymb", SqlDbType.VarChar).Value = bhavCopy.TckrSymb;
            command.Parameters.Add("@SctySrs", SqlDbType.VarChar).Value = bhavCopy.SctySrs;
            //command.Parameters.Add("@XpryDt", SqlDbType.VarChar).Value = bhavCopy.XpryDt ;
            //command.Parameters.Add("@FininstrmActlXpryDt", SqlDbType.VarChar).Value = bhavCopy.FininstrmActlXpryDt;
            //command.Parameters.Add("@StrkPric", SqlDbType.VarChar).Value = bhavCopy.StrkPric;
            //command.Parameters.Add("@OptnTp", SqlDbType.VarChar).Value = bhavCopy.OptnTp;
            command.Parameters.Add("@FinInstrmNm", SqlDbType.VarChar).Value = bhavCopy.FinInstrmNm;
            command.Parameters.Add("@OpnPric", SqlDbType.VarChar).Value = bhavCopy.OpnPric;
            command.Parameters.Add("@HghPric", SqlDbType.VarChar).Value = bhavCopy.HghPric;
            command.Parameters.Add("@LwPric", SqlDbType.VarChar).Value = bhavCopy.LwPric;
            command.Parameters.Add("@ClsPric", SqlDbType.VarChar).Value = bhavCopy.ClsPric;
            command.Parameters.Add("@LastPric", SqlDbType.VarChar).Value = bhavCopy.LastPric;
            command.Parameters.Add("@PrvsClsgPric", SqlDbType.VarChar).Value = bhavCopy.PrvsClsgPric;
            command.Parameters.Add("@UndrlygPric", SqlDbType.VarChar).Value = bhavCopy.UndrlygPric;
            command.Parameters.Add("@SttlmPric", SqlDbType.VarChar).Value = bhavCopy.SttlmPric;
            command.Parameters.Add("@OpnIntrst", SqlDbType.VarChar).Value = bhavCopy.OpnIntrst;
            command.Parameters.Add("@ChngInOpnIntrst", SqlDbType.VarChar).Value = bhavCopy.ChngInOpnIntrst;
            command.Parameters.Add("@TtlTradgVol", SqlDbType.VarChar).Value = bhavCopy.TtlTradgVol;
            command.Parameters.Add("@TtlTrfVal", SqlDbType.VarChar).Value = bhavCopy.TtlTrfVal;
            command.Parameters.Add("@TtlNbOfTxsExctd", SqlDbType.VarChar).Value = bhavCopy.TtlNbOfTxsExctd;
            command.Parameters.Add("@SsnId", SqlDbType.VarChar).Value = bhavCopy.SsnId;
            command.Parameters.Add("@NewBrdLotQty", SqlDbType.VarChar).Value = bhavCopy.NewBrdLotQty;

            return ExecuteStoredProcedure(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString, command);
            //return DBWrapper.ReturnDS(command, "[LKP_Middleware_Config].[dbo].[usp_InstiEXCEL]");
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
        public class BhavCopy
        {
            public string TradDt { get; set; }
            public string BizDt { get; set; }
            public string Sgmt { get; set; }
            public string Src { get; set; }
            public string FinInstrmTp { get; set; }
            public string FinInstrmId { get; set; }
            public string ISIN { get; set; }
            public string TckrSymb { get; set; }
            public string SctySrs { get; set; }
            public string XpryDt { get; set; } // Nullable in case Expiry Date is optional
            public string FininstrmActlXpryDt { get; set; }
            public string StrkPric { get; set; }
            public string OptnTp { get; set; }
            public string FinInstrmNm { get; set; }
            public string OpnPric { get; set; }
            public string HghPric { get; set; }
            public string LwPric { get; set; }
            public string ClsPric { get; set; }
            public string LastPric { get; set; }
            public string PrvsClsgPric { get; set; }
            public string UndrlygPric { get; set; }
            public string SttlmPric { get; set; }
            public string OpnIntrst { get; set; }
            public string ChngInOpnIntrst { get; set; }
            public string TtlTradgVol { get; set; }
            public string TtlTrfVal { get; set; }
            public string TtlNbOfTxsExctd { get; set; }
            public string SsnId { get; set; }
            public string NewBrdLotQty { get; set; }
        }
        protected void Moti_AMCReport_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DropDownList ddlBuySell = (DropDownList)e.Row.FindControl("ddlBuySell");
                TextBox txtClientCode = (TextBox)e.Row.FindControl("txtClientcode");
                TextBox txtNoofBaskets = (TextBox)e.Row.FindControl("txtNoofBaskets");

                // Assuming SCHEMENAME is the second column (index 1), or you can use DataBinder
                string buySellValue = DataBinder.Eval(e.Row.DataItem, "SCHEMENAME")?.ToString();

                if (!string.IsNullOrEmpty(buySellValue) && (buySellValue == "B" || buySellValue == "S"))
                {
                    ddlBuySell.SelectedValue = buySellValue;
                }

                bool isRowEdited = (Convert.ToInt32(txtNoofBaskets.Text) > 1) ? true : false;

                // Access the Download button (ButtonField becomes LinkButton in Controls[0] of the cell)
                if (isRowEdited)
                {
                    // Adjust cell index if your ButtonField is not the first column (e.g. use correct column index)
                    LinkButton lnkDownload = e.Row.Cells[9].Controls[0] as LinkButton;
                    if (lnkDownload != null)
                    {
                        lnkDownload.BackColor = System.Drawing.Color.Yellow;
                    }
                }
            }
        }

        protected void Moti_AMCReport_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            try
            {
                if (e.CommandName == "View")
                {
                    // Ensure the row index is available
                    int rowIndex = Convert.ToInt32(e.CommandArgument);

                    // Access the correct row using the row index
                    GridViewRow row = Moti_AMCReport.Rows[rowIndex];

                    string _TabName = ((Label)row.FindControl("lblSCHEMECODE")).Text;  // Example field
                    string Link = $"MotiTabName.aspx?TabName={HttpUtility.UrlEncode(_TabName)}";
                    //Response.Redirect(Link, false);
                    //Context.ApplicationInstance.CompleteRequest(); // Prevents ThreadAbortException
                    string script = $"window.open('{Link}', '_blank');";
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "OpenNewTab", script, true);
                }
                else if (e.CommandName == "Download")
                {
                    int rowIndex = Convert.ToInt32(e.CommandArgument);
                    GridViewRow row = Moti_AMCReport.Rows[rowIndex];
                    string _TabName = ((Label)row.FindControl("lblSCHEMECODE")).Text;
                    string _ClientCode = ((TextBox)row.FindControl("txtClientcode")).Text;
                    DataSet dts = new DataSet();

                    SqlParameter[] para = new SqlParameter[2];
                    para[0] = new SqlParameter("@Option", "MOTIEXCELINDI");
                    para[1] = new SqlParameter("@TabName", _TabName);

                    dts = DBWrapper.ReturnDS(para, "usp_InstiEXCEL");
                    if (dts.Tables.Count > 0 && dts.Tables[0].Rows.Count > 0)
                    {
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using (var package = new OfficeOpenXml.ExcelPackage()) // EPPlus Library
                        {
                            //ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                            //worksheet.Cells.LoadFromDataTable(dts.Tables[0], true); // Load data

                            Response.Clear();
                            Response.ContentType = "text/csv";
                            Response.AddHeader("content-disposition", "attachment; filename=" + _TabName + "_" + _ClientCode + ".csv");
                            Response.ContentEncoding = System.Text.Encoding.UTF8;

                            StringBuilder csvBuilder = new StringBuilder();

                            // Assuming your datatable is dts.Tables[0]
                            DataTable dt = dts.Tables[0];

                            // Add column headers
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                csvBuilder.Append(dt.Columns[i].ColumnName);
                                if (i < dt.Columns.Count - 1)
                                    csvBuilder.Append(",");
                            }
                            csvBuilder.AppendLine();

                            // Add rows
                            foreach (DataRow rows in dt.Rows)
                            {
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    var field = rows[i].ToString();

                                    // Escape double quotes and commas if needed
                                    if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
                                    {
                                        field = "\"" + field.Replace("\"", "\"\"") + "\"";
                                    }

                                    csvBuilder.Append(field);
                                    if (i < dt.Columns.Count - 1)
                                        csvBuilder.Append(",");
                                }
                                csvBuilder.AppendLine();
                            }

                            Response.Write(csvBuilder.ToString());
                            Response.Flush();
                            Response.End();
                        }
                    }
                    else
                    {
                        // Show message if no data found
                        ScriptManager.RegisterStartupScript(this, GetType(), "alertMessage", "alert('No data found to export!');", true);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        protected void RowValueChanged(object sender, EventArgs e)
        {
            Control ctrl = (Control)sender;
            GridViewRow row = (GridViewRow)ctrl.NamingContainer;

            if (row != null)
            {
                // Get controls from row
                Label lblSchemeCode = (Label)row.FindControl("lblSCHEMECODE");
                DropDownList ddlBuySell = (DropDownList)row.FindControl("ddlBuySell");
                TextBox txtClientCode = (TextBox)row.FindControl("txtClientcode");
                TextBox txtNoofBaskets = (TextBox)row.FindControl("txtNoofBaskets");
                Label lblQuantity = (Label)row.FindControl("lblQUANTITY");
                Label lblInitialQUANTITY = (Label)row.FindControl("lblInitialQUANTITY");
                TextBox lbltxtTradeInstruction = (TextBox)row.FindControl("txtTradeInstruction");

                // Extract values
                string schemeCode = lblSchemeCode.Text;
                string buySell = ddlBuySell.SelectedValue;
                string clientCode = txtClientCode.Text;
                string TradeInstructions = lbltxtTradeInstruction.Text;
                int noOfBaskets = int.TryParse(txtNoofBaskets.Text, out int parsedBaskets) ? parsedBaskets : 0;
                int initialQuantity = GetInitialQuantityForRow(Convert.ToInt32(lblInitialQUANTITY.Text));
                int updatedQuantity = initialQuantity * noOfBaskets;
                lblQuantity.Text = updatedQuantity.ToString();

                // Save to database
                SqlParameter[] para = new SqlParameter[7];
                para[0] = new SqlParameter("@Option", "UPDATEINSTIMAIN");
                para[1] = new SqlParameter("@ClientCode", clientCode);
                para[2] = new SqlParameter("@SCHEMENAME", buySell);
                para[3] = new SqlParameter("@Qty", updatedQuantity);
                para[4] = new SqlParameter("@TabName", schemeCode);
                para[5] = new SqlParameter("@#ofBaskets", noOfBaskets);
                para[6] = new SqlParameter("@TradeInstructions", TradeInstructions);

                var res = DBWrapper.ReturnDS(para, "usp_InstiEXCEL");
                ShowGrid();
            }
        }
        private int GetInitialQuantityForRow(int rowIndex)
        {

            return rowIndex;
        }
    }
}
