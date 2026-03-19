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
using static ChoiceDealing.MotiAMCBasketFile;

namespace ChoiceDealing
{
    public partial class Groww : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnView_Click(object sender, EventArgs e)
        {
            ShowGrid();
        }

        private void ShowGrid()
        {
            try
            {
                DataSet dts = new DataSet();

                SqlParameter[] para = new SqlParameter[1];
                para[0] = new SqlParameter("@Option", "GROWVIEW");

                dts = DBWrapper.ReturnDS(para, "usp_GrowEXCEL");
                if (dts.Tables.Count > 0)
                {
                    Groww_AMCReport.DataSource = dts;
                    Groww_AMCReport.DataBind();
                    Groww_AMCReport.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        protected void btnDownload_Click(object sender, EventArgs e)
        {
            try
            {
                DataSet dts = new DataSet();

                SqlParameter[] para = new SqlParameter[1];
                para[0] = new SqlParameter("@Option", "VIEW");

                dts = DBWrapper.ReturnDS(para, "usp_GrowEXCEL");
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
                if (fileUpload.HasFiles)
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    bool isFirstFile = true;

                    foreach (HttpPostedFile postedFile in fileUpload.PostedFiles)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(postedFile.FileName);
                        string filePath = Server.MapPath("~/FileUploads/" + postedFile.FileName);

                        postedFile.SaveAs(filePath);

                        FileInfo existingFile = new FileInfo(filePath);

                        using (ExcelPackage package = new ExcelPackage(existingFile))
                        {
                            var worksheet = package.Workbook.Worksheets[0];

                            if (worksheet.Dimension == null)
                                continue;

                            DataTable formattedTable = new DataTable();
                            formattedTable.Columns.Add("TabName", typeof(string));

                            int maxColumns = Math.Min(worksheet.Dimension.End.Column - 1, 9);
                            int rowCount = worksheet.Dimension.End.Row;

                            // Header row
                            for (int col = 2; col <= maxColumns + 1; col++)
                            {
                                var columnName = worksheet.Cells[2, col].Text.Trim();
                                formattedTable.Columns.Add(string.IsNullOrEmpty(columnName) ? $"Column{col}" : columnName);
                            }

                            // Data rows
                            for (int row = 2; row <= rowCount; row++)
                            {
                                if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text))
                                    break;

                                DataRow newRow = formattedTable.NewRow();

                                // Tabname = FileName
                                newRow["TabName"] = fileName;

                                for (int col = 2; col <= maxColumns + 1; col++)
                                {
                                    newRow[col - 1] = worksheet.Cells[row, col].Text;
                                }

                                formattedTable.Rows.Add(newRow);
                            }
                            formattedTable.Rows.RemoveAt(0);

                            if (formattedTable.Rows.Count > 0)
                            {
                                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString))
                                {
                                    connection.Open();

                                    if (isFirstFile)
                                    {
                                        new SqlCommand("TRUNCATE TABLE tbl_GrowTabNames", connection).ExecuteNonQuery();
                                        new SqlCommand("TRUNCATE TABLE tbl_GrowExcelMain", connection).ExecuteNonQuery();
                                        isFirstFile = false;
                                    }
                                }

                                foreach (DataRow row in formattedTable.Rows)
                                {
                                    Instireminagtabs instireminagtabs = new Instireminagtabs
                                    {
                                        Option = "REMAININGTABS",
                                        Tabname = fileName,   // <-- FILE NAME instead of sheet
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

                                // Summary calculation
                                int lineCount = formattedTable.Rows.Count;
                                decimal totalUnits = 0;

                                foreach (DataRow row in formattedTable.Rows)
                                {
                                    if (formattedTable.Columns.Count > 6 &&
                                        decimal.TryParse(row[6]?.ToString(), out decimal unit))
                                    {
                                        totalUnits += unit;
                                    }
                                }

                                InstiMain main = new InstiMain
                                {
                                    TabName = fileName,
                                    ClientCode = "",
                                    Transactions = totalUnits.ToString(),
                                    Qty = totalUnits.ToString(),
                                    Linecount = lineCount.ToString(),
                                };

                                DataTable table1 = INSTIMain(main);
                            }
                        }

                        File.Delete(filePath);
                    }

                    lblMessage.Text = "All files processed successfully!";
                }
                else
                {
                    lblMessage.Text = "Please select files to upload.";
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = "Error: " + ex.Message;
            }
        }

        protected void btnAllinOne_Click(object sender, EventArgs e)
        {
            try
            {
                if (fileUpload.HasFiles)
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (ExcelPackage combinedPackage = new ExcelPackage())
                    {
                        foreach (HttpPostedFile postedFile in fileUpload.PostedFiles)
                        {
                            string fileName = Path.GetFileNameWithoutExtension(postedFile.FileName);
                            string filePath = Server.MapPath("~/FileUploads/" + postedFile.FileName);

                            // Save temporarily
                            postedFile.SaveAs(filePath);

                            using (ExcelPackage sourcePackage = new ExcelPackage(new FileInfo(filePath)))
                            {
                                var sourceSheet = sourcePackage.Workbook.Worksheets[0];

                                if (sourceSheet == null || sourceSheet.Dimension == null)
                                    continue;

                                var newSheet = combinedPackage.Workbook.Worksheets.Add(fileName);

                                int rowCount = sourceSheet.Dimension.End.Row;
                                int colCount = sourceSheet.Dimension.End.Column;

                                for (int row = 1; row <= rowCount; row++)
                                {
                                    for (int col = 1; col <= colCount; col++)
                                    {
                                        newSheet.Cells[row, col].Value = sourceSheet.Cells[row, col].Value;
                                    }
                                }
                            }

                            File.Delete(filePath);
                        }

                        // Send file to browser
                        byte[] fileBytes = combinedPackage.GetAsByteArray();

                        Response.Clear();
                        Response.Buffer = true;
                        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        Response.AddHeader("content-disposition", "attachment; filename=CombinedExcel.xlsx");
                        Response.BinaryWrite(fileBytes);
                        Response.Flush();
                        Response.End();
                    }
                }
                else
                {
                    lblMessage.Text = "Please select files to upload.";
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = "Error: " + ex.Message;
            }
        }
        public DataTable InstiTabs(Instireminagtabs Inputmodel)
        {
            SqlCommand command = new SqlCommand("usp_GrowEXCEL");
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
            SqlCommand command = new SqlCommand("usp_GrowEXCEL");
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

        protected void Groww_AMCReport_RowDataBound(object sender, GridViewRowEventArgs e)
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

        protected void Groww_AMCReport_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            try
            {
                if (e.CommandName == "View")
                {
                    // Ensure the row index is available
                    int rowIndex = Convert.ToInt32(e.CommandArgument);

                    // Access the correct row using the row index
                    GridViewRow row = Groww_AMCReport.Rows[rowIndex];

                    string _TabName = ((Label)row.FindControl("lblSCHEMECODE")).Text;  // Example field
                    string Link = $"GrowwTabName.aspx?TabName={HttpUtility.UrlEncode(_TabName)}";
                    //Response.Redirect(Link, false);
                    //Context.ApplicationInstance.CompleteRequest(); // Prevents ThreadAbortException
                    string script = $"window.open('{Link}', '_blank');";
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "OpenNewTab", script, true);
                }
                else if (e.CommandName == "Download")
                {
                    int rowIndex = Convert.ToInt32(e.CommandArgument);
                    GridViewRow row = Groww_AMCReport.Rows[rowIndex];
                    string _TabName = ((Label)row.FindControl("lblSCHEMECODE")).Text;
                    string _ClientCode = ((TextBox)row.FindControl("txtClientcode")).Text;
                    DataSet dts = new DataSet();

                    SqlParameter[] para = new SqlParameter[2];
                    para[0] = new SqlParameter("@Option", "GROWWEXCELINDI");
                    para[1] = new SqlParameter("@TabName", _TabName);

                    dts = DBWrapper.ReturnDS(para, "usp_GrowEXCEL");
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
                                    // Escape double quotes and commas if needed
                                    //if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
                                    //{
                                    //    field = "\"" + field.Replace("\"", "\"\"") + "\"";
                                    //}
                                    field = "=\"" + field + "\"";

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

                var res = DBWrapper.ReturnDS(para, "usp_GrowEXCEL");
                ShowGrid();
            }
        }
        private int GetInitialQuantityForRow(int rowIndex)
        {

            return rowIndex;
        }
    }
}