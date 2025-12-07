using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
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
            ShowGrid();
        }

        private void ShowGrid()
        {
            try
            {
                DataSet dts = new DataSet();

                SqlParameter[] para = new SqlParameter[1];
                para[0] = new SqlParameter("@OPTION", "DSPVIEW");

                dts = DBWrapper.ReturnDS(para, "proc_DSP");
                if (dts.Tables.Count > 0)
                {
                    DSPReport.DataSource = dts;
                    DSPReport.DataBind();
                    DSPReport.Visible = true;
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
                                    if (formattedTable.Columns.Count > 6 && decimal.TryParse(row[5]?.ToString(), out decimal unit))
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
            try
            {
                DataSet dts = new DataSet();

                SqlParameter[] para = new SqlParameter[1];
                para[0] = new SqlParameter("@Option", "VIEW");

                dts = DBWrapper.ReturnDS(para, "proc_DSP");
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
                        Response.AddHeader("content-disposition", "attachment; filename=DSPFile.csv");

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

        protected void DSPReport_RowDataBound(object sender, GridViewRowEventArgs e)
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

        protected void DSPReport_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            try
            {
                if (e.CommandName == "View")
                {
                    // Ensure the row index is available
                    int rowIndex = Convert.ToInt32(e.CommandArgument);

                    // Access the correct row using the row index
                    GridViewRow row = DSPReport.Rows[rowIndex];

                    string _TabName = ((Label)row.FindControl("lblSCHEMECODE")).Text;  // Example field
                    string Link = $"DSPTabName.aspx?TabName={HttpUtility.UrlEncode(_TabName)}";
                    //Response.Redirect(Link, false);
                    //Context.ApplicationInstance.CompleteRequest(); // Prevents ThreadAbortException
                    string script = $"window.open('{Link}', '_blank');";
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "OpenNewTab", script, true);
                }
                else if (e.CommandName == "Download")
                {
                    int rowIndex = Convert.ToInt32(e.CommandArgument);
                    GridViewRow row = DSPReport.Rows[rowIndex];
                    string _TabName = ((Label)row.FindControl("lblSCHEMECODE")).Text;
                    string _ClientCode = ((TextBox)row.FindControl("txtClientcode")).Text;
                    DataSet dts = new DataSet();

                    SqlParameter[] para = new SqlParameter[2];
                    para[0] = new SqlParameter("@OPTION", "DSPEXCELINDI");
                    para[1] = new SqlParameter("@SCHEMECODE", _TabName);

                    dts = DBWrapper.ReturnDS(para, "proc_DSP");
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
                //int initialQuantity = GetInitialQuantityForRow(Convert.ToInt32(lblInitialQUANTITY.Text));
                decimal initialQuantity = GetInitialQuantityForRow(Convert.ToDecimal(lblInitialQUANTITY.Text));
                decimal updatedQuantity = initialQuantity * noOfBaskets;
                lblQuantity.Text = updatedQuantity.ToString();

                // Save to database
                SqlParameter[] para = new SqlParameter[7];
                para[0] = new SqlParameter("@OPTION", "UPDATEDSPMAIN");
                para[1] = new SqlParameter("@CLIENTCODE", clientCode);
                para[2] = new SqlParameter("@BUYSELL", buySell);
                para[3] = new SqlParameter("@QUANTITY", lblQuantity.Text);
                para[4] = new SqlParameter("@SCHEMECODE", SqlDbType.VarChar, 100)
                { Value = (schemeCode ?? "").Trim() };
                para[5] = new SqlParameter("@NOOFBASKETS", noOfBaskets);
                para[6] = new SqlParameter("@TradeInstructions", TradeInstructions);

                var res = DBWrapper.ReturnDS(para, "proc_DSP");
                ShowGrid();
            }
        }
        private decimal GetInitialQuantityForRow(decimal rowIndex)
        {

            return rowIndex;
        }
    }
}