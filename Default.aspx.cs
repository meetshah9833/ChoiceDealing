using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ChoiceDealing
{
    public partial class _Default : Page
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
            //try
            //{
            //    DataSet dts = new DataSet();

            //    SqlParameter[] para = new SqlParameter[1];
            //    para[0] = new SqlParameter("@Option", "ORMVIEW");

            //    dts = DBWrapper.ReturnDS(para, "[LKP_Middleware_Config].[dbo].[usp_InstiEXCEL]");
            //    if (dts.Tables.Count > 0)
            //    {
            //        Moti_ORMSReport.DataSource = dts;
            //        Moti_ORMSReport.DataBind();
            //        Moti_ORMSReport.Visible = true;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    if (lblMotiEQFile.HasFile && fileUpload.HasFile)
            //    {
            //        // Set EPPlus license
            //        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //        // Save and process lblMotiEQFile (EQT file)
            //        string eqtFileName = Path.GetFileName(lblMotiEQFile.PostedFile.FileName);
            //        string eqtFilePath = Server.MapPath("~/CTCL_Certificates/" + eqtFileName);
            //        lblMotiEQFile.SaveAs(eqtFilePath);

            //        DataTable formattedTable = new DataTable();
            //        using (var package = new ExcelPackage(new FileInfo(eqtFilePath)))
            //        {
            //            var worksheet = package.Workbook.Worksheets[0];
            //            int colCount = worksheet.Dimension.End.Column;
            //            int rowCount = worksheet.Dimension.End.Row;

            //            Dictionary<string, int> columnIndexes = new Dictionary<string, int>();

            //            // Dynamically add columns from the header row
            //            for (int col = 1; col <= colCount; col++)
            //            {
            //                string header = worksheet.Cells[1, col].Text?.Trim();
            //                if (!string.IsNullOrEmpty(header) && !formattedTable.Columns.Contains(header))
            //                {
            //                    columnIndexes[header] = col;
            //                    formattedTable.Columns.Add(header);
            //                }
            //            }

            //            // Fill the DataTable with data from the sheet
            //            for (int row = 2; row <= rowCount; row++)
            //            {
            //                DataRow dr = formattedTable.NewRow();
            //                foreach (var kvp in columnIndexes)
            //                {
            //                    string colName = kvp.Key;
            //                    int colIndex = kvp.Value;
            //                    dr[colName] = worksheet.Cells[row, colIndex].Text?.Trim();
            //                }
            //                formattedTable.Rows.Add(dr);
            //            }
            //        }

            //        using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString))
            //        {
            //            connection.Open();
            //            SqlCommand truncateCmd = new SqlCommand("TRUNCATE TABLE tbl_MotiORMExcel", connection);
            //            truncateCmd.ExecuteNonQuery();

            //            foreach (DataRow row in formattedTable.Rows)
            //            {
            //                MotiEQExcel instiMain = new MotiEQExcel
            //                {
            //                    Date = row[0]?.ToString(),
            //                    BuySell = row[1]?.ToString(),
            //                    ScripName = row[2]?.ToString(),
            //                    ISIN = row[3]?.ToString(),
            //                    Exchange = row[4]?.ToString(),
            //                    Qty = row[5]?.ToString(),
            //                    MarketPrice = row[6]?.ToString(),
            //                    BrokerSEBINo = row[7]?.ToString(),
            //                    Scheme = row[8]?.ToString(),
            //                    UCC = row[9]?.ToString(),
            //                    Brokerage = row[10]?.ToString(),
            //                    STT_Amt = row[11]?.ToString(),
            //                    NetAmount = row[12]?.ToString(),
            //                    AssetClass = row[13]?.ToString(),
            //                    YTM = row[14]?.ToString(),
            //                    InstrumentHoldingType = row[15]?.ToString(),
            //                    FXRate = row[16]?.ToString(),
            //                    AccruedInterest = row[17]?.ToString(),
            //                    CounterPartyShortName = row[18]?.ToString()
            //                };

            //                MotiORMEq(instiMain);
            //            }
            //        }

            //        File.Delete(eqtFilePath);
            //        ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('File Uploaded!');", true);

            //        // Process fileUpload (OMS file)
            //        string omsFileName = Path.GetFileName(fileUpload.PostedFile.FileName);
            //        string omsFilePath = Server.MapPath("~/CTCL_Certificates/" + omsFileName);
            //        fileUpload.SaveAs(omsFilePath);
            //        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //        DataTable omsTable = new DataTable();
            //        using (var package = new ExcelPackage(new FileInfo(omsFilePath)))
            //        {
            //            var worksheet = package.Workbook.Worksheets[0];
            //            int colCount = worksheet.Dimension.End.Column;
            //            int rowCount = worksheet.Dimension.End.Row;

            //            for (int col = 1; col <= colCount; col++)
            //            {
            //                string header = worksheet.Cells[1, col].Text?.Trim();
            //                if (!string.IsNullOrEmpty(header))
            //                    omsTable.Columns.Add(header);
            //                else
            //                    omsTable.Columns.Add("Column" + col); // Fallback if header is empty
            //            }

            //            for (int row = 2; row <= rowCount; row++)
            //            {
            //                bool isEmptyRow = true;
            //                DataRow dr = omsTable.NewRow();

            //                for (int col = 1; col <= colCount; col++)
            //                {
            //                    var cellValue = worksheet.Cells[row, col]?.Text?.Trim();
            //                    if (!string.IsNullOrEmpty(cellValue))
            //                        isEmptyRow = false;

            //                    dr[col - 1] = cellValue;
            //                }

            //                if (isEmptyRow)
            //                    break; // Stop processing if entire row is empty

            //                omsTable.Rows.Add(dr);
            //            }

            //        }

            //        using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString))
            //        {
            //            connection.Open();
            //            SqlCommand truncateCmd = new SqlCommand("TRUNCATE TABLE MotiEQExcel", connection);
            //            truncateCmd.ExecuteNonQuery();

            //            foreach (DataRow row in omsTable.Rows)
            //            {
            //                MotiORMExcel instiMain = new MotiORMExcel
            //                {
            //                    Exchange = row["Exchange"]?.ToString(),
            //                    Sub_Ac_Code_Client_Id = row["Sub A/c Code/Client Id"]?.ToString(),
            //                    Sub_Ac_Name = row["Sub A/c Name"]?.ToString(),
            //                    Scrip_Name = row["Scrip Name"]?.ToString(),
            //                    Buy_Sell = row["Buy/Sell"]?.ToString(),
            //                    Order_Type = row["Order Type"]?.ToString(),
            //                    Order_Qty = row["Order Qty"]?.ToString(),
            //                    ISIN_Code = row["ISIN Code"]?.ToString(),
            //                    Settlor_NSE = row["Settlor (NSE)"]?.ToString(),
            //                    OMSID = row["OMSID"]?.ToString(),
            //                    Date = row["Date"]?.ToString(),
            //                    Mapped_Scrip_Code = row["Mapped Scrip Code"]?.ToString()
            //                };

            //                MotiEq(instiMain);
            //            }
            //        }

            //        File.Delete(omsFilePath);
            //        ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('File Uploaded!');", true);
            //    }
            //    else
            //    {
            //        ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('Please select the files for uploading!');", true);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    ScriptManager.RegisterStartupScript(this, GetType(), "alertMessage", $"alert('Error: {ex.Message}');", true);
            //}
        }

        public DataTable MotiORMEq(MotiEQExcel motiEQExcel)
        {
            SqlCommand command = new SqlCommand("usp_InstiEXCEL");
            command.CommandType = CommandType.StoredProcedure;

            command.Parameters.Add("@Option", SqlDbType.VarChar).Value = "ORMEXCELDATA";
            command.Parameters.Add("@Date", SqlDbType.VarChar).Value = motiEQExcel.Date;
            command.Parameters.Add("@BuySell", SqlDbType.VarChar).Value = motiEQExcel.BuySell;
            command.Parameters.Add("@SecurityName", SqlDbType.VarChar).Value = motiEQExcel.ScripName;
            command.Parameters.Add("@ISIN", SqlDbType.VarChar).Value = motiEQExcel.ISIN;
            command.Parameters.Add("@Exchange", SqlDbType.VarChar).Value = motiEQExcel.Exchange;
            command.Parameters.Add("@Qty", SqlDbType.VarChar).Value = motiEQExcel.Qty;
            command.Parameters.Add("@MarketPrice", SqlDbType.VarChar).Value = motiEQExcel.MarketPrice;
            command.Parameters.Add("@BrokerSEBINo", SqlDbType.VarChar).Value = motiEQExcel.BrokerSEBINo;
            command.Parameters.Add("@Scheme", SqlDbType.VarChar).Value = motiEQExcel.Scheme;
            command.Parameters.Add("@UCC", SqlDbType.VarChar).Value = motiEQExcel.UCC;
            command.Parameters.Add("@Brokerage", SqlDbType.VarChar).Value = motiEQExcel.Brokerage;
            command.Parameters.Add("@Amt", SqlDbType.VarChar).Value = motiEQExcel.Amt;
            command.Parameters.Add("@STT_Amt", SqlDbType.VarChar).Value = motiEQExcel.STT_Amt;
            command.Parameters.Add("@NetAmount", SqlDbType.VarChar).Value = motiEQExcel.NetAmount;
            command.Parameters.Add("@AssetClass", SqlDbType.VarChar).Value = motiEQExcel.AssetClass;
            command.Parameters.Add("@YTM", SqlDbType.VarChar).Value = motiEQExcel.YTM;
            command.Parameters.Add("@InstrumentHoldingType", SqlDbType.VarChar).Value = motiEQExcel.InstrumentHoldingType;
            command.Parameters.Add("@FXRate", SqlDbType.VarChar).Value = motiEQExcel.FXRate;
            command.Parameters.Add("@AccruedInterest", SqlDbType.VarChar).Value = motiEQExcel.AccruedInterest;
            command.Parameters.Add("@CounterPartyShortName", SqlDbType.VarChar).Value = motiEQExcel.CounterPartyShortName;
            return ExecuteStoredProcedure(ConfigurationManager.ConnectionStrings["ConnMiddleWare"].ConnectionString, command);
        }

        public DataTable MotiEq(MotiORMExcel motiEQExcel)
        {
            SqlCommand command = new SqlCommand("usp_InstiEXCEL");
            command.CommandType = CommandType.StoredProcedure;

            command.Parameters.Add("@Option", SqlDbType.VarChar).Value = "MotiEQ";
            command.Parameters.Add("@Exchange", SqlDbType.VarChar).Value = motiEQExcel.Exchange;
            command.Parameters.Add("@Sub_Ac_Code_Client_Id", SqlDbType.VarChar).Value = motiEQExcel.Sub_Ac_Code_Client_Id;
            command.Parameters.Add("@Sub_Ac_Name", SqlDbType.VarChar).Value = motiEQExcel.Sub_Ac_Name;
            command.Parameters.Add("@Scrip_Name", SqlDbType.VarChar).Value = motiEQExcel.Scrip_Name;
            command.Parameters.Add("@Buy_Sell", SqlDbType.VarChar).Value = motiEQExcel.Buy_Sell;
            command.Parameters.Add("@Order_Type", SqlDbType.VarChar).Value = motiEQExcel.Order_Type;
            command.Parameters.Add("@Order_Qty", SqlDbType.VarChar).Value = motiEQExcel.Order_Qty;
            command.Parameters.Add("@ISIN_Code", SqlDbType.VarChar).Value = motiEQExcel.ISIN_Code;
            command.Parameters.Add("@Settlor_NSE", SqlDbType.VarChar).Value = motiEQExcel.Settlor_NSE;
            command.Parameters.Add("@OMSID", SqlDbType.VarChar).Value = motiEQExcel.OMSID;
            command.Parameters.Add("@Date", SqlDbType.VarChar).Value = motiEQExcel.Date;
            command.Parameters.Add("@Mapped_Scrip_Code", SqlDbType.VarChar).Value = motiEQExcel.Mapped_Scrip_Code;

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
        public class MotiEQExcel
        {
            public string Date { get; set; }
            public string BuySell { get; set; }
            public string ScripName { get; set; }
            public string ISIN { get; set; }
            public string Exchange { get; set; }
            public string Qty { get; set; }
            public string MarketPrice { get; set; }
            public string BrokerSEBINo { get; set; }
            public string Scheme { get; set; }
            public string UCC { get; set; }
            public string Brokerage { get; set; }
            public string Amt { get; set; }
            public string STT_Amt { get; set; }
            public string NetAmount { get; set; }
            public string AssetClass { get; set; }
            public string YTM { get; set; }
            public string InstrumentHoldingType { get; set; }
            public string FXRate { get; set; }
            public string AccruedInterest { get; set; }
            public string CounterPartyShortName { get; set; }
        }

        public class MotiORMExcel
        {
            public string Tree { get; set; }
            public string Algo_Type { get; set; }
            public string Exchange { get; set; }
            public string Sub_Ac_Code_Client_Id { get; set; }
            public string Sub_Ac_Name { get; set; }
            public string Scrip_Name { get; set; }
            public string Buy_Sell { get; set; }
            public string Order_Type { get; set; }
            public string Status { get; set; }
            public string Order_Qty { get; set; }
            public string Order_Price { get; set; }
            public string Soft_Limit { get; set; }
            public string Trade_Instruction { get; set; }
            public string Released_Qty { get; set; }
            public string Unreleased_Qty { get; set; }
            public string Executed_Qty { get; set; }
            public string NSE_ATP { get; set; }
            public string Arrival_Price { get; set; }
            public string End_Time { get; set; }
            public string Dealing_Instruction { get; set; }
            public string ISIN_Code { get; set; }
            public string UnReleased_Value { get; set; }
            public string Executed_Value { get; set; }
            public string Time { get; set; }
            public string Algo_Desc { get; set; }
            public string Main_Ac_Name { get; set; }
            public string Settlor_NSE { get; set; }
            public string Basket_Id { get; set; }
            public string OMSID { get; set; }
            public string Parent_Ac_Name { get; set; }
            public string Part_Type { get; set; }
            public string Short_Code { get; set; }
            public string Date { get; set; }
            public string Parent_Ac_Code { get; set; }
            public string Main_Ac_Code { get; set; }
            public string Mapped_Scrip_Code { get; set; }
            public string Instrument_Name { get; set; }
            public string Expiry_Date { get; set; }
            public string Strike_Price { get; set; }
            public string Option_Type { get; set; }
            public string Dollar_Figure { get; set; }
            public string Percentage_Volume { get; set; }
            public string PV_Comparision { get; set; }
            public string Exchange_Pending_Qty { get; set; }
            public string Percentage_Pending { get; set; }
            public string Balance_Qty { get; set; }
            public string BSE_ATP { get; set; }
            public string Conversion_Rate { get; set; }
            public string Exchange_Ratio { get; set; }
            public string Last_Modified_Time { get; set; }
            public string Slice_Qty { get; set; }
            public string FII_Watch { get; set; }
            public string NRI_Watch { get; set; }
            public string NSE_Executed_Qty { get; set; }
            public string BSE_Executed_Qty { get; set; }
            public string Route_Dealer { get; set; }
            public string Vol_Phase { get; set; }
            public string Block_Fig { get; set; }
            public string Combined_ATP { get; set; }
            public string Valid_Till { get; set; }
            public string Previous_Executed_Qty { get; set; }
            public string Previous_Value { get; set; }
            public string Todays_Combined_ATP { get; set; }
            public string Previous_Combined_ATP { get; set; }
            public string Validity { get; set; }
            public string NSE_Manual_Fill_Qty { get; set; }
            public string BSE_Manual_Fill_Qty { get; set; }
            public string Order_Value { get; set; }
            public string Released_Value { get; set; }
            public string Exchange_Pending_Value { get; set; }
            public string Balance_Value { get; set; }
            public string Slice_Value { get; set; }
            public string NSE_Executed_Value { get; set; }
            public string BSE_Executed_Value { get; set; }
            public string EOMSID { get; set; }
            public string Booking_Ref { get; set; }
            public string Target_Dealer_Id { get; set; }
            public string Source_Dealer_Id { get; set; }
            public string Grab_Status { get; set; }
            public string Pvt { get; set; }
            public string Sales_Trader_Id { get; set; }
            public string Assign_Dealer_Id { get; set; }
            public string Display_Price { get; set; }
            public string Execution_Style { get; set; }
            public string DQ_Percentage { get; set; }
            public string Display_Size { get; set; }
            public string Max_POV { get; set; }
            public string Start_Time { get; set; }
            public string From_Price1 { get; set; }
            public string To_Price1 { get; set; }
            public string POV_Percent1 { get; set; }
            public string From_Price2 { get; set; }
            public string To_Price2 { get; set; }
            public string POV_Percent2 { get; set; }
            public string From_Price3 { get; set; }
            public string To_Price3 { get; set; }
            public string POV_Percent3 { get; set; }
            public string Distribution_Type { get; set; }
            public string Max_Slice_Qty { get; set; }
            public string Block_Order_Qty { get; set; }
            public string Max_Block_Order_Qty { get; set; }
            public string Threshold_Percent { get; set; }
            public string Initial_Participation { get; set; }
            public string Max_Participation { get; set; }
            public string Int_Tick { get; set; }
            public string NSE_Released_Qty { get; set; }
            public string NSE_Released_Value { get; set; }
            public string BSE_Released_Qty { get; set; }
            public string BSE_Released_Value { get; set; }
            public string Combined_Released_Qty { get; set; }
            public string Combined_Released_Value { get; set; }
            public string NSE_Pending_Qty { get; set; }
            public string BSE_Pending_Qty { get; set; }
            public string Distribution { get; set; }
            public string Completion_Quantity_Percentage { get; set; }
            public string Completion_Price { get; set; }
            public string Switch_Price { get; set; }
            public string Min_Participation { get; set; }
            public string Change_in_Price1_Percentage { get; set; }
            public string Change_in_Participation1_Percentage { get; set; }
            public string Change_in_Price2_Percentage { get; set; }
            public string Change_in_Participation2_Percentage { get; set; }
            public string Parent_OMSID { get; set; }
            public string NOE_ID { get; set; }
            public string Sender_Details { get; set; }
            public string NSE_Incr_Vol { get; set; }
            public string BSE_Incr_Vol { get; set; }
            public string Combined_Incr_Vol { get; set; }
            public string Incr_Participation_Percentage { get; set; }
            public string Executed_Qty_Percentage { get; set; }
            public string Incr_Algo_Quantity { get; set; }
            public string Execution_Exchange { get; set; }
            public string Bottom_Price { get; set; }
            public string Price_Jump { get; set; }
            public string Available_Quantity { get; set; }
            public string Block_Volume { get; set; }
        }
        protected void btnDownload_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    DataSet dts = new DataSet();

            //    SqlParameter[] para = new SqlParameter[1];
            //    para[0] = new SqlParameter("@Option", "ORMDOWNLOAD");

            //    dts = DBWrapper.ReturnDS(para, "[LKP_Middleware_Config].[dbo].[usp_InstiEXCEL]");
            //    if (dts.Tables.Count > 0 && dts.Tables[0].Rows.Count > 0)
            //    {
            //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //        using (var package = new OfficeOpenXml.ExcelPackage()) // EPPlus Library
            //        {
            //            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
            //            worksheet.Cells.LoadFromDataTable(dts.Tables[0], true); // Load data

            //            Response.Clear();
            //            Response.ContentType = "text/csv";
            //            Response.AddHeader("content-disposition", "attachment; filename=MotiEQOMSFile.csv");

            //            System.Text.StringBuilder csvContent = new StringBuilder();
            //            DataTable dt = dts.Tables[0];
            //            for (int i = 0; i < dt.Columns.Count; i++)
            //            {
            //                csvContent.Append(dt.Columns[i].ColumnName);
            //                if (i < dt.Columns.Count - 1)
            //                    csvContent.Append(",");
            //            }
            //            csvContent.AppendLine();
            //            foreach (DataRow row in dt.Rows)
            //            {
            //                for (int i = 0; i < dt.Columns.Count; i++)
            //                {
            //                    csvContent.Append(row[i]?.ToString().Replace(",", " ")); // Prevent CSV format issues
            //                    if (i < dt.Columns.Count - 1)
            //                        csvContent.Append(",");
            //                }
            //                csvContent.AppendLine();
            //            }
            //            //using (var stream = new MemoryStream())
            //            //{
            //            //    package.SaveAs(stream);
            //            //    stream.WriteTo(Response.OutputStream);
            //            //}
            //            byte[] csvBytes = Encoding.UTF8.GetBytes(csvContent.ToString());

            //            Response.OutputStream.Write(csvBytes, 0, csvBytes.Length);
            //            Response.Flush();
            //            Response.End();
            //        }
            //    }
            //    else
            //    {
            //        // Show message if no data found
            //        ScriptManager.RegisterStartupScript(this, GetType(), "alertMessage", "alert('No data found to export!');", true);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    ScriptManager.RegisterStartupScript(this, GetType(), "alertMessage", "alert('" + ex + "');", true);
            //}
        }
    }
}