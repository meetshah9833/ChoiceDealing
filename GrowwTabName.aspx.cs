using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ChoiceDealing
{
    public partial class GrowwTabName : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string tabName = Request.QueryString["TabName"];
                if (!string.IsNullOrEmpty(tabName))
                {
                    BindGrid(tabName);
                }
            }
        }

        private void BindGrid(string tabName)
        {
            try
            {
                DataSet dts = new DataSet();

                SqlParameter[] para = new SqlParameter[2];
                para[0] = new SqlParameter("@Option", "TABNAMES");
                para[1] = new SqlParameter("@TabName", tabName);

                dts = DBWrapper.ReturnDS(para, "usp_InstiEXCEL");
                if (dts.Tables.Count > 0)
                {
                    GrowwTabNameReport.DataSource = dts;
                    GrowwTabNameReport.DataBind();
                    GrowwTabNameReport.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}