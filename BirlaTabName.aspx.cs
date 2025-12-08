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
    public partial class BirlaTabName : System.Web.UI.Page
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
                para[0] = new SqlParameter("@OPTION", "TABNAMES");
                para[1] = new SqlParameter("@SCHEMECODE", tabName);

                dts = DBWrapper.ReturnDS(para, "proc_BIRLA");
                if (dts.Tables.Count > 0)
                {
                    BirlaTabNameReport.DataSource = dts;
                    BirlaTabNameReport.DataBind();
                    BirlaTabNameReport.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}