using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void btnCSV_Click(object sender, ImageClickEventArgs e)
    {
        Response.ClearContent();
        Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", "Customers.csv"));
        Response.ContentType = "application/text";
        gvdetails.AllowPaging = false;
        gvdetails.DataBind();
        StringBuilder strbldr = new StringBuilder();
        for (int i = 0; i < gvdetails.Columns.Count; i++)
        {
            //separting header columns text with comma operator
            strbldr.Append(gvdetails.Columns[i].HeaderText + ',');
        }
        //appending new line for gridview header row
        strbldr.Append("\n");
        for (int j = 0; j < gvdetails.Rows.Count; j++)
        {
            for (int k = 0; k < gvdetails.Columns.Count; k++)
            {
                //separating gridview columns with comma
                strbldr.Append(gvdetails.Rows[j].Cells[k].Text + ',');
            }
            //appending new line for gridview rows
            strbldr.Append("\n");
        }
        Response.Write(strbldr.ToString());
        Response.End();
    }
}
