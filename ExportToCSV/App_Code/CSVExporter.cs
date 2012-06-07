using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Collections.Generic;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Text;

/// <summary>
/// Summary description for CSVExporter
/// </summary>
public class CSVExporter
{
    public static void WriteToCSV(List<Person> personList)
    {
        string attachment = "attachment; filename=PerosnList.csv";
        HttpContext.Current.Response.Clear();
        HttpContext.Current.Response.ClearHeaders();
        HttpContext.Current.Response.ClearContent();
        HttpContext.Current.Response.AddHeader("content-disposition", attachment);
        HttpContext.Current.Response.ContentType = "text/csv";
        HttpContext.Current.Response.AddHeader("Pragma", "public");

        WriteColumnName();

        foreach (Person item in personList)
        {
            WriteUserInfo(item);
        }

        HttpContext.Current.Response.End();
    }

    private static void WriteUserInfo(Person item)
    {
        StringBuilder strb = new StringBuilder();

        AddComma(item.Name, strb);
        AddComma(item.Family, strb);
        AddComma(item.Age.ToString(), strb);
        AddComma(string.Format("{0:C2}", item.Salary), strb);

        HttpContext.Current.Response.Write(strb.ToString());
        HttpContext.Current.Response.Write(Environment.NewLine);
    }

    private static void AddComma(string item, StringBuilder strb)
    {
        strb.Append(item.Replace(',', ' '));
        strb.Append(" ,");
    }
    private static void WriteColumnName()
    {
        string str = "Name, Family, Age, Salary";
        HttpContext.Current.Response.Write(str);
        HttpContext.Current.Response.Write(Environment.NewLine);
    }
}

