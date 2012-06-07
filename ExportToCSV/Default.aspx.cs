using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections.Generic;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void ButtonExport_Click(object sender, EventArgs e)
    {
        List<Person> list = new List<Person>();
        list.Add(new Person("Emad", "Yazdan", 31, 100.45m));
        list.Add(new Person("Ati", "Yazdan", 31, 300.8m));
        list.Add(new Person("Parsa", "Yazdan", 31, 200m));
        list.Add(new Person("Ati", "Taher", 31, 400.20m));
        list.Add(new Person("Eli", "Moghadam", 14, 10.25m));
        CSVExporter.WriteToCSV(list);
    }
}
