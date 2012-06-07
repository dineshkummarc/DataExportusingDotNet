using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// Summary description for Person
/// </summary>
public class Person
{
    private string m_Name;

    public string Name
    {
        get { return m_Name; }
        set { m_Name = value; }
    }
    private string m_Family;

    public string Family
    {
        get { return m_Family; }
        set { m_Family = value; }
    }
    private int m_Age;

    public int Age
    {
        get { return m_Age; }
        set { m_Age = value; }
    }

    private decimal m_Salary;

    public decimal Salary
    {
        get { return m_Salary; }
        set { m_Salary = value; }
    }

	public Person()
	{
		//
		// TODO: Add constructor logic here
		//
	}
    public Person(string name,string family,int age,decimal salary)
    {
        m_Name = name;
        m_Family = family;
        m_Age = age;
        m_Salary = salary;
    }
}
