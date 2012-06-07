Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Data.OleDb
Imports RKLib.ExportData

Public Class EmployeeInfo
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents lblScenario1 As System.Web.UI.WebControls.Label
    Protected WithEvents dgEmployee1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents btnExport1 As System.Web.UI.WebControls.Button
    Protected WithEvents lblScenario2 As System.Web.UI.WebControls.Label
    Protected WithEvents dgEmployee2 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents btnExport2 As System.Web.UI.WebControls.Button
    Protected WithEvents lblScenario3 As System.Web.UI.WebControls.Label
    Protected WithEvents dgEmployee3 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents btnExport3 As System.Web.UI.WebControls.Button
    Protected WithEvents lblError As System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it Imports the code editor.
        InitializeComponent()
    End Sub

#End Region

    Dim dsEmployee As New DataSet

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            lblError.Text = ""
            If Not IsPostBack Then
                get_AllEmployees()
                Session("dsEmployee") = dsEmployee

                dgEmployee1.DataSource = dsEmployee.Tables("Employee")
                dgEmployee1.DataBind()

                dgEmployee2.DataSource = dsEmployee.Tables("Employee")
                dgEmployee2.DataBind()

                dgEmployee3.DataSource = dsEmployee.Tables("Employee")
                dgEmployee3.DataBind()
            End If
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try

    End Sub

    Private Function get_AllEmployees() As DataSet

        Try
            ' Get the employee details
            Dim strSql As String = "SELECT EmployeeID, LastName, FirstName, Format(BirthDate, 'dd-MMM-yyyy') As BirthDate, Format(HireDate, 'dd-MMM-yyyy') As HireDate, Address, PostalCode FROM Employees"
            Dim objConn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.MapPath("Database\\Employee.mdb"))
            Dim daEmp As New OleDbDataAdapter(strSql, objConn)
            daEmp.Fill(dsEmployee, "Employee")
            Return dsEmployee
        Catch Ex As Exception
            Throw Ex
        End Try

    End Function

    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click

        ' Export all the details

        Try
            ' Get the datatable to export			
            Dim dtEmployee As DataTable = (CType(Session("dsEmployee"), DataSet)).Tables("Employee").Copy()

            ' Export all the details to CSV
            Dim objExport As New RKLib.ExportData.Export("Web")
            objExport.ExportDetails(dtEmployee, Export.ExportFormat.CSV, "EmployeesInfo1.csv")
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try

    End Sub

    Private Sub btnExport2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport2.Click

        ' Export the details of specified columns

        Try
            ' Get the datatable to export			
            Dim dtEmployee As DataTable = (CType(Session("dsEmployee"), DataSet)).Tables("Employee").Copy()

            ' Specify the column list to export
            Dim iColumns() As Integer = New Integer() {1, 2, 3, 5, 6}

            ' Export the details of specified columns to Excel
            Dim objExport As New RKLib.ExportData.Export("Web")
            objExport.ExportDetails(dtEmployee, iColumns, Export.ExportFormat.Excel, "EmployeesInfo2.xls")
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try

    End Sub

    Private Sub btnExport3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport3.Click

        ' Export the details of specified columns with specified headers

        Try
            ' Get the datatable to export			
            Dim dtEmployee As DataTable = (CType(Session("dsEmployee"), DataSet)).Tables("Employee").Copy()

            ' Specify the column list and headers to export
            Dim iColumns() As Integer = New Integer() {1, 2, 3, 5, 6}
            Dim sHeaders() As String = New String() {"LastName", "FirstName", "DOB", "Address", "ZipCode"}

            ' Export the details of specified columns with specified headers to CSV
            Dim objExport As New RKLib.ExportData.Export("Web")
            objExport.ExportDetails(dtEmployee, iColumns, sHeaders, Export.ExportFormat.CSV, "EmployeesInfo3.csv")
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try

    End Sub

End Class
