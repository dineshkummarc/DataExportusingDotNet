Imports System.Data.OleDb
Public Class WebForm1
    Inherits System.Web.UI.Page
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents lbPublishers As System.Web.UI.WebControls.ListBox
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents dgBooks As System.Web.UI.WebControls.DataGrid
    Protected WithEvents lnkbtnCSV As PNayak.Web.UI.WebControls.ExportLinkButton
    Protected WithEvents lnkbtnExcel As PNayak.Web.UI.WebControls.ExportLinkButton
    Protected WithEvents btnLoad As System.Web.UI.WebControls.Button
    Protected WithEvents btnCSV As PNayak.Web.UI.WebControls.ExportButton
    Protected WithEvents btnExcel As PNayak.Web.UI.WebControls.ExportButton
    Protected WithEvents btnTestPostback As System.Web.UI.WebControls.Button

    Dim _ds As DataSet = New DataSet()


#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then
            filldataset()
            dgBooks.DataSource = _ds.Tables("Books")
            dgBooks.DataBind()

            lbPublishers.DataSource = _ds.Tables("Publishers")
            lbPublishers.DataTextField = _ds.Tables("Publishers").Columns(1).ColumnName
            lbPublishers.DataBind()

            btnCSV.Dataview = _ds.Tables("Books").DefaultView

            btnExcel.Dataview = _ds.Tables("Books").DefaultView
            btnExcel.FileNameToExport = "MyBooks.xls"
            btnExcel.ExportType = PNayak.Web.UI.WebControls.ExportButton.ExportTypeEnum.Excel
            btnExcel.Separator = PNayak.Web.UI.WebControls.ExportButton.SeparatorTypeEnum.TAB
            btnExcel.Delimiter = """"

            lnkbtnCSV.Dataview = _ds.Tables("Publishers").DefaultView

            lnkbtnExcel.Dataview = _ds.Tables("Publishers").DefaultView
            lnkbtnExcel.CharSet = "iso-8859-2" 'Test central europian characters
        End If
    End Sub


    Sub filldataset()

        Dim conn As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.MapPath("books.mdb"))
        conn.Open()

        Dim SQL As String = "SELECT BookID, Title FROM Books"
        Dim adapter1 As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(SQL, conn))
        adapter1.Fill(_ds, "Books")

        SQL = "SELECT * FROM Publishers"
        adapter1 = New OleDbDataAdapter(New OleDbCommand(SQL, conn))
        adapter1.Fill(_ds, "Publishers")

        conn.Close()

    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If _ds Is Nothing OrElse _ds.Tables.Count = 0 Then
            filldataset()
        End If
        btnCSV.Dataview = _ds.Tables("Publishers").DefaultView
        lnkbtnCSV.Dataview = _ds.Tables("Books").DefaultView
    End Sub
End Class
