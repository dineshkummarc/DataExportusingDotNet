'@ --------------------------------------------------------------------------------------
'@ Prashant Nayak's Export Button
'@ Copyright (C) 2004 Prashant Nayak. All rights reserved.
'@ http://PrashantNayak.FreeServers.com/ASPNETControls/ExportButton.html
'@ For further reference see http://www.codeproject.com/ and search for 
'@ "Excel Export button control"
'@
'@ Code change history:
'@ ~~~~~~~~~~~~~~~~~~~~
'@ 1. Bug fix - 12/07/2004 
'@    Unicode export and all properties use ViewState as its data store.
'@ --------------------------------------------------------------------------------------
Imports System.ComponentModel
Imports System.Web.UI
Imports System.Text

<Assembly: TagPrefix("PNayak.Web.UI.WebControls", "pnwc")> 
Namespace PNayak.Web.UI.WebControls

    '@ <summary>An extended web button control to download data in Excel or CSV format.
    '@ Drop this button on ASP.NET page, set its Dataview property and when clicked it will
    '@ display a download dialog box for downloading CSV or XLS file.</summary>
    '@ <remarks><c>ExportButton</c> extends standard Button control to implement export
    '@ functionality.
    '@ [Reference: http://support.microsoft.com/default.aspx?scid=kb;EN-US;317719]
    '@ </remarks>
    '@ <example>
    '@   Following is an example of using Export button control on ASP.NET web page.
    '@   <code>
    '@   <![CDATA[
    '@   <PNWebControls:ExportLinkButton id="btnExcel" runat="server" DataView='MyDataView' ExportType="CSV" ></PNWebControls:ExportLinkButton>
    '@   ]]>
    '@  
    '@   'Code Behind
    '@   '-------------------- 
    '@    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '@       Dim ds As Dataset
    '@       if Not Page.IsPostback then
    '@          ds=FillDataset()
    '@       End if
    '@  
    '@       'Set Export button properties
    '@     
    '@       btnExcel.Dataview = ds.Tables("Books").DefaultView
    '@       btnExcel.FileNameToExport = "Books.xls"
    '@       btnExcel.ExportType = PNayak.Web.UI.WebControls.ExportButton.ExportTypeEnum.Excel
    '@       btnExcel.Separator = PNayak.Web.UI.WebControls.ExportButton.SeparatorTypeEnum.TAB
    '@  
    '@    End Sub
    '@ </code>
    '@ </example>

    <DefaultProperty("ExportType"), ToolboxData("<{0}:ExportButton runat=server></{0}:ExportButton>")> _
    Public Class ExportButton
        Inherits System.Web.UI.WebControls.Button

#Region "Declaration Section"
        'Type for export format
        Public Enum ExportTypeEnum
            Excel
            CSV
        End Enum

        'Separator Type
        Public Enum SeparatorTypeEnum
            [Comma]
            [TAB]
        End Enum

        Private _exportData As Dataview
        Private Const C_HTTP_HEADER_CONTENT As String = "Content-Disposition"
        Private Const C_HTTP_ATTACHMENT As String = "attachment;filename="
        Private Const C_HTTP_INLINE As String = "inline;filename="
        Private Const C_HTTP_CONTENT_TYPE_OCTET As String = "application/octet-stream"
        Private Const C_HTTP_CONTENT_TYPE_EXCEL As String = "application/ms-excel"
        Private Const C_HTTP_CONTENT_LENGTH As String = "Content-Length"
        Private Const C_QUERY_PARAM_CRITERIA As String = "Criteria"
        Private Const C_ERROR_NO_RESULT As String = "Data not found."

#End Region

        '@ <summary>
        '@ Assign dataview for exporting data.
        '@ </summary>
        '@ <Remark>Dataset or Datatable can also be used as shown in the example.
        '@ Dataview is used especially when sort order needs to be preseved in the export.</Remark>
        '@ <example>[For Example: ExportButton1.Dataview=Dataset.Tables(0).DefaultView]</example>
        <Bindable(False), Category("ExportButton"), Browsable(False)> _
        Public Property [Dataview]() As Dataview
            Get
                Dim o As Object = ViewState("data")
                If o Is Nothing Then
                    Return Nothing
                Else
                    Dim t As DataTable
                    t = CType(o, DataTable)
                    Return t.DefaultView
                End If
            End Get

            Set(ByVal Value As Dataview)
                ViewState.Add("data", Value.Table)
            End Set
        End Property

        '@ <summary>
        '@ Set filename to save the exported data.
        '@ </summary>
        '@ <example>[For Example: SalesLedger.csv]</example>
        <Bindable(False), Category("ExportButton"), Description("File name to download."), Browsable(True)> _
        Public Property FileNameToExport() As String
            Get
                Dim o As Object = ViewState("filename")
                If o Is Nothing Then
                    Return "ExportData.csv" 'default (CSV extension used as all other default properties return CSV and Comma)
                Else
                    Return CType(o, String)
                End If
            End Get

            Set(ByVal Value As String)
                If Value Is Nothing OrElse Value.Trim.Length = 0 Then
                    Throw New ArgumentNullException("FileNameToExport", "Provide valid file name for export.")
                End If
                ViewState.Add("filename", Value)
            End Set
        End Property


        '@ <summary>
        '@ Set ExportType to either EXCEL or CSV save the exported data.
        '@ </summary>
        '@ <example>[For Example: SalesLedger.xls]</example>
        <Bindable(False), Category("ExportButton"), Description("File format to export (Excel or CSV)."), Browsable(True), DefaultValue("CSV")> _
        Public Property ExportType() As ExportTypeEnum
            Get
                Dim obj As Object = viewstate("exporttype")
                If obj Is Nothing Then
                    Return ExportTypeEnum.CSV 'default
                Else
                    Return CType(obj, ExportTypeEnum)
                End If
            End Get

            Set(ByVal Value As ExportTypeEnum)
                viewstate("exporttype") = Value
            End Set
        End Property

        '@ <summary>
        '@ Set valid separator if "," is not required
        '@ </summary>
        '@ <example>[For Example: SalesLedger.xls]</example>
        <Bindable(False), Category("ExportButton"), Description("Usually Comma can be used for CSV and Tab for Excel.")> _
        Public Property Separator() As SeparatorTypeEnum
            Get
                Dim obj As Object = viewstate("separator")
                If obj Is Nothing Then
                    Return SeparatorTypeEnum.Comma 'default
                Else
                    Return CType(obj, SeparatorTypeEnum)
                End If
            End Get
            Set(ByVal Value As SeparatorTypeEnum)
                viewstate("separator") = Value
            End Set
        End Property

        '@ <summary>
        '@ Valid delimiters like single or double quote can be used.
        '@ </summary>
        <Bindable(False), Category("ExportButton"), Description("Single or Double quoted can be used to delimit data items.")> _
        Public Property Delimiter() As String
            Get
                Dim o As Object = ViewState("delimiter")
                If o Is Nothing Then
                    Return String.Empty
                Else
                    Return CType(o, String)
                End If
            End Get

            Set(ByVal Value As String)
                ViewState.Add("delimiter", Value)
            End Set
        End Property

        '@ <summary>
        '@ Valid Character set for exporting international characters
        '@ </summary>
        <Bindable(False), Category("ExportButton"), Description("Used for various character sets. (Ex: 'iso-8859-2' used for Cenral Europian Characters)")> _
        Public Property CharSet() As String
            Get
                Dim o As Object = ViewState("charset")
                If o Is Nothing Then
                    Return String.Empty
                Else
                    Return CType(o, String)
                End If
            End Get

            Set(ByVal Value As String)
                ViewState.Add("charset", Value)
            End Set
        End Property


        'Overriding the click event for executing custom export function to avoid
        'building separate ASPX page export functionality.
        Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
            Call CustomExport()
        End Sub

        ' Exports data to a custom format
        Private Sub CustomExport()
            Dim response As System.Web.HttpResponse = System.Web.HttpContext.Current.Response
            response.Clear()
            ' Add the header that specifies the default filename 
            ' for the Download/SaveAs dialog
            response.AddHeader(C_HTTP_HEADER_CONTENT, C_HTTP_ATTACHMENT & FileNameToExport)

            ' Specify that the response is a stream that cannot be read _
            ' by the client and must be downloaded
            If ExportType = ExportTypeEnum.CSV Then
                response.ContentType = C_HTTP_CONTENT_TYPE_OCTET
            Else
                response.ContentType = C_HTTP_CONTENT_TYPE_EXCEL
            End If

            Dim _exportContent As String = String.Empty
            _exportData = [Dataview]
            If (Not _exportData Is Nothing) AndAlso _exportData.Table.Rows.Count > 0 Then
                If Separator = SeparatorTypeEnum.TAB Then
                    _exportContent = ConvertDataViewToString(_exportData, Delimiter, vbTab)
                Else
                    _exportContent = ConvertDataViewToString(_exportData, Delimiter, ",")
                End If
            End If
            If _exportContent.Length <= 0 Then
                _exportContent = C_ERROR_NO_RESULT
            End If

            Dim Encoding As New System.Text.UnicodeEncoding
            response.AddHeader(C_HTTP_CONTENT_LENGTH, Encoding.GetByteCount(_exportContent).ToString())
            response.BinaryWrite(Encoding.GetBytes(_exportContent))
            response.Charset = CharSet

            ' Stop execution of the current page
            response.End()
        End Sub



    End Class
End Namespace