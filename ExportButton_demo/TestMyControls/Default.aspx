<%@ Register TagPrefix="pnwc" Namespace="PNayak.Web.UI.WebControls" Assembly="PNayak.Web.UI.WebControls.ExportButton" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Default.aspx.vb" Inherits="TestMyControls.WebForm1"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Prashant Nayak's Export Button control</title>
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<P><span style="FONT-WEIGHT: bold; FONT-SIZE: 14pt; FONT-FAMILY: Verdana; BACKGROUND-COLOR: silver; BORDER-BOTTOM-STYLE: solid">Reusable 
				Export button control for&nbsp;Excel/CSV File<BR>
			</span>
			<br>
			<font color="gray" size="1"><b>Description: </b>An extended button control to 
				export Excel or CSV file&nbsp;</font><br>
			<font color="gray" size="1"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Author: </b>
				Prashant Nayak, Copy right © 2004 </font>
		</P>
		<form id="Form1" method="post" runat="server">
			<asp:label id="Label2" runat="server" Font-Size="Medium" Font-Bold="True">Books:</asp:label><BR>
			<asp:datagrid id="dgBooks" runat="server" CellPadding="3" BackColor="White" BorderWidth="1px"
				BorderStyle="None" BorderColor="#CCCCCC">
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#669999"></SelectedItemStyle>
				<ItemStyle ForeColor="#000066"></ItemStyle>
				<HeaderStyle Font-Bold="True" ForeColor="White" BackColor="#006699"></HeaderStyle>
				<FooterStyle ForeColor="#000066" BackColor="White"></FooterStyle>
				<PagerStyle HorizontalAlign="Left" ForeColor="#000066" BackColor="White" Mode="NumericPages"></PagerStyle>
			</asp:datagrid>
			<P><pnwc:exportbutton id="btnCSV" runat="server" Text="Export to CSV"></pnwc:exportbutton><pnwc:exportbutton id="btnExcel" runat="server" Text="Export to Excel"></pnwc:exportbutton><BR>
				<BR>
				<asp:button id="btnLoad" runat="server" Text="Dynamically load Publishers"></asp:button><BR>
				<asp:label id="Label1" runat="server" Font-Size="Medium" Font-Bold="True"> Publishers:</asp:label><BR>
				<asp:listbox id="lbPublishers" runat="server" BackColor="#FFFFC0" Width="249px"></asp:listbox><BR>
				<BR>
				<pnwc:exportlinkbutton id="lnkbtnCSV" runat="server" FileNameToExport="Publishers.csv" ExportType="CSV"
					Separator="Comma">Export to CSV</pnwc:exportlinkbutton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<pnwc:exportlinkbutton id="lnkbtnExcel" runat="server" FileNameToExport="Publishers.xls" ExportType="Excel"
					Separator="TAB">Export to EXCEL</pnwc:exportlinkbutton><BR>
				<BR>
				<asp:button id="btnTestPostback" runat="server" Text="Test Postback"></asp:button></P>
		</form>
	</body>
</HTML>
