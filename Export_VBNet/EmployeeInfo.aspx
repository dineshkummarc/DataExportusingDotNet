<%@ Page Language="vb" AutoEventWireup="false" Codebehind="EmployeeInfo.aspx.vb" Inherits="ExportDemo_VBNet.EmployeeInfo" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>EmployeeInfo - VB.Net</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<P><span style="FONT-WEIGHT: bold; FONT-SIZE: 14pt; FONT-FAMILY: Garamond">Exporting 
				DataGrid into CSV / Excel File
				<br>
				Author : Rama Krishna </span>
		</P>
		<form id="Form1" method="post" runat="server">
			<asp:label id="lblScenario1" runat="server" Font-Bold="True" Font-Size="12pt" Font-Names="Garamond">Scenario # 1</asp:label><BR>
			<span style="FONT-SIZE: 12pt; FONT-FAMILY: Garamond">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* 
				Get all employee details from database and diplay in datagrid.
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* Export all employee details into CSV 
				format. </span>
			<BR>
			<asp:datagrid id="dgEmployee1" runat="server" AutoGenerateColumns="False" BorderColor="#CCCCCC"
				BorderStyle="None" BorderWidth="1px" BackColor="White" CellPadding="3" Height="136px">
				<FooterStyle ForeColor="#000066" BackColor="White"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#669999"></SelectedItemStyle>
				<ItemStyle ForeColor="Black"></ItemStyle>
				<HeaderStyle Font-Names="Garamond" Font-Bold="True" ForeColor="White" BackColor="MidnightBlue"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="EmployeeID" HeaderText="EmployeeID"></asp:BoundColumn>
					<asp:BoundColumn DataField="LastName" HeaderText="LastName"></asp:BoundColumn>
					<asp:BoundColumn DataField="FirstName" HeaderText="FirstName"></asp:BoundColumn>
					<asp:BoundColumn DataField="BirthDate" HeaderText="BirthDate"></asp:BoundColumn>
					<asp:BoundColumn DataField="HireDate" HeaderText="HireDate"></asp:BoundColumn>
					<asp:BoundColumn DataField="Address" HeaderText="Address"></asp:BoundColumn>
					<asp:BoundColumn DataField="PostalCode" HeaderText="PostalCode"></asp:BoundColumn>
				</Columns>
				<PagerStyle HorizontalAlign="Left" ForeColor="#000066" BackColor="White" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:button id="btnExport1" runat="server" Text="Export to CSV"></asp:button><BR>
			<BR>
			<BR>
			<asp:label id="lblScenario2" runat="server" Font-Bold="True" Font-Size="12pt" Font-Names="Garamond">Scenario # 2</asp:label><BR>
			<span style="FONT-SIZE: 12pt; FONT-FAMILY: Garamond">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* 
				Get all employee details from database.
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* Hide EmployeeID and HireDate in datagrid. 
				But they should be used for internal reference.
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* Export all employee details except 
				EmployeeID and HireDate into Excel format. </span>
			<BR>
			<asp:datagrid id="dgEmployee2" runat="server" AutoGenerateColumns="False" BorderColor="#CCCCCC"
				BorderStyle="None" BorderWidth="1px" BackColor="White" CellPadding="3" Height="136px">
				<FooterStyle ForeColor="#000066" BackColor="White"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#669999"></SelectedItemStyle>
				<ItemStyle ForeColor="Black"></ItemStyle>
				<HeaderStyle Font-Names="Garamond" Font-Bold="True" ForeColor="White" BackColor="MidnightBlue"></HeaderStyle>
				<Columns>
					<asp:BoundColumn Visible="False" DataField="EmployeeID" HeaderText="EmployeeID"></asp:BoundColumn>
					<asp:BoundColumn DataField="LastName" HeaderText="LastName"></asp:BoundColumn>
					<asp:BoundColumn DataField="FirstName" HeaderText="FirstName"></asp:BoundColumn>
					<asp:BoundColumn DataField="BirthDate" HeaderText="BirthDate"></asp:BoundColumn>
					<asp:BoundColumn Visible="False" DataField="HireDate" HeaderText="HireDate"></asp:BoundColumn>
					<asp:BoundColumn DataField="Address" HeaderText="Address"></asp:BoundColumn>
					<asp:BoundColumn DataField="PostalCode" HeaderText="PostalCode"></asp:BoundColumn>
					<asp:TemplateColumn HeaderText="Details">
						<ItemTemplate>
							<asp:Label ID="lnkDetails" Runat="server" Font-Underline="True" Font-Name="Garamond" ForeColor="#6600ff"
								ToolTip="Dummy link. Uses EmployeeID & HireDate for further processing.">Click Here</asp:Label>
						</ItemTemplate>
					</asp:TemplateColumn>
				</Columns>
				<PagerStyle HorizontalAlign="Left" ForeColor="#000066" BackColor="White" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:button id="btnExport2" runat="server" Text="Export to Excel"></asp:button><BR>
			<BR>
			<BR>
			<asp:label id="lblScenario3" runat="server" Font-Bold="True" Font-Size="12pt" Font-Names="Garamond">Scenario # 3</asp:label><BR>
			<span style="FONT-SIZE: 12pt; FONT-FAMILY: Garamond">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* 
				Get all employee details from database.
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* Hide EmployeeID and HireDate in datagrid. 
				But they should be used for internal reference.
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* Export all employee details except 
				EmployeeID and HireDate with costmized headers
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;like 'BirthDate' to 'DOB' 
				and 'PostalCode' to 'ZipCode' into CSV format. </span>
			<BR>
			<asp:datagrid id="dgEmployee3" runat="server" AutoGenerateColumns="False" BorderColor="#CCCCCC"
				BorderStyle="None" BorderWidth="1px" BackColor="White" CellPadding="3" Height="136px">
				<FooterStyle ForeColor="#000066" BackColor="White"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#669999"></SelectedItemStyle>
				<ItemStyle ForeColor="Black"></ItemStyle>
				<HeaderStyle Font-Names="Garamond" Font-Bold="True" ForeColor="White" BackColor="MidnightBlue"></HeaderStyle>
				<Columns>
					<asp:BoundColumn Visible="False" DataField="EmployeeID" HeaderText="EmployeeID"></asp:BoundColumn>
					<asp:BoundColumn DataField="LastName" HeaderText="LastName"></asp:BoundColumn>
					<asp:BoundColumn DataField="FirstName" HeaderText="FirstName"></asp:BoundColumn>
					<asp:BoundColumn DataField="BirthDate" HeaderText="BirthDate"></asp:BoundColumn>
					<asp:BoundColumn Visible="False" DataField="HireDate" HeaderText="HireDate"></asp:BoundColumn>
					<asp:BoundColumn DataField="Address" HeaderText="Address"></asp:BoundColumn>
					<asp:BoundColumn DataField="PostalCode" HeaderText="PostalCode"></asp:BoundColumn>
					<asp:TemplateColumn HeaderText="Details">
						<ItemTemplate>
							<asp:Label ID="Label1" Runat="server" Font-Underline="True" Font-Name="Garamond" ForeColor="#6600ff"
								ToolTip="Dummy link. Uses EmployeeID & HireDate for further processing.">Click Here</asp:Label>
						</ItemTemplate>
					</asp:TemplateColumn>
				</Columns>
				<PagerStyle HorizontalAlign="Left" ForeColor="#000066" BackColor="White" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:button id="btnExport3" runat="server" Text="Export to CSV"></asp:button><BR>
			<BR>
			<asp:label id="lblError" style="Z-INDEX: 101; LEFT: 432px; POSITION: absolute; TOP: 16px" runat="server"
				Font-Names="Garamond" ForeColor="Red">Error Message</asp:label><BR>
		</form>
	</body>
</HTML>
