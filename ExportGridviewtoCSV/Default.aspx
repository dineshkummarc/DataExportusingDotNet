<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title></title>
</head>
<body>
<form id="form1" runat="server">
<div>
<table>
<tr>
<td align="right">
<asp:ImageButton ID="btnCSV" runat="server" ImageUrl="~/CSVImage.jpg" 
onclick="btnCSV_Click" />
<asp:ImageButton ID="btnWord" runat="server" ImageUrl="~/PDF.jpg" Width="32px" Height="32px"/>
</td>
</tr>
<tr>
<td>
<asp:GridView runat="server" ID="gvdetails" DataSourceID="dsdetails"  AllowPaging="true" AllowSorting="true" AutoGenerateColumns="false">
<RowStyle BackColor="#EFF3FB" />
<FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
<PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
<HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
<AlternatingRowStyle BackColor="White" />
<Columns>
<asp:BoundField DataField="UserId" HeaderText="UserId" />
<asp:BoundField DataField="UserName" HeaderText="UserName" />
<asp:BoundField DataField="LastName" HeaderText="LastName" />
<asp:BoundField DataField="Location" HeaderText="Location" />
</Columns>
</asp:GridView>
</td>
</tr>
</table>
<asp:SqlDataSource ID="dsdetails" runat="server" ConnectionString="<%$ConnectionStrings:dbconnection %>"
SelectCommand="select * from UserInformation"/>
</div>
</form>
</body>
</html>