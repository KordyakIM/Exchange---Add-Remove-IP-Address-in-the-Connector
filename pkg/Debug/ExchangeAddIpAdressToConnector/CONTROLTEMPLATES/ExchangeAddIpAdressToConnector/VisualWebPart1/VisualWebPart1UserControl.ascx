<%@ Assembly Name="ExchangeAddIpAdressToConnector, Version=1.0.0.0, Culture=neutral, PublicKeyToken=18e91c879bf703ed" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="VisualWebPart1UserControl.ascx.cs" Inherits="ExchangeAddIpAdressToConnector.VisualWebPart1.VisualWebPart1UserControl" %>
<asp:Table id="Table1" runat="server">
    <asp:TableRow>
        <asp:TableCell>
            <asp:ListBox ID="ListBoxMailboxServers" runat="server" Height="219px" Width="240px" >
                <asp:ListItem>snrzfex01</asp:ListItem>
                <asp:ListItem>snrzfex02</asp:ListItem>
                <asp:ListItem>snrzfex03</asp:ListItem>
                <asp:ListItem>snrzfex04</asp:ListItem>
                <asp:ListItem>snrzfex05</asp:ListItem>
                <asp:ListItem>sntzfex01</asp:ListItem>
                <asp:ListItem>sntzfex02</asp:ListItem>
                <asp:ListItem>sntzfex03</asp:ListItem>
                <asp:ListItem>sntzfex04</asp:ListItem>
                <asp:ListItem>sntzfex05</asp:ListItem>
            </asp:ListBox>
        </asp:TableCell>
        <asp:TableCell>
            <asp:ListBox ID="ListBoxConnectors" runat="server" Height="219px" Width="316px"></asp:ListBox>
        </asp:TableCell>
        <asp:TableCell>
            <asp:ListBox ID="ListBoxIPAddress" runat="server" Height="219px" Width="316px"></asp:ListBox>
        </asp:TableCell>
    </asp:TableRow>
    <asp:TableRow>
        <asp:TableCell>
            <asp:Button OnClick="button_ReceiveConnectors" ID="button1" runat="server" Text="Download receive connectors"></asp:Button>
        </asp:TableCell>
        <asp:TableCell>
            <asp:Button OnClick="button_ReceiveIPAddress" ID="button4" runat="server" Text="Download IP Address"></asp:Button>
        </asp:TableCell>
        <asp:TableCell>
            <br />
        </asp:TableCell>
    </asp:TableRow>
    <asp:TableRow>
        <asp:TableCell>
            <br />
        </asp:TableCell>
    </asp:TableRow>
</asp:Table>
<asp:Table id="Table2" runat="server">
    <asp:TableRow>
        <asp:TableCell >
            <asp:Label ID="Label_UserName" runat="server" Text="UserName:"></asp:Label>
            <asp:Textbox ID="TextBox_UserName" runat="server" Text=""></asp:Textbox>
        </asp:TableCell>
    </asp:TableRow>
    <asp:TableRow>
        <asp:TableCell >
            <asp:Label ID="Password" runat="server" Text="Password:"></asp:Label>
            <asp:Textbox ID="TextBox_Password" runat="server" Text="" TextMode="Password"></asp:Textbox>
        </asp:TableCell>
    </asp:TableRow>
</asp:Table>
<asp:Table id="Table3" runat="server">
    <asp:TableRow>
        <asp:TableCell >
            <asp:Label ID="Label_IPRange" runat="server" Text="IP Range:"></asp:Label>
            <asp:TextBox ID="TextBox_IPRange" runat="server" Width="316px"></asp:TextBox>
        </asp:TableCell>
    </asp:TableRow>
    <asp:TableRow>
        <asp:TableCell >
            <asp:Button OnClick="button_AddIP" ID="button2" runat="server" Text="Add IP Address"></asp:Button>
            <asp:Button OnClick="button_RemoveIP" ID="button3" runat="server" Text="Remove IP Address"></asp:Button>
        </asp:TableCell>
    </asp:TableRow>
</asp:Table>
<br />
<asp:Label ID="Label2" runat="server" Text="Result:"></asp:Label>
<br />
<asp:Textbox ID="ResultBox" TextMode="MultiLine" runat="server" Height="250px" Width="800px" Text=""></asp:Textbox>
