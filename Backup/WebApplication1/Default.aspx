<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApplication1._Default" Debug="true"  %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>成型機產能分析系統</title>
    <style type="text/css">
        #form1
        {
            width: 1096px;
            height: 676px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server" method="post">
    <p style="height: 30px">
        <asp:Label ID="Label3" runat="server" BackColor="#00CCFF" BorderStyle="Dotted" 
            Text="成型機產能分析系統-恒耀工業股份有限公司製作"></asp:Label>
    </p>
    <p>
    <asp:Button ID="Button1" runat="server" Text="系統設定" Visible="False" />
&nbsp;&nbsp;
    <asp:Button ID="Btn_Download" runat="server" Text="檔案下載" 
            onclick="Btn_Download_Click" />
    </p>
    <asp:Panel ID="Panel3" runat="server" 
        
        style="z-index: 1; left: 225px; top: 64px; position: absolute; height: 207px; width: 193px">
        <asp:Label ID="Label1" runat="server" Text="選擇時間"></asp:Label>
        <br />
        <asp:Calendar ID="dateTimeAnalysis" runat="server" BackColor="White" 
        BorderColor="#999999" CellPadding="4" DayNameFormat="Shortest" 
        Font-Names="Verdana" Font-Size="8pt" ForeColor="Black" Height="92px" 
        onselectionchanged="dateTimeAnalysis_SelectionChanged" Width="16px" 
            CaptionAlign="Left">
            <SelectedDayStyle BackColor="#666666" Font-Bold="True" ForeColor="White" />
            <SelectorStyle BackColor="#CCCCCC" />
            <WeekendDayStyle BackColor="#FFFFCC" />
            <TodayDayStyle BackColor="#CCCCCC" ForeColor="Black" />
            <OtherMonthDayStyle ForeColor="#808080" />
            <NextPrevStyle VerticalAlign="Bottom" />
            <DayHeaderStyle BackColor="#CCCCCC" Font-Bold="True" Font-Size="7pt" />
            <TitleStyle BackColor="#999999" BorderColor="Black" Font-Bold="True" />
        </asp:Calendar>
    </asp:Panel>
    <asp:Panel ID="Panel2" runat="server" 
        
        
        
        style="z-index: 1; left: 13px; top: 283px; position: absolute; height: 238px; width: 579px">
        <asp:Label ID="Lb_Analysis" runat="server" Text="產能分析"></asp:Label>
        <asp:GridView ID="GridView_Analysis" runat="server" 
    onrowcommand="GridView_Analysis_RowCommand" BackColor="White" BorderColor="#DEDFDE" 
            BorderStyle="None" BorderWidth="1px" CellPadding="4" ForeColor="Black" 
            GridLines="Vertical" Width="576px" Height="26px">
            <RowStyle BackColor="#F7F7DE" />
            <Columns>
                <asp:ButtonField ButtonType="Button" CausesValidation="True" ShowHeader="True" 
                Text="機台詳細資料" />
            </Columns>
            <FooterStyle BackColor="#CCCC99" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="Panel1" runat="server" 
        
        style="z-index: 1; left: 643px; top: 283px; position: absolute; height: 227px; width: 429px">
        <asp:Label ID="LbDetail" runat="server" Text="機台詳細產能" Visible="False"></asp:Label>
        <asp:GridView ID="GridView_Detail" runat="server" BackColor="White" BorderColor="#DEDFDE" 
            BorderStyle="None" BorderWidth="1px" CellPadding="4" ForeColor="Black" 
            GridLines="Vertical" Height="195px" Width="255px">
            <RowStyle BackColor="#F7F7DE" />
            <FooterStyle BackColor="#CCCC99" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
    </asp:Panel>
    <br />
    <asp:Panel ID="Panel4" runat="server" Height="113px" Width="168px">
        <asp:RadioButtonList ID="RadioBtn_Viewdetails" runat="server" Height="61px" 
        Width="100px" oninit="RadioBtn_Viewdetails_Init" 
    CausesValidation="True" 
    onselectedindexchanged="RadioBtn_Viewdetails_SelectedIndexChanged">
            <asp:ListItem>顯示昨天</asp:ListItem>
            <asp:ListItem>顯示今天</asp:ListItem>
            <asp:ListItem Selected="True">顯示兩天</asp:ListItem>
        </asp:RadioButtonList>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="Btn_UpdateInformation" runat="server" 
            onclick="Btn_UpdateInformation_Click" Text="更新" />
    </asp:Panel>
    </form>
</body>
</html>
