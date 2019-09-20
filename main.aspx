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
        cmd.CommandText = "select * from [userinfo] where [username]='" & username & "' and [rndtext]='" & rndtext & "'"
        dr = cmd.ExecuteReader()
        dr.Read()
        If dr("UserType") = 1 Then
            Button_SystemManagement.Visible = True
        Else
            Button_SystemManagement.Visible = False
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

    Protected Sub Button_ReceivedBox_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        iframe1.Attributes.Item("src") = "box.aspx?id=" + CStr(box.receivedbox)
    End Sub

    Protected Sub Button_SendBox_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        iframe1.Attributes.Item("src") = "box.aspx?id=" + CStr(box.sendbox)
    End Sub
    Protected Sub Button_WriteMail_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        iframe1.Attributes.Item("src") = "writemail.aspx"
    End Sub
    Protected Sub Button_ChangePassword_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error Resume Next
        Dim opass As String
        Dim npass As String
        Dim npass2 As String
        Dim n1 As Integer
        Dim n2 As Integer
        Dim f As Integer
        Dim ni As Integer
        Dim nis As String
        opass = Trim(InputBox("请输入原密码：", "修改密码"))
        If opass = "" Then
            Exit Sub
        End If
        npass = Trim(InputBox("请输入新密码：", "修改密码"))
        If npass = "" Then
            Exit Sub
        End If
        npass2 = Trim(InputBox("请再次输入新密码：", "修改密码"))
        If npass2 = "" Then
            Exit Sub
        End If
        n1 = FunctionClass.GetRnd(1, 99)
        n2 = FunctionClass.GetRnd(1, 99)
        f = FunctionClass.GetRnd(1, 3)
        nis = Trim(InputBox(CStr(n1) & IIf(f = 1, "+", IIf(f = 2, "-", "*")) & CStr(n2) & "=?", "修改密码"))
        If nis Is Nothing Then
            nis = ""
        End If
        If nis = "" Then
            Exit Sub
        End If
        ni = Int(nis)
        If npass <> npass2 Then
            MsgBox("两次输入的密码不一致，请重新输入！", 4096, "错误")
            Exit Sub
        End If
        cmd.CommandText = "select count(*) as count001 from userinfo where username='" & Trim(Request.Cookies("nbmailsystem")("username").ToString()) & "'"
        dr = cmd.ExecuteReader()
        dr.Read()
        If dr("count001") <= 0 Then
            FunctionClass.DisplayErr(weberr.sqlerr)
        Else
            dr.Close()
            cmd.CommandText = "select * from userinfo where username='" & Trim(Request.Cookies("nbmailsystem")("username").ToString()) & "'"
            dr = cmd.ExecuteReader()
            dr.Read()
            If LCase(dr("password")) <> LCase(FunctionClass.MD5String(opass)) Then
                MsgBox("原密码错误！", 4096, "错误")
                dr.Close()
            Else
                dr.Close()
                cmd.CommandText = "update [userinfo] set [password]='" & FunctionClass.MD5String(npass) & "',[rndtext]='" & FunctionClass.Getrndtext(100) & "' where [username]='" & Request.Cookies("nbmailsystem")("username").ToString() & "'"
                MsgBox(cmd.CommandText)
                If cmd.ExecuteNonQuery() > 0 Then
                    MsgBox("密码修改成功！", 4096, "成功")
                    Response.Redirect("index.aspx")
                Else
                    FunctionClass.DisplayErr(weberr.accessdberr)
                End If
            End If
        End If
    End Sub
    Protected Sub Button_logoff_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error Resume Next
        Response.Cookies("nbmailsystem")("username") = ""
        Response.Cookies("nbmailsystem")("password") = ""
        Response.Redirect("index.aspx")
    End Sub
    Protected Sub Button_SystemManagement_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error Resume Next
        cmd.CommandText = "select * from userinfo where username='" & Request.Cookies("nbmailsystem")("username").ToString() & "'"
        dr = cmd.ExecuteReader()
        dr.Read()
        If dr("UserType") <> 1 Then
            MsgBox("本选项只有管理员可以访问！", 4096, "错误")
            dr.Close()
        Else
            dr.Close()
            Response.Redirect("SystemManagement.aspx")
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
            On Error Resume Next
            cmd.CommandText = "select * from WebInfo"
            dr = cmd.ExecuteReader()
            If dr.HasRows() = True Then
                dr.Read()
                Response.Write("<br /><br /><center><img src=logo.jpg><font color=#FFFFFF size=12px>" + dr("comname") + "内部邮件系统欢迎您！</font></center><br /><br />")
            End If
            dr.Close()
        %>
    </div>  
    <table width ="auto" style="width: 744px">
    <tr valign ="top">
    <td style="width: 173px; height: 533px">
    <asp:Button ID="Button_WriteMail" runat="server" OnClick="Button_WriteMail_Click" Text="写邮件" /> 
    <br/>
    <asp:Button ID="Button_ReceivedBox" runat="server" OnClick="Button_ReceivedBox_Click" Text="收件箱" /> 
    <br/>
    <asp:Button ID="Button_SendBox" runat="server" Text="发件箱" OnClick="Button_SendBox_Click" />
    <br/>
    <asp:Button ID="Button_ChangePassword" runat="server" OnClick="Button_ChangePassword_Click" Text="修改密码" /> 
    <br />
    <asp:Button ID="Button_SystemManagement" runat="server" OnClick="Button_SystemManagement_Click" Text="系统管理" Visible="False" /> 
    <br />
    <asp:Button ID="Button_logoff" runat="server" OnClick="Button_logoff_Click" Text="注销" /> 
    </td>
    <td style="height: 533px; width: 1032px;">
    <iframe   id= "iframe1"   runat="server" style="width: 680px; height: 536px"> 
    </iframe>
    </td>
    </tr>
    </table>
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>