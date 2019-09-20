Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Imports system.Web
Imports System.Net.Mail
Public Enum weberr
    dbconnerr
    sqlerr
    parerr
    accessdberr
    regfalse
End Enum
Public Enum logrecord
    logins
    logine
    loginpe
    dbconne
    sqle
End Enum
Public Enum box
    sendbox
    receivedbox
    draftbox
End Enum
Public Class FunctionClass
    Public Sub DisplayErr(ByVal ErrNumber As weberr)
        HttpContext.Current.Response.Redirect("error.aspx?id=" + CStr(ErrNumber))
    End Sub
    Public Function GetUrl() As String
        Dim url As String, host_url As String, no_http As String  '定义变量
        Dim webdir As String
        url = HttpContext.Current.Request.Url.ToString '获取当前页的URL
        no_http = url.Substring(url.IndexOf("//") + 2) '截取去掉HTTP://的URL
        host_url = "http://" & no_http.Substring(0, no_http.IndexOf("/")) + webdir '组合成当前网站的域名 
        Return host_url
    End Function
    Public Sub sendmail(ByVal email As String, ByVal subject As String, ByVal mailbody As String)
        '创建发件连接,根据你的发送邮箱的SMTP设置填充
        Dim smtp As New System.Net.Mail.SmtpClient("smtp.126.com", 25)
        '发件邮箱身份验证,参数分别为 发件邮箱登录名和密码
        smtp.Credentials = New System.Net.NetworkCredential("nbmailsystem", "lzrnbmailsystem")
        '创建邮件
        Dim mail As New System.Net.Mail.MailMessage()
        '邮件主题
        mail.Subject = subject
        '主题编码
        mail.SubjectEncoding = System.Text.Encoding.GetEncoding("GB2312")
        '邮件正文件编码
        mail.BodyEncoding = System.Text.Encoding.GetEncoding("GB2312")
        '发件人邮箱
        mail.From = New System.Net.Mail.MailAddress("nbmailsystem@126.com")
        '邮件优先级
        mail.Priority = System.Net.Mail.MailPriority.Normal
        'HTML格式的邮件,为false则发送纯文本邮箱
        mail.IsBodyHtml = True
        '邮件内容
        mail.Body = mailbody
        '添加收件人,如果有多个,可以多次添加
        mail.To.Add(email)
        '定义附件,参数为附件文件名,包含路径,推荐使用绝对路径
        '如果不需要附件,下面三行可以不要
        '发送邮件
        Try
            smtp.Send(mail)
        Catch
            MsgBox("发送失败")
        Finally
            mail.Dispose()
        End Try
    End Sub
    Public Function checkemail(ByVal email As String) As Boolean
        Dim a As Integer
        Dim b As Boolean
        Dim c As Integer
        b = False
        c = 0
        email = LCase(Trim(email))
        If email = "" Then
            Return False
        ElseIf Mid(email, Len(email), 1) = "@" Or Mid(email, Len(email), 1) = "." Then
            Return False
        ElseIf (Asc(Mid(email, 1, 1)) > 122 Or Asc(Mid(email, 1, 1)) < 97) And IsNumeric(Mid(email, 1, 1)) = False Then
            Return False
        End If
        For a = 1 To Len(email)
            If (Asc(Mid(email, a, 1)) > 122 Or Asc(Mid(email, a, 1)) < 97) And IsNumeric(Mid(email, a, 1)) = False And Mid(email, a, 1) <> "_" And Mid(email, a, 1) <> "-" And Mid(email, a, 1) <> "@" And Mid(email, a, 1) <> "." Then
                Return False
            ElseIf Mid(email, a, 1) = "@" Then
                If b = True Then
                    Return False
                ElseIf Mid(email, a + 1, 1) = "." Then
                    Return False
                Else
                    b = True
                End If
            ElseIf Mid(email, a, 1) = "." Then
                c = c + 1
                If c > 3 Then
                    Return False
                ElseIf Mid(email, a + 1, 1) = "@" Then
                    Return False
                ElseIf b = False Then
                    Return False
                End If
            End If
        Next
        Return True
    End Function
    Public Function MD5String(ByVal InputString As String) As String
        Dim MD5
        Dim dataToHash As Byte() = (New System.Text.ASCIIEncoding).GetBytes(InputString)
        Dim hashvalue As Byte() = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(dataToHash)
        Dim i As Integer
        MD5 = ""
        For i = 0 To 15
            MD5 += Right("00" + Hex(hashvalue(i)), 2)
        Next
        Return MD5
    End Function
    Public Function Getrndtext(ByVal len As Long) As String
        Dim a As Integer
        Dim i As Integer
        Dim rndtext As String
        rndtext = ""
        For i = 1 To len
            a = GetRnd(0, 100) Mod 2
            If a = 0 Then
                a = GetRnd(0, 9)
                rndtext = rndtext + CStr(a)
            Else
                a = GetRnd(97, 122)
                rndtext = rndtext + Chr(a)
            End If
        Next
        Return rndtext
    End Function
    Public Function GetRnd(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Randomize()
        Return Int(Rnd() * (Max - Min + 1) + Min)
    End Function
    Public Function IsTrue(ByVal Text As String) As Boolean
        Dim i As Integer
        Text = Trim(Text)
        If Len(Text) = 0 Then
            Return True
        Else
            For i = 1 To Len(Text)
                If Mid(Text, i, 1) = "" Then
                    Return False
                End If
            Next
            Return True
        End If
    End Function
    Public Function GetIP() As String
        Dim ip As String
        ip = Trim(HttpContext.Current.Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
        If ip = "" Then
            ip = Trim(HttpContext.Current.Request.ServerVariables("REMOTE_ADDR"))
        End If
	If ip Is Nothing Then
	    ip=""
	End If
        Return ip
    End Function
    Public Sub CheckSql()
        Dim JK1986_Sql As String
        Dim JK_Sql As String()
        Dim k As String
        Dim jk As Integer
        JK1986_Sql = "exec↓select↓drop↓alter↓exists↓union↓and↓or↓xor↓order↓mid↓asc↓execute↓xp_cmdshell↓insert↓update↓delete↓join↓declare↓char↓sp_oacreate↓wscript.shell↓xp_regwrite↓'↓;↓--↓/↓*"
        JK_Sql = JK1986_Sql.Split("↓")
        For Each k In JK_Sql
            '-----------------------防 GET 注入-----------------------
            If System.Web.HttpContext.Current.Request.QueryString.ToString() <> "" Then
                For jk = 0 To System.Web.HttpContext.Current.Request.QueryString.Count - 1
                    If IsTrue(System.Web.HttpContext.Current.Request.QueryString(System.Web.HttpContext.Current.Request.QueryString.Keys(jk).ToString())) = False Then
                        DisplayErr(weberr.sqlerr)
                    End If
                    If jk > System.Web.HttpContext.Current.Request.QueryString.Count Then Exit For
                Next
            End If
            '-----------------------防 Post 注入-----------------------
            If System.Web.HttpContext.Current.Request.Form.ToString() <> "" Then
                For jk = 0 To System.Web.HttpContext.Current.Request.Form.Count - 1
                    If IsTrue(System.Web.HttpContext.Current.Request.Form(System.Web.HttpContext.Current.Request.Form.Keys(jk).ToString())) = False Then
                        DisplayErr(weberr.sqlerr)
                    End If
                    If jk > System.Web.HttpContext.Current.Request.Form.Count Then Exit For
                Next
            End If
            '-----------------------防 Cookies 注入-----------------------
            If System.Web.HttpContext.Current.Request.Cookies.ToString() <> "" Then
                If System.Web.HttpContext.Current.Request.Cookies("nbmailsystem")("username").ToString() <> Nothing Then
                    If IsTrue(System.Web.HttpContext.Current.Request.Cookies("nbmailsystem")("username").ToString()) = False Then
                        DisplayErr(weberr.sqlerr)
                    End If
                ElseIf System.Web.HttpContext.Current.Request.Cookies("nbmailsystem")("rndtext").ToString() <> Nothing Then
                    If IsTrue(System.Web.HttpContext.Current.Request.Cookies("nbmailsystem")("rndtext").ToString()) = False Then
                        DisplayErr(weberr.sqlerr)
                    End If
                End If
            End If
        Next
    End Sub
End Class
