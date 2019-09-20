Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Imports system.Web
Public Class dbconn
    Dim conn As New OleDbConnection
    Dim IsErr As Boolean
    Dim ErrText As String
    Dim FunctionClass As New FunctionClass
    Public Sub Create()
        conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=; Data Source=" & HttpContext.Current.Server.MapPath("nbmailsystem.aspx")
        IsErr = False
        ErrText = ""
        Try
            conn.Open()
        Catch ex As Exception
            IsErr = True
            ErrText = ex.Message
            FunctionClass.DisplayErr(weberr.dbconnerr)
        End Try
    End Sub
    Public Function GetConn() As OleDbConnection
        Return conn
    End Function
    Public Function GetIsErr() As Boolean
        Return IsErr
    End Function
    Public Function GetErrText() As String
        Return ErrText
    End Function
    Public Sub Close()
        conn.Close()
        conn = Nothing
        IsErr = False
        ErrText = ""
    End Sub
End Class

