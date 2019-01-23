Imports System
Imports Microsoft.VisualBasic
Imports System.Data.SQLite
Imports System.Text

Public Module Runtime
    Public db As SQLiteConnection = New SQLiteConnection()

    Public Sub OpenSqlite()
        db.ConnectionString = "Data Source=data.db;Version=3;"
        db.Open()	
    End Sub

    Public Sub CloseSqlite()
        db.Close()
    End Sub

    Public Sub SaveData(message As String)
        OpenSqlite()

        Dim sql As StringBuilder = New StringBuilder($"Insert Into 'Messages' VALUES ( ")

        sql.Append("""" + DateTime.Now.ToString() + """ , ")
        sql.Append("""" + message + """")
        sql.Append(")")

        Dim stmt As New SQLiteCommand(sql.ToString(), db)
        stmt.ExecuteNonQuery()

        CloseSqlite()
    End Sub

End Module


Public Class Debug
    Public Shared Sub Print(message As String)
        Console.WriteLine(message)
    End Sub

    Public Shared Sub Print(format As String, ByVal ParamArray args() As Object)
        Dim arg As String = format
        For i As Integer = 0 To UBound(args, 1)
            arg = arg & " " & args(i)
        Next i

        Console.WriteLine(arg)
    End Sub
End Class