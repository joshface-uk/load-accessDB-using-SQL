Imports System.Data.OleDb
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connection As OleDbConnection
        Dim command As OleDbCommand
        Dim data_reader As OleDbDataReader
        connection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\Josh\Desktop\Users.accdb")
        connection.Open()
        command = New OleDbCommand("SELECT * FROM Users", connection)
        data_reader = command.ExecuteReader

        If data_reader.HasRows Then
            While data_reader.Read
                Dim NewItem As New ListViewItem
                NewItem.Text = data_reader.GetValue(1) 'Forename
                NewItem.SubItems.Add(data_reader.GetValue(2)) 'Surname
                NewItem.SubItems.Add(data_reader.GetValue(3)) 'Age
                NewItem.SubItems.Add(data_reader.GetValue(4)) 'Gender
                ListView1.Items.Add(NewItem)
            End While
        End If
    End Sub
End Class
