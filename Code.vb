Imports System.Data.OleDb

Public Class LoginActivity

    Private Sub LoginActivity_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadAllLogonActivity()
    End Sub

    Private Sub LoadAllLogonActivity()
        listview_AllLogonActivity.Items.Clear()
        Dim connection As OleDbConnection
        Dim command As OleDbCommand
        Dim data_reader As OleDbDataReader
        connection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=S:\Data\Application\.data\.secure\LogonActivityLog.accdb")
        connection.Open()
        command = New OleDbCommand("SELECT * FROM LoginActivityLog", connection)
        data_reader = command.ExecuteReader

        If data_reader.HasRows Then
            While data_reader.Read
                Dim NewItem As New ListViewItem()
                NewItem.Text = data_reader.GetValue(1) 'Date
                NewItem.SubItems.Add(data_reader.GetValue(2)) 'Time
                NewItem.SubItems.Add(data_reader.GetValue(3)) 'Username
                NewItem.SubItems.Add(data_reader.GetValue(4)) 'Success/Fail
                NewItem.SubItems.Add(data_reader.GetValue(5)) 'Message
                listview_AllLogonActivity.Items.Add(NewItem)
            End While
        End If
    End Sub

    Private Sub UpdateData_Tick(sender As Object, e As EventArgs) Handles UpdateData.Tick
        LoadAllLogonActivity()
    End Sub
End Class