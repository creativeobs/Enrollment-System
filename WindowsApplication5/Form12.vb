Imports System.Data.OleDb
Public Class Form12
    Public acc_db = "School.accdb"
    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.BackColor = Color.Transparent
        Label2.BackColor = Color.Transparent
        Label3.BackColor = Color.Transparent
        Label4.BackColor = Color.Transparent
        Label5.BackColor = Color.Transparent

        Panel1.BackColor = Color.FromArgb(100, Color.LightYellow)
        Panel2.BackColor = Color.FromArgb(100, Color.LightPink)
    End Sub

    Private Sub DataGridShow(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Account WHERE Username='" & str & "'", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox2.Text = Nothing Then
            MessageBox.Show("Enter your username first ", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        TextBox1.Enabled = True

        If TextBox1.Text = DataGridView1.Rows(0).Cells(9).Value Then
            MessageBox.Show("Your Password is " & DataGridView1.Rows(0).Cells(2).Value, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf TextBox1.Text <> DataGridView1.Rows(0).Cells(9).Value Then
            MessageBox.Show("Wrong Security Answer", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        Label5.Text = Nothing
        TextBox1.Enabled = False
        Button1.Enabled = False
        TextBox1.Text = ""
        TextBox2.Text = ""

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridShow(TextBox2.Text)
        If DataGridView1.RowCount > 1 Then
            Label5.Text = DataGridView1.Rows(0).Cells(8).Value
        Else
            MessageBox.Show("NO user found!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Button1.Enabled = True
        TextBox1.Enabled = True
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) 

    End Sub
End Class