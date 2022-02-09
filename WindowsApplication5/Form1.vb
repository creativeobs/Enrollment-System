Imports System.Data.OleDb
Public Class Form1

    Public Namee As String
    Public emp As String
    Public dept As String
    Public path As String
    Public acc_db = "School.accdb"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "admin" And TextBox2.Text = "school" Then
            Form2.Show()
            TextBox1.Text = Nothing
            TextBox2.Text = Nothing
            Hide()
            Exit Sub
        End If

        If TextBox1.TextLength > 0 And TextBox2.TextLength > 0 Then
            DataGridShow(TextBox1.Text)
            If DataGridView1.RowCount > 1 Then
                If TextBox1.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString And TextBox2.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString Then
                    DataGridShow2(DataGridView1.CurrentRow.Cells(0).Value)
                    If DataGridView1.CurrentRow.Cells(5).Value.ToString = Nothing Then
                        MessageBox.Show("Must assign section first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    Else
                        DataGridShow(TextBox1.Text)
                        TextBox1.Text = Nothing
                        TextBox2.Text = Nothing
                        emp = DataGridView1.CurrentRow.Cells(0).Value.ToString
                        Namee = DataGridView1.CurrentRow.Cells(3).Value.ToString + ", " + DataGridView1.CurrentRow.Cells(4).Value.ToString + " " + DataGridView1.CurrentRow.Cells(5).Value.ToString
                        dept = DataGridView1.CurrentRow.Cells(6).Value.ToString
                        path = DataGridView1.CurrentRow.Cells(7).Value.ToString
                        Form10.Show()
                        Hide()
                    End If

                Else
                    TextBox1.Focus()
                    MessageBox.Show("Invalid Username or Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    TextBox1.Text = Nothing
                    TextBox2.Text = Nothing
                End If

            Else
                TextBox1.Focus()
                MessageBox.Show("No such account is found", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox1.Text = Nothing
                TextBox2.Text = Nothing
            End If

        ElseIf TextBox1.TextLength = 0 Or TextBox2.TextLength = 0 Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox1.Focus()
            TextBox1.Text = Nothing
            TextBox2.Text = Nothing
        End If
        TextBox2.Focus()
        TextBox1.Focus()
    End Sub

    Private Sub DataGridShow2(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Teacher WHERE [Employee Number]='" & str & "'", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LinkLabel1.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        LinkLabel2.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Panel3.BackColor = Color.FromArgb(100, Color.LightYellow)
        Panel2.BackColor = Color.FromArgb(100, Color.Gray)
        Panel4.BackColor = Color.FromArgb(100, Color.LightPink)
        Panel1.BackColor = Color.FromArgb(100, Color.Coral)
        Label1.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label2.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        TextBox1.Focus()
    End Sub

    Private Sub DataGridShow(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Account WHERE Username='" & str & "'", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        If TextBox2.UseSystemPasswordChar = False Then
            TextBox2.UseSystemPasswordChar = True
            PictureBox5.Image = My.Resources.eye2
        Else
            PictureBox5.Image = My.Resources.eye1
            TextBox2.UseSystemPasswordChar = False
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim b As New Form12
        b.ShowDialog(Me)
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim c As New Form16
        c.ShowDialog(Me)
    End Sub

End Class
