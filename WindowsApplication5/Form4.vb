Imports System.Data.OleDb

Public Class Form4

    Public acc_db = "School.accdb"

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel1.BackColor = Color.FromArgb(100, Color.LightYellow)
        Panel2.BackColor = Color.FromArgb(100, Color.LightPink)
        dtgrd()
        Label5.Text = Form9.str3(0)
        Label4.Text = Form9.str3(1)
        Label6.Text = Form9.str3(4)
        Label9.Text = Form9.str3(3)
        Label10.Text = Form9.str3(2)
        Label1.BackColor = Color.Transparent
        Label2.BackColor = Color.Transparent
        Label3.BackColor = Color.Transparent
        Label4.BackColor = Color.Transparent
        Label5.BackColor = Color.Transparent
        Label6.BackColor = Color.Transparent
        Label7.BackColor = Color.Transparent
        Label8.BackColor = Color.Transparent
        Label9.BackColor = Color.Transparent
        Label10.BackColor = Color.Transparent
        Label11.BackColor = Color.Transparent
        Panel3.BackColor = Color.FromArgb(100, Color.LightBlue)
        Label12.BackColor = Color.Transparent
        Label12.Text = (DataGridView1.RowCount - 1).ToString
    End Sub

    Private Sub dtgrd()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("SELECT * FROM Student WHERE [Section] LIKE '" & Form9.str3(0) & "%'", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
        DataGridView1.Columns("Path").Visible = False
        For x As Integer = 0 To DataGridView1.RowCount - 1
            If x Mod 2 Then
                DataGridView1.Rows(x).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView1.Rows(x).DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            End If
        Next
        For Each Column In DataGridView1.Columns
            Column.width = 250
            DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 12)
            Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Column.SortMode = DataGridViewColumnSortMode.NotSortable
            Column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
    End Sub

    Private Sub DataGridShow(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Account WHERE [Employee Number]= '" & str & "';", conn)
        da.Fill(dt)
        DataGridView2.DataSource = dt.DefaultView
        conn.Close()
    End Sub

End Class