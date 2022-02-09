Imports System.Data.OleDb

Public Class Form11

    Dim na As String = "NA"
    Dim str(100) As String
    Public acc_db = "School.accdb"
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Form1.path = "NA" Then
            PictureBox3.Image = My.Resources.Addphoto
        Else
            PictureBox3.Image = Image.FromFile(Form1.path)
        End If
        DataGridShow(Form1.emp)
        Label5.Text = DataGridView2.Rows(0).Cells(0).Value
        Label4.Text = Form1.Namee
        Label6.Text = DataGridView2.Rows(0).Cells(4).Value
        Label9.Text = DataGridView2.Rows(0).Cells(3).Value
        Label10.Text = DataGridView2.Rows(0).Cells(2).Value
        Panel3.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel4.BackColor = Color.FromArgb(100, Color.LightBlue)
        Panel5.BackColor = Color.FromArgb(100, Color.LightBlue)
        Panel1.BackColor = Color.FromArgb(100, Color.LightYellow)
        Panel2.BackColor = Color.FromArgb(100, Color.LightPink)
        dtgrd()
        Label13.BackColor = Color.Transparent
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
        Label12.BackColor = Color.Transparent
        Label12.Text = (DataGridView1.RowCount - 1).ToString
    End Sub

    Private Sub DataGridShow(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from [Section] WHERE [Employee Number]= '" & str & "';", conn)
        da.Fill(dt)
        DataGridView2.DataSource = dt.DefaultView
        conn.Close()
    End Sub

    Private Sub dtgrd()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("SELECT * FROM Student WHERE [Section] LIKE '" & na & "%'", conn)
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

    Private Sub MASTERLISTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MASTERLISTToolStripMenuItem.Click
        Dim ans As Integer = MessageBox.Show("All unsaved data will be loss, do you wish to proceed ?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form10.Show()
            Close()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        For x As Integer = 0 To DataGridView1.RowCount - 1
            str(x) = DataGridView1.CurrentRow.Cells(x).Value.ToString
        Next
        If str(0) = Nothing Then
            Exit Sub
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If str(0) = Nothing Then
            MessageBox.Show("Click a data first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim ans As Integer = MessageBox.Show("Are you sure you want to add this student?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then

            Dim conn As New OleDbConnection
            Dim sqlquery As New OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =  " + acc_db
            Dim ct As Integer = 0
            While ct <> 2
                sqlquery.Connection = conn
                conn.Open()
                Try
                    If ct = 0 Then
                        sqlquery.CommandText = "UPDATE [Student] SET [Section] = '" & Label5.Text & "' WHERE [Student Number] = '" & str(0) & "'"
                    ElseIf ct = 1 Then
                        sqlquery.CommandText = "UPDATE [Student] SET [Class Adviser] = '" & Form1.Namee & "' WHERE [Student Number] = '" & str(0) & "'"
                    End If
                    sqlquery.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                conn.Close()
                ct = ct + 1
            End While
            MessageBox.Show("Successfully Added", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

            For x As Integer = 0 To 18
                str(x) = Nothing
            Next
            dtgrd()
            Label12.Text = (DataGridView1.RowCount - 1).ToString
        End If

    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click, Panel22.Click, PictureBox2.Click
        Dim ans As Integer = MessageBox.Show("Do you wish to Log Out ?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form1.Show()
            Close()
        End If
    End Sub
End Class