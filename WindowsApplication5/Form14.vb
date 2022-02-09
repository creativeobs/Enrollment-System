Imports System.Data.OleDb


Public Class Form14
    Public acc_db = "School.accdb"
    Dim str(100) As String

    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel2.BackColor = Color.FromArgb(150, Color.LightPink)
        Panel1.BackColor = Color.FromArgb(190, Color.LightYellow)
        Label16.BackColor = Color.Transparent
        Panel7.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel5.BackColor = Color.FromArgb(100, Color.LightBlue)
        Label1.BackColor = Color.Transparent
        Label2.BackColor = Color.Transparent
        Label3.BackColor = Color.Transparent
        Label4.BackColor = Color.Transparent
        Label1.Text = Form1.Namee
        Label2.Text = Form1.emp
        Label3.Text = Form1.dept

        If Form1.path = "NA" Then
            PictureBox1.Image = My.Resources.Addphoto
        Else
            PictureBox1.Image = Image.FromFile(Form1.path)
        End If
        dtgrd()
    End Sub

    Private Sub Panel22_Paint(sender As Object, e As EventArgs) Handles PictureBox2.Click, Panel22.Click, Label17.Click
        Dim ans As Integer = MessageBox.Show("Do you want to log out?", "Log out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form1.Show()
            Close()
        End If
    End Sub

    Private Sub dtgrd()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Grades", conn)
        da.Fill(dt)
        DataGridView2.DataSource = dt.DefaultView
        conn.Close()
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        For x As Integer = 0 To DataGridView2.RowCount - 1
            If x Mod 2 Then
                DataGridView2.Rows(x).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView2.Rows(x).DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            End If
        Next
        DataGridView2.Columns("Path").Visible = False
        For Each Column In DataGridView2.Columns
            DataGridView2.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 12)
            Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Column.SortMode = DataGridViewColumnSortMode.NotSortable
            Column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Dim lastMark As Integer = -1
        Dim currentRank As Integer = 0
        Dim atSameRank As Integer = 1
        For x As Integer = 0 To DataGridView2.RowCount - 1

            Dim currentMark As Integer = Convert.ToInt32(DataGridView2.Columns("2nd Sem Final Grade"))

            If currentMark <> lastMark Then

                lastMark = currentMark
                currentRank = currentRank + atSameRank
                atSameRank = 0

            Else
                atSameRank = atSameRank + 1
                DataGridView2.Rows(x).Cells("Rank").Value = currentRank
            End If
        Next


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        For x As Integer = 0 To DataGridView2.ColumnCount - 1
            str(x) = DataGridView2.CurrentRow.Cells(x).Value.ToString
        Next
        If str(0) = Nothing Then
            Exit Sub
        End If
        Label18.Text = str(1)
        Label6.Text = str(0)
        Label19.Text = str(2)
        If str(11) = "NA" Then
            PictureBox3.Image = My.Resources.Addphoto
        Else
            PictureBox3.Image = Image.FromFile(str(11))
        End If
    End Sub

    Private Sub GradeCalculatorToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim ans As Integer = MessageBox.Show("All unsaved data will be lose, proceed ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form13.Show()
            Close()
        End If
    End Sub

    Private Sub AttendanceToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles AttendanceToolStripMenuItem.Click
        Dim ans As Integer = MessageBox.Show("All unsaved data will be lose, proceed ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form15.Show()
            Close()
        End If
    End Sub


End Class