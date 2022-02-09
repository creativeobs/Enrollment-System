Imports System.Data.OleDb
Public Class Form5

    Dim str(3) As String
    Dim cho As Integer = 0
    Public acc_db = "School.accdb"

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel7.BackColor = Color.FromArgb(100, Color.Coral)
        Label7.BackColor = Color.Transparent
        TextBox1.Enabled = False
        LineShape1.BorderColor = Color.Gray
        LineShape1.SelectionColor = Color.Transparent
        Panel1.BackColor = Color.FromArgb(190, Color.LightYellow)
        Panel2.BackColor = Color.FromArgb(150, Color.LightPink)
        PictureBox1.BackColor = Color.Transparent
        Label1.BackColor = Color.Transparent
        Label2.BackColor = Color.Transparent
        Label3.BackColor = Color.Transparent
        Label4.BackColor = Color.Transparent
        Label5.BackColor = Color.Transparent
        Label6.BackColor = Color.Transparent
        Panel6.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel3.BackColor = Color.FromArgb(100, Color.LightBlue)
        Panel5.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel4.BackColor = Color.FromArgb(100, Color.LightBlue)
        dtgrd()
        btnCancel.Enabled = False
        btnSave.Enabled = False
    End Sub

    Private Sub dtgrd()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Rooms", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        For x As Integer = 0 To DataGridView1.RowCount - 1
            If x Mod 2 Then
                DataGridView1.Rows(x).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView1.Rows(x).DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            End If
        Next
        For Each Column In DataGridView1.Columns
            DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 12)
            Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Column.SortMode = DataGridViewColumnSortMode.NotSortable
            Column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        btnDelete.Enabled = False
        btnEdit.Enabled = False
        txtRoomno.Enabled = True
        txtFloor.Enabled = True
        txtDesc.Enabled = True
        cho = 1
        txtRoomno.Focus()
        btnCancel.Enabled = True
        btnSave.Enabled = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        Dim ans As Integer = MessageBox.Show("All unsaved data will be loss, do you wish to proceed ?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form2.Show()
            txtRoomno.Text = Nothing
            txtDesc.Text = Nothing
            txtFloor.Text = Nothing
            txtRoomno.Enabled = False
            txtFloor.Enabled = False
            txtDesc.Enabled = False
            btnAdd.Enabled = True
            btnDelete.Enabled = True
            btnEdit.Enabled = True
            btnCancel.Enabled = False
            btnSave.Enabled = False
            DataGridView1.Enabled = True
            Close()
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If txtRoomno.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        ElseIf txtFloor.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        ElseIf txtDesc.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
      
        If cho = 1 Then
            DataGridShow(txtRoomno.Text)
            If DataGridView2.RowCount > 1 Then
                MessageBox.Show("Room Number is already listed.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim ct As Integer = 0
            Dim conn As New OleDbConnection
            Dim sqlquery As New OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =" + acc_db
            sqlquery.Connection = conn
            conn.Open()
            Try
                sqlquery.CommandText = "Insert into Rooms ([Room Number], [Description], [Floor]) VALUES ('" & txtRoomno.Text & "', '" & txtDesc.Text & "', '" & txtFloor.Text & "');"
                sqlquery.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conn.Close()

            MessageBox.Show("Successfully Added", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ElseIf cho = 2 Then

            Dim ct As Integer = 0
            While ct <> 2

                Dim conn As New OleDbConnection
                Dim sqlquery As New OleDbCommand
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =" + acc_db
                sqlquery.Connection = conn
                conn.Open()
                Try
                    If ct = 0 Then
                        sqlquery.CommandText = "Delete from Rooms WHERE [Room Number] = '" & str(0) & "';"
                    ElseIf ct = 1 Then
                        sqlquery.CommandText = "Insert into Rooms ([Room Number], [Description], [Floor]) VALUES ('" & txtRoomno.Text & "', '" & txtDesc.Text & "', '" & txtFloor.Text & "');"
                    End If
                    sqlquery.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                conn.Close()
                ct = ct + 1
            End While
            MessageBox.Show("Successfully Edited", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For x As Integer = 0 To 3
                str(x) = Nothing
            Next
        End If

a:
        btnDelete.Enabled = True
        btnAdd.Enabled = True
        DataGridView1.Enabled = True
        btnEdit.Enabled = True
        txtRoomno.Text = Nothing
        txtDesc.Text = Nothing
        txtFloor.Text = Nothing
        txtRoomno.Enabled = False
        txtFloor.Enabled = False
        txtDesc.Enabled = False
        btnCancel.Enabled = False
        btnSave.Enabled = False
        dtgrd()

    End Sub

    Private Sub DataGridShow(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Rooms WHERE [Room Number]= '" & str & "';", conn)
        da.Fill(dt)
        DataGridView2.DataSource = dt.DefaultView
        conn.Close()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If str(0) = Nothing Then
            MessageBox.Show("Click a data first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        btnDelete.Enabled = False
        btnAdd.Enabled = False
        DataGridView1.Enabled = False
            txtRoomno.Enabled = True
            txtFloor.Enabled = True
            txtDesc.Enabled = True
            txtRoomno.Text = str(0)
            txtDesc.Text = str(1)
            txtFloor.Text = str(2)
        txtRoomno.Focus()
        cho = 2
        btnCancel.Enabled = True
        btnSave.Enabled = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        For x As Integer = 0 To DataGridView1.ColumnCount - 1
            str(x) = DataGridView1.CurrentRow.Cells(x).Value.ToString
        Next
        If str(0) = Nothing Then
            Exit Sub
        End If
    End Sub

    Private Sub btnCancel_Click_1(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim ans As Integer = MessageBox.Show("Cancel?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            txtRoomno.Text = Nothing
            txtDesc.Text = Nothing
            txtFloor.Text = Nothing
            txtRoomno.Enabled = False
            txtFloor.Enabled = False
            txtDesc.Enabled = False
            btnAdd.Enabled = True
            btnDelete.Enabled = True
            btnEdit.Enabled = True
            btnCancel.Enabled = False
            btnSave.Enabled = False
            DataGridView1.Enabled = True
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If str(0) = Nothing Then
            MessageBox.Show("Click a data first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Dim ans As Integer = MessageBox.Show("Are you sure you want to delete this room?", "Room", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
                Dim conn As New OleDbConnection
                Dim sqlquery As New OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =" + acc_db
            sqlquery.Connection = conn
                conn.Open()
                Try
                sqlquery.CommandText = "DELETE FROM Rooms WHERE [Room Number] = '" & str(0) & "';"
                sqlquery.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conn.Close()       
            For x As Integer = 0 To 3
                str(x) = Nothing
            Next
            MessageBox.Show("Successfully Deleted", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dtgrd()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.TextLength >= 1 Then
            Dim conn As New OleDb.OleDbConnection
            Dim sqlquery As New OleDb.OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            da = New OleDbDataAdapter("SELECT * FROM Rooms WHERE [" & ComboBox1.Text & "] LIKE '" & TextBox1.Text & "%'", conn)
            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            conn.Close()
            For x As Integer = 0 To DataGridView1.RowCount - 1
                If x Mod 2 Then
                    DataGridView1.Rows(x).DefaultCellStyle.BackColor = Color.White
                Else
                    DataGridView1.Rows(x).DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
                End If
            Next
        Else
            dtgrd()
        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        If ComboBox1.Text = "" Then
            TextBox1.Enabled = False
        Else
            TextBox1.Enabled = True
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class