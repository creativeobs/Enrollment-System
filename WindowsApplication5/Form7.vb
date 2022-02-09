Imports System.Data.OleDb
Public Class Form7

    Dim str(100) As String
    Dim cho As Integer = 0
    Dim ofd As New OpenFileDialog
    Dim path1 As String
    Dim namee As String
    Public acc_db = "School.accdb"
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label16.BackColor = Color.Transparent
        PictureBox2.BackColor = Color.FromArgb(100, Color.LimeGreen)
        PictureBox2.Enabled = False
        Panel9.BackColor = Color.FromArgb(150, Color.Coral)
        Panel7.BackColor = Color.FromArgb(100, Color.Coral)
        Label9.BackColor = Color.Transparent
        TextBox1.Enabled = False
        Panel1.BackColor = Color.FromArgb(190, Color.LightYellow)
        Panel2.BackColor = Color.FromArgb(150, Color.LightPink)
        PictureBox1.BackColor = Color.Transparent
        Label2.BackColor = Color.Transparent
        Label3.BackColor = Color.Transparent
        Label4.BackColor = Color.Transparent
        Label5.BackColor = Color.Transparent
        Label6.BackColor = Color.Transparent
        Label7.BackColor = Color.Transparent
        Label8.BackColor = Color.Transparent
        Label10.BackColor = Color.Transparent
        Label11.BackColor = Color.Transparent
        Label12.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Label14.BackColor = Color.Transparent
        Label17.BackColor = Color.Transparent
        Label18.BackColor = Color.Transparent
        Label1.BackColor = Color.Transparent
        Label15.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Label19.BackColor = Color.Transparent
        Label20.BackColor = Color.Transparent
        Panel6.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel3.BackColor = Color.FromArgb(100, Color.LightBlue)
        Panel5.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel4.BackColor = Color.FromArgb(100, Color.LightBlue)
        Panel3.Visible = False
        dtgrd()
        btnCancel.Enabled = False
        btnSave.Enabled = False
    End Sub

    Private Sub dtgrd()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Student", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
        DataGridView1.Columns("Path").Visible = False
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
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

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        ofd.Filter = "Choose Image(*.JPG;*.PNG;*.GIF)|*.jpg;*.png;*.gif"
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            PictureBox2.Image = Image.FromFile(ofd.FileName)
            If txtLN.Text = Nothing And txtFN.Text = Nothing And txtMN.Text = Nothing Then
                Label16.Text = "No Name"
            Else
                Label16.Text = txtLN.Text & ", " & txtFN.Text & " " & txtMN.Text
                Label16.Left = (Label16.Parent.Width \ 2) - (Label16.Width \ 2)
            End If
        End If
        path1 = ofd.FileName
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        btnDelete.Enabled = False
        btnEdit.Enabled = False
        cho = 1
        btnCancel.Enabled = True
        btnSave.Enabled = True
        PictureBox2.Enabled = True
        Label16.Text = "Add Photo"
        Label16.Left = Label16.Parent.Width \ 2 - Label16.Width \ 2
        While Me.Size.Height < 700
            Me.Size = New Point(1268, Me.Size.Height + 10)
        End While
        Panel3.Visible = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        Dim ans As Integer = MessageBox.Show("All unsaved data will be loss, do you wish to proceed ?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form2.Show()
            Close()
        End If
    End Sub

    Private Sub DataGridShow(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Student WHERE [Student Number]= '" & str & "';", conn)
        da.Fill(dt)
        DataGridView2.DataSource = dt.DefaultView
        conn.Close()
    End Sub

    Private Sub DataGridShow2(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Student WHERE [LRN]= '" & str & "';", conn)
        da.Fill(dt)
        DataGridView2.DataSource = dt.DefaultView
        conn.Close()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If str(0) = Nothing Then
            MessageBox.Show("Click a data first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        DataGridView1.Enabled = False
        btnDelete.Enabled = False
        btnAdd.Enabled = False
        cho = 2
        txtStudno.Text = str(0)
        LRN.Text = str(1)
        txtLN.Text = str(2)
        txtFN.Text = str(3)
        txtMN.Text = str(4)
        MI.Text = str(4).Substring(0, 1)
        AD.Text = str(5)
        GL.Text = str(6)
        f1.Text = str(9)
        f2.Text = str(10)
        f3.Text = str(11)
        m1.Text = str(12)
        m2.Text = str(13)
        m3.Text = str(14)
        g1.Text = str(15)
        g2.Text = str(16)
        g3.Text = str(17)

        btnCancel.Enabled = True
        btnSave.Enabled = True
        PictureBox2.Enabled = True
        namee = str(2) & ", " & str(3) & " " & str(4)
        If str(18) = "NA" Then
            GoTo a
        End If
        Label16.Text = namee
        PictureBox2.Image = Image.FromFile(str(18))
a:
        Label16.Left = Label16.Parent.Width \ 2 - Label16.Width \ 2
        While Me.Size.Height < 700
            Me.Size = New Point(1268, Me.Size.Height + 10)
        End While
        Panel3.Visible = True
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

            txtStudno.Text = Nothing
            LRN.Text = Nothing
            txtLN.Text = Nothing
            txtFN.Text = Nothing
            txtMN.Text = Nothing
            MI.Text = Nothing
            AD.Text = Nothing
            GL.Text = Nothing
            f1.Text = Nothing
            f2.Text = Nothing
            f3.Text = Nothing
            m1.Text = Nothing
            m2.Text = Nothing
            m3.Text = Nothing
            g1.Text = Nothing
            g2.Text = Nothing
            g3.Text = Nothing
            btnAdd.Enabled = True
            btnDelete.Enabled = True
            btnEdit.Enabled = True
            btnCancel.Enabled = False
            DataGridView1.Enabled = True
            btnSave.Enabled = False
            PictureBox2.Enabled = False
            Label16.Text = "Photo"
            PictureBox2.Image = My.Resources.Resources.Addphoto
            Label16.Left = Label16.Parent.Width \ 2 - Label16.Width \ 2
            DataGridView1.Enabled = True
            While Me.Size.Height > 377
                Me.Size = New Point(1268, Me.Size.Height - 10)
            End While
            Panel3.Visible = False
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If str(0) = Nothing Then
            MessageBox.Show("Click a data first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Dim ans As Integer = MessageBox.Show("Are you sure you want to delete Student data?", "Student", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then

            Dim conn As New OleDbConnection
            Dim sqlquery As New OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =  " + acc_db
            sqlquery.Connection = conn
            Dim ct = 0
            While ct <> 2
                conn.Open()
                Try
                    If ct = 0 Then
                        sqlquery.CommandText = "DELETE FROM Student WHERE [Student Number] = '" & str(0) & "';"
                    ElseIf ct = 1 Then
                        sqlquery.CommandText = "DELETE FROM Student WHERE [Student Number] = '" & str(0) & "';"
                    End If

                    sqlquery.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                conn.Close()
                ct = ct + 1
            End While

            MessageBox.Show("Successfully Deleted", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dtgrd()
            For x As Integer = 0 To 18
                str(x) = Nothing
            Next
            PictureBox2.Enabled = False
            Label16.Text = "Photo"
            Label16.Left = Label16.Parent.Width \ 2 - Label16.Width \ 2
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.TextLength >= 1 Then
            Dim conn As New OleDb.OleDbConnection
            Dim sqlquery As New OleDb.OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            da = New OleDbDataAdapter("SELECT * FROM Student WHERE [" & ComboBox1.Text & "] LIKE '" & TextBox1.Text & "%'", conn)
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


    Private Sub txtLN_TextChanged(sender As Object, e As EventArgs) Handles txtLN.TextChanged
        If txtLN.Text = Nothing And txtFN.Text = Nothing And txtMN.Text = Nothing Then
            Label16.Text = "No Name"
            Label16.Left = (Label16.Parent.Width \ 2) - (Label16.Width \ 2)
        Else
            Label16.Text = txtLN.Text & ", " & txtFN.Text & " " & txtMN.Text
            Label16.Left = (Label16.Parent.Width \ 2) - (Label16.Width \ 2)
        End If
    End Sub

    Private Sub txtFN_TextChanged(sender As Object, e As EventArgs) Handles txtFN.TextChanged
        If txtLN.Text = Nothing And txtFN.Text = Nothing And txtMN.Text = Nothing Then
            Label16.Text = "No Name"
            Label16.Left = (Label16.Parent.Width \ 2) - (Label16.Width \ 2)
        Else
            Label16.Text = txtLN.Text & ", " & txtFN.Text & " " & txtMN.Text
            Label16.Left = (Label16.Parent.Width \ 2) - (Label16.Width \ 2)
        End If
    End Sub

    Private Sub txtMN_TextChanged(sender As Object, e As EventArgs) Handles txtMN.TextChanged
        If txtLN.Text = Nothing And txtFN.Text = Nothing And txtMN.Text = Nothing Then
            Label16.Text = "No Name"
            Label16.Left = (Label16.Parent.Width \ 2) - (Label16.Width \ 2)
            MI.Text = Nothing
        Else
            MI.Text = txtMN.Text.Substring(0, 1)
            Label16.Text = txtLN.Text & ", " & txtFN.Text & " " & txtMN.Text
            Label16.Left = (Label16.Parent.Width \ 2) - (Label16.Width \ 2)
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If txtStudno.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        ElseIf txtLN.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        ElseIf txtFN.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        ElseIf txtMN.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        ElseIf LRN.Text = Nothing Or AD.Text = Nothing Or GL.Text = Nothing Or g1.Text = Nothing Or g3.Text = Nothing Or g2.Text = Nothing Or f1.Text = Nothing Or f2.Text = Nothing Or f3.Text = Nothing Or m1.Text = Nothing Or m2.Text = Nothing Or m3.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If path1 = Nothing Then
            path1 = "NA"
        End If

        Dim na As String = "NA"

        If cho = 1 Then

            DataGridShow(txtStudno.Text)

            If DataGridView2.RowCount > 1 Then
                MessageBox.Show("Student is already listed.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            DataGridShow2(LRN.Text)

            If DataGridView2.RowCount > 1 Then
                MessageBox.Show("Student is already listed.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If


            Dim conn As New OleDbConnection
            Dim sqlquery As New OleDbCommand

            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =  " + acc_db
            sqlquery.Connection = conn
            conn.Open()

            Try

                sqlquery.CommandText = "Insert into Student ([Student Number], [LRN], [Last Name], [First Name], [Middle Name], [Address], [Grade Level], [Class Adviser], [Section], [Father], [Father's Occupation], [Father's Contact Number], [Mother], [Mother's Occupation], [Mother's Contact Number], [Guardian], [Guardian's Occupation], [Guardian's Contact Number], [Path]) VALUES ('" & txtStudno.Text & "', '" & LRN.Text & "', '" & txtLN.Text & "', '" & txtFN.Text & "', '" & txtMN.Text & "', '" & AD.Text & "', '" & GL.Text & "', '" & na & "', '" & na & "', '" & f1.Text & "', '" & f2.Text & "', '" & f3.Text & "', '" & m1.Text & "', '" & m2.Text & "', '" & m3.Text & "', '" & g1.Text & "', '" & g2.Text & "', '" & g3.Text & "', '" & path1 & "');"
                sqlquery.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conn.Close()

            MessageBox.Show("Successfully Added", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)


        ElseIf cho = 2 Then

            DataGridShow(txtStudno.Text)

            If DataGridView2.RowCount > 1 And str(0) <> txtStudno.Text Then
                MessageBox.Show("Student is already listed.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            DataGridShow2(LRN.Text)

            If DataGridView2.RowCount > 1 And str(1) <> LRN.Text Then
                MessageBox.Show("Student is already listed.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim ct As Integer = 0
            While ct <> 2

                Dim conn As New OleDbConnection
                Dim sqlquery As New OleDbCommand
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =  " + acc_db
                sqlquery.Connection = conn
                conn.Open()
                Try
                    If ct = 0 Then
                        sqlquery.CommandText = "Delete from Student WHERE [Student Number] = '" & str(0) & "';"
                    ElseIf ct = 1 Then
                        sqlquery.CommandText = "Insert into Student ([Student Number], [LRN], [Last Name], [First Name], [Middle Name], [Address], [Grade Level], [Class Adviser], [Section], [Father], [Father's Occupation], [Father's Contact Number], [Mother], [Mother's Occupation], [Mother's Contact Number], [Guardian], [Guardian's Occupation], [Guardian's Contact Number], [Path]) VALUES ('" & txtStudno.Text & "', '" & LRN.Text & "', '" & txtLN.Text & "', '" & txtFN.Text & "', '" & txtMN.Text & "', '" & AD.Text & "', '" & GL.Text & "', '" & na & "', '" & na & "', '" & f1.Text & "', '" & f2.Text & "', '" & f3.Text & "', '" & m1.Text & "', '" & m2.Text & "', '" & m3.Text & "', '" & g1.Text & "', '" & g2.Text & "', '" & g3.Text & "', '" & path1 & "');"
                    End If
                    sqlquery.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                conn.Close()
                ct = ct + 1
            End While
            For x As Integer = 0 To 18
                str(x) = Nothing
            Next
            MessageBox.Show("Successfully Edited", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

a:
        btnDelete.Enabled = True
        btnAdd.Enabled = True
        btnEdit.Enabled = True
        txtStudno.Text = Nothing
        txtLN.Text = Nothing
        txtFN.Text = Nothing
        txtMN.Text = Nothing
        LRN.Text = Nothing
        AD.Text = Nothing
        GL.Text = Nothing
        g1.Text = Nothing
        g3.Text = Nothing
        txtStudno.Text = Nothing
        LRN.Text = Nothing
        txtLN.Text = Nothing
        txtFN.Text = Nothing
        txtMN.Text = Nothing
        MI.Text = Nothing
        AD.Text = Nothing
        GL.Text = Nothing
        f1.Text = Nothing
        f2.Text = Nothing
        f3.Text = Nothing
        m1.Text = Nothing
        m2.Text = Nothing
        m3.Text = Nothing
        g1.Text = Nothing
        g2.Text = Nothing
        g3.Text = Nothing
        btnCancel.Enabled = False
        DataGridView1.Enabled = True
        btnSave.Enabled = False
        PictureBox2.Image = My.Resources.Resources.Addphoto
        PictureBox2.Enabled = False
        Label16.Text = "Photo"
        Label16.Left = Label16.Parent.Width \ 2 - Label16.Width \ 2
        dtgrd()
        While Me.Size.Height > 377
            Me.Size = New Point(1268, Me.Size.Height - 10)
        End While
        Panel3.Visible = False
    End Sub


End Class