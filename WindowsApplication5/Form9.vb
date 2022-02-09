Imports System.Data.OleDb


Public Class Form9

    Public str(100) As String
    Public str2(100) As String
    Public str3(100) As String
    Dim cho As Integer = 0
    Public acc_db = "School.accdb"

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label9.BackColor = Color.Transparent
        txtdepttttt.BackColor = Color.Transparent
        Panel1.BackColor = Color.FromArgb(190, Color.LightYellow)
        Panel8.BackColor = Color.FromArgb(150, Color.LightPink)
        Label1.BackColor = Color.Transparent
        Label2.BackColor = Color.Transparent
        Label3.BackColor = Color.Transparent
        Label4.BackColor = Color.Transparent
        Label5.BackColor = Color.Transparent
        Label6.BackColor = Color.Transparent
        Label8.BackColor = Color.Transparent
        txtdepttttt.BackColor = Color.Transparent
        Label11.BackColor = Color.Transparent     
        Label10.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Label14.BackColor = Color.Transparent
        Label7.BackColor = Color.Transparent
        Label18.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        en.BackColor = Color.Transparent
        ln.BackColor = Color.Transparent
        fn.BackColor = Color.Transparent
        mn.BackColor = Color.Transparent
        d.BackColor = Color.Transparent
        sn.BackColor = Color.Transparent
        da.BackColor = Color.Transparent
        t.BackColor = Color.Transparent
        rn.BackColor = Color.Transparent
        Label15.BackColor = Color.Transparent
        Panel7.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel2.BackColor = Color.FromArgb(100, Color.LightBlue)
        Label12.BackColor = Color.Transparent
        Panel6.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel3.BackColor = Color.FromArgb(100, Color.LightBlue)
        Panel5.BackColor = Color.FromArgb(100, Color.LimeGreen)
        Panel4.BackColor = Color.FromArgb(100, Color.LightBlue)
        dtgrd1()
        dtgrd2()
        dtgrd3()
    End Sub

    Private Sub dtgrd1()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from [Teacher]", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
        DataGridView1.Columns("Account").Visible = False
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
        For x As Integer = 0 To DataGridView2.RowCount - 1
            If x Mod 2 Then
                DataGridView2.Rows(x).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView2.Rows(x).DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            End If
        Next
    End Sub

    Private Sub dtgrd2()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from [Section]", conn)
        da.Fill(dt)
        DataGridView2.DataSource = dt.DefaultView
        conn.Close()
        DataGridView2.Columns("Employee Number").Visible = False
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        For Each Column In DataGridView2.Columns
            Column.width = 250
            DataGridView2.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 12)
            Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Column.SortMode = DataGridViewColumnSortMode.NotSortable
            Column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        For x As Integer = 0 To DataGridView2.RowCount - 1
            If x Mod 2 Then
                DataGridView2.Rows(x).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView2.Rows(x).DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            End If
        Next
    End Sub

    Private Sub dtgrd3()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from [Section]", conn)
        da.Fill(dt)
        DataGridView3.DataSource = dt.DefaultView
        conn.Close()
        DataGridView3.Columns("Employee Number").Visible = False
        DataGridView3.Columns("Day").Visible = False
        DataGridView3.Columns("Time").Visible = False
        DataGridView3.Columns("Room Number").Visible = False
        DataGridView3.Sort(DataGridView3.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        For Each Column In DataGridView3.Columns
            Column.width = 250
            DataGridView3.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 12)
            Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Column.SortMode = DataGridViewColumnSortMode.NotSortable
            Column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next

        For x As Integer = 0 To DataGridView3.RowCount - 1
            If x Mod 2 Then
                DataGridView3.Rows(x).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView3.Rows(x).DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            End If
        Next
    End Sub


    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        Dim ans As Integer = MessageBox.Show("All unsaved data will be loss, do you wish to proceed ?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form2.Show()
            Close()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        For x As Integer = 0 To DataGridView1.ColumnCount - 1
            str(x) = DataGridView1.CurrentRow.Cells(x).Value.ToString
        Next
        If str(0) = Nothing Then
            Exit Sub
        End If
        en.Text = str(0)
        ln.Text = str(1)
        fn.Text = str(2)
        mn.Text = str(3)
        d.Text = str(4)
        cho = 1
    End Sub

    Private Sub DataGridView1_CellContent3Click(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        For x As Integer = 0 To DataGridView2.ColumnCount - 1
            str2(x) = DataGridView2.CurrentRow.Cells(x).Value.ToString
        Next
        If str2(0) = Nothing Then
            Exit Sub
        End If
        sn.Text = str2(0)
        da.Text = str2(2)
        t.Text = str2(3)
        rn.Text = str2(4)
        cho = 2
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        For x As Integer = 0 To DataGridView3.ColumnCount - 1
            str3(x) = DataGridView3.CurrentRow.Cells(x).Value.ToString
        Next
        If str3(0) = Nothing Then
            Exit Sub
        End If
        cho = 3
    End Sub


    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        If str(0) = Nothing And str2(0) = Nothing And str3(0) = Nothing Then
            MessageBox.Show("Click a data first", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim ans As Integer = MessageBox.Show("Are you sure you want to delete this data?", "data", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then

            Dim conn As New OleDbConnection
            Dim sqlquery As New OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =  " + acc_db
            sqlquery.Connection = conn
            conn.Open()
            Try
                If cho = 1 Then
                    sqlquery.CommandText = "DELETE FROM Teacher WHERE [Employee Number] = '" & str(0) & "';"
                ElseIf cho = 2 Then
                    sqlquery.CommandText = "DELETE FROM [Section] WHERE [Section] = '" & str2(0) & "';"
                ElseIf cho = 3 Then
                    sqlquery.CommandText = "DELETE FROM [Section] WHERE [Section] = '" & str3(0) & "';"
                End If
                sqlquery.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conn.Close()

            MessageBox.Show("Successfully Deleted", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dtgrd1()
            dtgrd2()
            dtgrd3()
            For x As Integer = 0 To 3
                str(x) = Nothing
                str2(x) = Nothing
                str3(x) = Nothing
            Next
            dtgrd1()
            dtgrd2()
            dtgrd3()
            en.Text = Nothing
            ln.Text = Nothing
            fn.Text = Nothing
            mn.Text = Nothing
            d.Text = Nothing
            sn.Text = Nothing
            da.Text = Nothing
            t.Text = Nothing
            rn.Text = Nothing
        End If

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click

        If str(0) = Nothing Or str2(0) = Nothing Then
            MessageBox.Show("Click a data first for the teacher and the Section", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If cho = 3 Then
            MessageBox.Show("Please Select only from the teacher and section", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim namee As String = str(1) & ", " & str(2) & " " & str(3)

        For x As Integer = 0 To DataGridView2.RowCount - 1
            If namee = DataGridView2.Rows(x).Cells(1).Value Then
                MessageBox.Show("The teacher already have a section", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If str2(1) <> "NA" Then
                MessageBox.Show("There is already a teacher in this section", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next

        Dim ans As Integer = MessageBox.Show("Are you sure you want to add this Section?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then

            Dim conn As New OleDbConnection
            Dim sqlquery As New OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =  " + acc_db

            Dim ct As Integer = 0
            While ct <> 3
                sqlquery.Connection = conn
                conn.Open()
                Try
                    If ct = 0 Then
                        sqlquery.CommandText = "UPDATE [Section] SET [Class Adviser] = '" & namee & "' WHERE [Section] = '" & str2(0) & "'"
                    ElseIf ct = 1 Then
                        sqlquery.CommandText = "UPDATE [Section] SET [Employee Number] = '" & str(0) & "' WHERE [Section] = '" & str2(0) & "'"
                    ElseIf ct = 2 Then
                        sqlquery.CommandText = "UPDATE [Teacher] SET [Section] = '" & str2(0) & "' WHERE [Employee Number] = '" & str(0) & "'"
                    End If
                    sqlquery.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                conn.Close()

                ct = ct + 1
            End While
            MessageBox.Show("Successfully Added", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dtgrd1()
            dtgrd2()
            dtgrd3()

            For x As Integer = 0 To 3
                str(x) = Nothing
                str2(x) = Nothing
                str3(x) = Nothing

            Next

            dtgrd1()
            dtgrd2()
            dtgrd3()
            en.Text = Nothing
            ln.Text = Nothing
            fn.Text = Nothing
            mn.Text = Nothing
            d.Text = Nothing
            sn.Text = Nothing
            da.Text = Nothing
            t.Text = Nothing
            rn.Text = Nothing
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
            da = New OleDbDataAdapter("SELECT * FROM [Section] WHERE [" & ComboBox1.Text & "] LIKE '" & TextBox1.Text & "%'", conn)
            da.Fill(dt)
            DataGridView3.DataSource = dt.DefaultView
            conn.Close()
            DataGridView3.Columns("Employee Number").Visible = False
            For x As Integer = 0 To DataGridView3.RowCount - 1
                If x Mod 2 Then
                    DataGridView3.Rows(x).DefaultCellStyle.BackColor = Color.White
                Else
                    DataGridView3.Rows(x).DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
                End If
            Next
        Else
            dtgrd1()
            dtgrd2()
            dtgrd3()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        If ComboBox1.Text = "" Then
            TextBox1.Enabled = False
        Else
            TextBox1.Enabled = True
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If str3(0) = Nothing Then
            MessageBox.Show("Click a data first in the Summary", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim b As New Form4
        b.ShowDialog(Me)
    End Sub

End Class