Imports System.Data.OleDb

Public Class Form16

    Public path As String
    Dim ofd As New OpenFileDialog

    Public acc_db = "School.accdb"
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If TextBox7.UseSystemPasswordChar = False Then
            TextBox7.UseSystemPasswordChar = True
            TextBox8.UseSystemPasswordChar = True
            PictureBox2.Image = My.Resources.eye2
        Else
            PictureBox2.Image = My.Resources.eye1
            TextBox7.UseSystemPasswordChar = False
            TextBox8.UseSystemPasswordChar = False
        End If
    End Sub

    Private Sub DataGridShow2(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Teacher WHERE [Employee Number]='" & str & "'", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
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

    Private Sub DataGridShow3(str As String)
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("Select * from Account WHERE [Employee Number]='" & str & "'", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        conn.Close()
    End Sub

    Private Sub Form16_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label10.BackColor = Color.Transparent
        Panel4.BackColor = Color.FromArgb(100, Color.LightPink)
        Label5.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label6.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label7.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label8.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label14.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label15.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label16.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label17.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label18.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Label9.BackColor = Color.FromKnownColor(KnownColor.Transparent)
        Panel8.BackColor = Color.FromArgb(100, Color.LightBlue)
        Panel6.BackColor = Color.FromArgb(100, Color.LimeGreen)
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        ofd.Filter = "Choose Image(*.JPG;*.PNG;*.GIF)|*.jpg;*.png;*.gif"
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            PictureBox3.Image = Image.FromFile(ofd.FileName)
            Label14.Text = "Your Photo"
            Label14.Left = (Label14.Parent.Width \ 2) - (Label14.Width \ 2)
        End If
        path = ofd.FileName
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "" Then
            TextBox11.Enabled = False
        Else
            TextBox11.Enabled = True
        End If
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox3.TextLength = 0 Or TextBox4.TextLength = 0 Or TextBox5.TextLength = 0 Or TextBox6.TextLength = 0 Or TextBox7.TextLength = 0 Or TextBox8.TextLength = 0 Or TextBox9.TextLength = 0 Or TextBox10.TextLength = 0 Or TextBox11.TextLength = 0 Or ComboBox1.Text = Nothing Then
            MessageBox.Show("Fill up all the fields", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        DataGridShow2(TextBox9.Text)
        If DataGridView1.RowCount = 1 Then
            MessageBox.Show("You are not authorized and not in the list of the teachers." & vbNewLine & "Please request for the admin first to add you in the list.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub

        ElseIf DataGridView1.RowCount > 1 And DataGridView1.CurrentRow.Cells(0).Value.ToString = TextBox9.Text And DataGridView1.CurrentRow.Cells(5).Value.ToString = "yes" Then
            MessageBox.Show("Employee Number already registered. Only 1 employee number per account is allowed.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        DataGridShow(TextBox6.Text)


        If DataGridView1.RowCount > 1 Then
            MessageBox.Show("The username has been taken", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If TextBox7.TextLength < 8 Then
            Dim ans As Integer = MessageBox.Show(" The Password strength is too weak, do you insist to continue?", "WARNING", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If ans = DialogResult.No Then
                Exit Sub
            End If
        End If

        If TextBox7.Text <> TextBox8.Text Then
            MessageBox.Show("Please Re-enter Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        If path = Nothing Then
            path = "NA"
        End If

        Dim ct = 0
        While ct <> 2
            Dim acc As String = "yes"
            Dim conn As New OleDbConnection
            Dim sqlquery As New OleDbCommand
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =  " + acc_db
            sqlquery.Connection = conn
            conn.Open()
            Try
                If ct = 0 Then
                    sqlquery.CommandText = "Insert into Account ([Employee Number], [Username], [Password], [Last Name], [First Name], [Middle Initial], [Department], [Photo], [Security], [Answer]) VALUES ( '" & TextBox9.Text & "', '" & TextBox6.Text & "', '" & TextBox7.Text & "', '" & TextBox5.Text & "', '" & TextBox4.Text & "', '" & TextBox3.Text & "', '" & TextBox10.Text & "', '" & path & "', '" & ComboBox1.Text & "', '" & TextBox11.Text & "');"
                Else
                    sqlquery.CommandText = "UPDATE Teacher SET [Account] = '" & acc & "' WHERE [Employee Number] = '" & TextBox9.Text & "'"
                End If
                sqlquery.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            ct += 1
        End While
        MessageBox.Show("You successfully added an account.", "account", MessageBoxButtons.OK, MessageBoxIcon.Information)


a:
        TextBox3.Text = Nothing
        TextBox4.Text = Nothing
        TextBox5.Text = Nothing
        TextBox6.Text = Nothing
        TextBox7.Text = Nothing
        TextBox8.Text = Nothing
        TextBox9.Text = Nothing
        TextBox10.Text = Nothing
        ComboBox1.Text = Nothing
        TextBox11.Text = Nothing
        PictureBox3.Image = My.Resources.Addphoto
        Label14.Text = "Add Photo"


    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.TextLength = 0 Then
            Label11.BackColor = Color.White
            Label12.BackColor = Color.White
            Label13.BackColor = Color.White
        ElseIf TextBox7.TextLength < 8 Then
            Label11.BackColor = Color.Red
            Label12.BackColor = Color.White
            Label13.BackColor = Color.White
        ElseIf TextBox7.TextLength < 12 Then
            Label12.BackColor = Color.Yellow
            Label11.BackColor = Color.White
            Label13.BackColor = Color.White
        Else
            Label13.BackColor = Color.Green
            Label11.BackColor = Color.White
            Label12.BackColor = Color.White
        End If
    End Sub

    Private Sub TextBox9_Leave(sender As Object, e As EventArgs) Handles TextBox9.Leave
        DataGridShow2(TextBox9.Text)
        If TextBox9.TextLength > 0 Then
            If DataGridView1.RowCount > 1 Then
                TextBox5.Text = DataGridView1.CurrentRow.Cells(1).Value
                TextBox4.Text = DataGridView1.CurrentRow.Cells(2).Value
                TextBox3.Text = DataGridView1.CurrentRow.Cells(3).Value
                TextBox10.Text = DataGridView1.CurrentRow.Cells(4).Value
            Else
                MessageBox.Show("No employee number registered.", "account", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox9.Text = ""
                Exit Sub
            End If
        Else
        End If

        DataGridShow3(TextBox9.Text)
        If TextBox9.TextLength > 0 Then
            If DataGridView1.RowCount > 1 Then
                MessageBox.Show("Employee already Registered.", "account", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox9.Text = ""
                TextBox5.Text = ""
                TextBox4.Text = ""
                TextBox3.Text = ""
                TextBox10.Text = ""
                Exit Sub
            End If
        Else
        End If

    End Sub

End Class