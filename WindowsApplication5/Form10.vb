Imports System.Data.OleDb
Imports Excl = Microsoft.Office.Interop.Excel.Constants
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form10

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

    Private Sub dtgrd()
        Dim conn As New OleDb.OleDbConnection
        Dim sqlquery As New OleDb.OleDbCommand
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + acc_db
        Dim ds As New DataSet
        Dim dt As New DataTable
        ds.Tables.Add(dt)
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("SELECT * FROM Student WHERE [Section] LIKE '" & Label5.Text & "%'", conn)
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

    Private Sub ADDSTUDENTSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDSTUDENTSToolStripMenuItem.Click
        Dim ans As Integer = MessageBox.Show("All unsaved data will be loss, do you wish to proceed ?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form11.Show()
            Close()
        End If
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click, Panel22.Click, PictureBox2.Click
        Dim ans As Integer = MessageBox.Show("Do you wish to Log Out ?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form1.Show()
            Close()
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ExportDgvToExcel(DataGridView1, Form1.Namee, Label5.Text)
    End Sub

    Public Sub ExportDgvToExcel(ByVal DGV As DataGridView, title As String, sortt As String)
        If DGV.Columns.Count = 1 Then
            Exit Sub
        End If

        Dim dset As New DataSet
        dset.Tables.Add()
        For i As Integer = 0 To DGV.ColumnCount - 1
            dset.Tables(0).Columns.Add(DGV.Columns(i).HeaderText)
        Next
        Dim dr1 As DataRow
        For i As Integer = 0 To DGV.RowCount - 1
            dr1 = dset.Tables(0).NewRow
            For j As Integer = 0 To DGV.Columns.Count - 1
                dr1(j) = DGV.Rows(i).Cells(j).Value
            Next
            dset.Tables(0).Rows.Add(dr1)
        Next

        Dim excel As New Excel.Application
        Dim wBook As Excel.Workbook
        Dim wSheet As Excel.Worksheet

        wBook = excel.Workbooks.Add()
        wSheet = CType(wBook.ActiveSheet(), Microsoft.Office.Interop.Excel.Worksheet)

        Dim dt As System.Data.DataTable = dset.Tables(0)
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 1
        Dim colnm As String = ""
        Dim counter As Integer = 0


        Dim filepath = My.Computer.FileSystem.CurrentDirectory

        wSheet.Shapes.AddPicture(filepath + "\icon.JPG",
             Microsoft.Office.Core.MsoTriState.msoFalse,
             Microsoft.Office.Core.MsoTriState.msoCTrue, 3, 5, 72, 72)
        excel.Range("A1:I1").Merge()
        excel.Range("A1:I1").Font.Bold = True
        excel.Range("A2:I2").Merge()
        excel.Range("A2:I2").Font.Bold = True
        excel.Range("A3:I3").Merge()
        excel.Range("A3:I3").Font.Bold = True
        excel.Range("A4:I4").Merge()
        excel.Range("A4:I4").Font.Bold = True
        excel.Range("A5:I5").Merge()
        excel.Range("A5:I5").Font.Bold = True
        excel.Range("A6:I6").Merge()
        excel.Range("A6:I6").Font.Bold = True
        excel.Range("A7:I7").Merge()
        excel.Range("A7:I7").Font.Bold = True
        excel.Range("A6:I6").Font.Size = 15
        excel.Range("A8:I8").Font.Bold = True
        excel.Range("A1:Z999").HorizontalAlignment = Excl.xlCenter
        excel.Range("A1:Z999").NumberFormat = "0"
        excel.Cells(1, 1) = "Date: " & MonthName(Date.Today.Month) & " " & Date.Today.Day & ", " & Date.Today.Year
        excel.Cells(2, 1) = "La Immaculada Conception School"
        excel.Cells(3, 1) = "PASIG CITY"
        excel.Cells(4, 1) = "Section: " & sortt
        excel.Cells(5, 1) = "Class Adviser: " & title
        excel.Cells(6, 1) = "MASTER LIST"

        For Each dc In dt.Columns
            colIndex = colIndex + 1
            If colIndex = 1 Or colIndex = 2 Or colIndex = 3 Or colIndex = 4 Or colIndex = 5 Or colIndex = 7 Or colIndex = 8 Or colIndex = 9 Then
                If colIndex = 1 Then
                    colnm = "Student Number"
                ElseIf colIndex = 2 Then
                    colnm = "Learner's Reference Number"
                ElseIf colIndex = 3 Then
                    colnm = "Last Name"
                ElseIf colIndex = 4 Then
                    colnm = "First Name"
                ElseIf colIndex = 5 Then
                    colnm = "Middle Name"
                ElseIf colIndex = 7 Then
                    colnm = "Grade Level"
                ElseIf colIndex = 8 Then
                    colnm = "Class Adviser"
                ElseIf colIndex = 9 Then
                    colnm = "Section"
                End If
                counter += 1
                excel.Cells(8, counter) = colnm
                wSheet.Cells(8, counter).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                wSheet.Cells(8, counter).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                wSheet.Cells(8, counter).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                wSheet.Cells(8, counter).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            End If
        Next

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            counter = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                If colIndex = 1 Or colIndex = 2 Or colIndex = 3 Or colIndex = 4 Or colIndex = 5 Or colIndex = 7 Or colIndex = 8 Or colIndex = 9 Then
                    counter += 1
                    excel.Cells(rowIndex + 7, counter) = dr(dc.ColumnName)
                    wSheet.Cells(rowIndex + 7, counter).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    wSheet.Cells(rowIndex + 7, counter).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    wSheet.Cells(rowIndex + 7, counter).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    wSheet.Cells(rowIndex + 7, counter).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                End If
            Next
        Next

        wSheet.Columns.AutoFit()
        Dim blnFileOpen As Boolean = False
        Try
            Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(filepath + "\Print_DATA.xlsx")
            fileTemp.Close()
        Catch ex As Exception
            blnFileOpen = False
        End Try

        If System.IO.File.Exists(filepath + "\Print_DATA.xlsx") Then
            System.IO.File.Delete(filepath + "\Print_DATA.xlsx")
        End If
        wBook.SaveAs(filepath + "\Print_DATA")
        excel.Visible = True

        excel.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized

        excel.Visible = True
        Dim ps As Excel.PageSetup
        ps = excel.ActiveSheet.PageSetup

        ps.Zoom = False
        ps.FitToPagesWide = 1

        ps.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait
        excel.ActiveWorkbook.PrintPreview()
        excel.Quit()
    End Sub

End Class