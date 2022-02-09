Public Class Form2

    Private Sub Form2_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Panel5.Visible = False    
        Panel8.BackColor = Color.FromArgb(150, Color.White)     
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel10.BackColor = Color.FromArgb(100, Color.Gray)
        Panel21.BackColor = Color.FromArgb(100, Color.CornflowerBlue)
        Panel8.BackColor = Color.FromArgb(150, Color.White)
        Panel19.BackColor = Color.FromArgb(150, Color.White)
        Label18.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Panel5.Visible = False
        Panel3.BackColor = Color.FromKnownColor(KnownColor.Brown)
        Panel2.BackColor = Color.FromArgb(150, Color.LightPink)
        Panel1.BackColor = Color.FromArgb(190, Color.LightYellow)
        PictureBox1.BackColor = Color.Transparent
    End Sub

    Private Sub Panel8_Paint(sender As Object, e As EventArgs) Handles Panel8.MouseClick, Label12.MouseClick, Panel8.MouseDoubleClick, Label12.MouseDoubleClick
        If Panel5.Visible = False And Panel4.Visible = False Then
            Panel8.BackColor = Color.FromKnownColor(KnownColor.Highlight)
            Panel5.Visible = True
        ElseIf Panel5.Visible = True And Panel4.Visible = True Then
            Panel8.BackColor = Color.FromArgb(150, Color.White)
            Panel5.Visible = False
        End If
        Panel4.BackColor = Color.FromArgb(150, Color.White)
        Panel18.BackColor = Color.FromArgb(150, Color.White)
        Panel15.BackColor = Color.FromArgb(150, Color.White)
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click, Panel18.Click
        Hide()
        Form6.Show()
    End Sub

    Private Sub label20_Click(sender As Object, e As EventArgs) Handles label20.Click, Panel15.Click
        Hide()
        Form7.Show()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click, Panel19.Click
        Hide()
        Form9.Show()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Panel4.Click, Label1.Click
        Hide()
        Form3.Show()
    End Sub

    Private Sub Panel18_Paint(sender As Object, e As EventArgs) Handles Panel18.MouseHover, Label6.MouseHover
        Panel4.BackColor = Color.FromArgb(150, Color.White)
        Panel18.BackColor = Color.FromKnownColor(KnownColor.Highlight)
        Panel15.BackColor = Color.FromArgb(150, Color.White)
    End Sub

    Private Sub Panel15_Paint(sender As Object, e As EventArgs) Handles Panel15.MouseHover, label20.MouseHover
        Panel4.BackColor = Color.FromArgb(150, Color.White)
        Panel18.BackColor = Color.FromArgb(150, Color.White)
        Panel15.BackColor = Color.FromKnownColor(KnownColor.Highlight)
    End Sub

    Private Sub Label1_MouseHover(sender As Object, e As EventArgs) Handles Label1.MouseHover, Panel4.MouseHover
        Panel4.BackColor = Color.FromKnownColor(KnownColor.Highlight)
        Panel18.BackColor = Color.FromArgb(150, Color.White)
        Panel15.BackColor = Color.FromArgb(150, Color.White)
    End Sub


    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click, Label17.Click, Panel22.Click
        Dim ans As Integer = MessageBox.Show("Do you want to log out?", "Log out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = DialogResult.Yes Then
            Form1.Show()
            Close()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label13.Text = Date.Now.ToString("dd-MM-yyyy    hh:mm:ss")
    End Sub

  
End Class