Module Module1


    Public Sub restrict(e As KeyPressEventArgs, str As TextBox)
        If Not Char.IsLetter(CChar(e.KeyChar)) AndAlso Not {" "c, ControlChars.Back}.Contains(CChar(e.KeyChar)) Then
            e.Handled = True
        End If

        If e.KeyChar = " " And str.TextLength = 0 Then
            e.Handled = True
        End If

        If e.KeyChar = " " And str.TextLength <> 0 Then
            If str.Text.Substring(str.TextLength - 1, 1) = " " Then
                e.Handled = True
            End If
        End If
    End Sub

    Public Sub restrictnum(e As KeyPressEventArgs)
        If Not Char.IsDigit(CChar(e.KeyChar)) AndAlso Not {ControlChars.Back}.Contains(CChar(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub


    Public Sub restrictspace(e As KeyPressEventArgs, str As TextBox)
        If e.KeyChar = " " And str.TextLength = 0 Then
            e.Handled = True
        End If

        If e.KeyChar = " " And str.TextLength <> 0 Then
            If str.Text.Substring(str.TextLength - 1, 1) = " " Then
                e.Handled = True
            End If
        End If
    End Sub

    Public Sub restrictletnum(e As KeyPressEventArgs, str As TextBox)
        If Not Char.IsLetterOrDigit(CChar(e.KeyChar)) AndAlso Not {" "c, ControlChars.Back}.Contains(CChar(e.KeyChar)) Then
            e.Handled = True
        End If

        If e.KeyChar = " " And str.TextLength = 0 Then
            e.Handled = True
        End If

        If e.KeyChar = " " And str.TextLength <> 0 Then
            If str.Text.Substring(str.TextLength - 1, 1) = " " Then
                e.Handled = True
            End If
        End If
    End Sub

    Public Sub min(a As TextBox, b As Integer)
        If a.TextLength < b And a.TextLength <> 0 Then
            MsgBox("Minimum required characters is not satisfied.", MsgBoxStyle.Exclamation)
            a.Clear()
            a.Focus()
        Else
        End If
    End Sub

End Module
