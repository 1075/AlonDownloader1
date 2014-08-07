Public Class Form3
    Dim int As Integer
    Private Sub CheckBox1_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles CheckBox1.Paint
        If CheckBox1.Checked Then
            e.Graphics.DrawImageUnscaled((My.Resources.checked), 0, 0)
        Else
            e.Graphics.DrawImageUnscaled((My.Resources.unchecked), 0, 0)
        End If
    End Sub
    Private Sub CheckBox2_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles CheckBox2.Paint
        If CheckBox2.Checked Then
            e.Graphics.DrawImageUnscaled((My.Resources.checked), 0, 0)
        Else
            e.Graphics.DrawImageUnscaled((My.Resources.unchecked), 0, 0)
        End If
    End Sub
    Private Sub CheckBox3_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles CheckBox3.Paint
        If CheckBox3.Checked Then
            e.Graphics.DrawImageUnscaled((My.Resources.checked), 0, 0)
        Else
            e.Graphics.DrawImageUnscaled((My.Resources.unchecked), 0, 0)
        End If
    End Sub

    Public Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        int += 1
        If int = 1 Then
            Label2.Text = "Please wait."
        ElseIf int = 2 Then
            Label2.Text = "Please wait.."
        Else
            Label2.Text = "Please wait..."
            int = 0
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If CheckBox1.Checked = False AndAlso CheckBox2.Checked = False AndAlso CheckBox3.Checked = False Or CheckBox3.Visible = False AndAlso CheckBox1.Checked = False AndAlso CheckBox2.Checked = False Then
            MsgBox("Are you sure you wouldn't like to get 'Perceptions' this week?" & vbNewLine & "You still haven't chosen an article!", MsgBoxStyle.YesNo, "Are you sure?")
            If MsgBoxResult.Yes Then
                Me.Close()
            Else

            End If
        Else
            Label2.Visible = True
            PictureBox2.Visible = True
            If CheckBox1.Checked Then
               
            End If
            If CheckBox2.Checked Then
                
            End If
            If CheckBox3.Checked Then

            End If
        End If
    End Sub
End Class