Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        f2.Left = f1.Left + 650
        f2.Top = f1.Top + 120
    End Sub
    Private Sub Form2_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        open_w = False
        f1.g.Clear(Color.Transparent)
        PictureBox1.Refresh()
    End Sub
End Class