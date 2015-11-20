Public Class Form_help

    Private Sub Form_help_Click(sender As Object, e As EventArgs) Handles Me.Click
        Me.Close()
    End Sub

    Private Sub Form_help_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(Form1.Location.X, Form1.Location.Y)
        InitializeMyForm()

        
    End Sub

    Public Sub InitializeMyForm()
        Me.BackColor = Color.FromArgb(0, 104, 91)
        ' Make the background color of form display transparently.
        TransparencyKey = BackColor

    End Sub
End Class