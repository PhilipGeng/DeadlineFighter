Public Class Form_setting

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
      

        Form1.ddlWithin = CInt(ComboBox_set1.Text)
        Form1.Label_ddlview.Text = "(within " & ComboBox_set1.Text & " days)"
        Form1.todoDisplayItems = CInt(ComboBox_set2.Text)
        Form1.Label_todoview.Text = "(first " & ComboBox_set2.Text & " items)"

        If (ComboBox_set3.Text = "Never") Then
            Form1.popUpEvery = -1
            Form1.Timer2.Enabled = False
        Else
            Form1.popUpEvery = CInt(ComboBox_set3.Text)
            Form1.Timer2.Enabled = True
        End If
        If Form1.popUpEvery > 0 Then
            Form1.Timer2.Interval = Form1.popUpEvery * 60 * 60 * 1000
        End If



        If (ComboBox_set4.Text = "Never") Then
            Form1.Timer2.Enabled = False
            Form1.popUpWithin = -1
        Else
            Form1.popUpWithin = CInt(ComboBox_set4.Text)
            Form1.Timer2.Enabled = True
        End If

        Form1.endDate = dtpicker_setting.Value

        If (RadioButton_image.Checked) Then
            Form1.TabPage1.BackgroundImage = My.Resources.IMAG0237
            Form1.isColor = False

        Else
            Form1.TabPage1.BackgroundImage = Nothing
            Form1.isColor = True
        End If


        Form1.renewDdlView()
        Form1.renewTodoView()

        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub Form_setting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox_set1.Text = Form1.ddlWithin
        ComboBox_set2.Text = Form1.todoDisplayItems
        If Form1.popUpEvery > 0 Then
            ComboBox_set3.Text = Form1.popUpEvery
        Else
            ComboBox_set3.Text = "Never"
        End If

        If Form1.popUpWithin > 0 Then
            ComboBox_set4.Text = Form1.popUpWithin
        Else
            ComboBox_set4.Text = "Never"
        End If


        If (Form1.isColor) Then
            RadioButton_color.Checked = True
        Else
            RadioButton_image.Checked = True
        End If
    End Sub




End Class