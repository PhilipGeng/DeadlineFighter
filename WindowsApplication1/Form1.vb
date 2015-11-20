Public Class Form1
    Public eventArr(100, 4) As String
    Public ddlArr(100, 3) As String
    Public ddlItems As Integer = 0
    Public eventItems As Integer = 0
    Public ddlDisplayItems As Integer = 3

    Public ddlWithin As Integer = 3
    Public todoDisplayItems As Integer = 5
    Public popUpEvery As Integer = 12
    Public popUpWithin As Double = 3

    Public endDate As Date = CDate("2014-4-26 23:59")
    Private score As Integer = 0
    Public isColor As Boolean = True




    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Interval = 10000
        renewDateTime()
        renewDdlArr()
        renewddl7()

        renewCalendar()

    End Sub


    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        notify()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        InitializeMyForm()


        TabControl1.DrawMode = TabDrawMode.OwnerDrawFixed

        renewDateTime()
        renewDdlView()
        renewTodoView()

        renewddl7()
        renewAllDdl()
        renewChart()


        dtpicker_addddl.CustomFormat = "yyyy-MM-dd HH:mm"

        MonthCalendar1.TodayDate = Now

        dtpicker_addevent1.CustomFormat = "yyyy-MM-dd HH:mm"
        dtpicker_addevent2.CustomFormat = "yyyy-MM-dd HH:mm"
        Form_setting.dtpicker_setting.CustomFormat = "yyyy-MM-dd"
        renewScore()

    End Sub



    Private Sub Button_setting_Click(sender As Object, e As EventArgs) Handles Button_setting.Click
        Form_setting.Show()

    End Sub

    Private Sub Button_exit_Click(sender As Object, e As EventArgs) Handles Button_exit.Click
        Me.Close()

    End Sub

    Public Function notify()
        Dim tempStr As String = ""
        Dim tempDif As Integer
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) = "A") Then
                tempDif = CInt(DateDiff("d", Now, ddlArr(i, 0)))
                If ((tempDif >= 0) And (tempDif <= popUpWithin)) Then
                    tempStr += ddlArr(i, 0) & "     " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next
        If (tempStr <> "") Then
            MsgBox(tempStr)
        End If
    End Function



    Public Function renewDdlArr()
        For i = 0 To ddlItems - 1
            If ((ddlArr(i, 2) = "A") And (CDate(ddlArr(i, 0)) < Now)) Then
                ddlArr(i, 2) = "N"
                renewDdlView()
                renewAllDdl()
            End If
        Next
    End Function


    Public Function renewFace()
        Dim temp As Integer = CInt(Label_ddl7.Text)
        If (temp = 0) Then
            PictureBox_face.Image = My.Resources.face0
        ElseIf (temp < 3) Then
            PictureBox_face.Image = My.Resources.face1
        ElseIf (temp < 6) Then
            PictureBox_face.Image = My.Resources.face2
        ElseIf (temp < 10) Then
            PictureBox_face.Image = My.Resources.face3
        ElseIf (temp < 15) Then
            PictureBox_face.Image = My.Resources.face4
        Else
            PictureBox_face.Image = My.Resources.face5
        End If
    End Function
    Public Function renewDateTime()
        Label_time.Text = Format(Now, "HH:mm")
        Label1_date.Text = Format(Now, "MMM d, yyyy dddd")
    End Function

    Public Function renewDdlView()
        TextBox_ddl.Text = ""
        
        Dim tempDif As Integer
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) = "A") Then
                tempDif = DateDiff("d", Now, CDate(ddlArr(i, 0)))
                If ((tempDif >= 0) And (tempDif <= ddlWithin)) Then
                    TextBox_ddl.Text += ddlArr(i, 0) & "      " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next


    End Function

    Public Function renewTodoView()
        TextBox_todo.Text = ""
        Dim temp As Integer = todoDisplayItems
        If todoDisplayItems > eventItems Then
            temp = eventItems
        End If
        For i = 0 To eventItems - 1
            If (eventArr(i, 3) = "A") Then
                TextBox_todo.Text += eventArr(i, 0) & "-" & eventArr(i, 1) & "     " & eventArr(i, 2) & vbCrLf
            End If
            temp -= 1
            If temp <= 0 Then
                Exit For
            End If
        Next
    End Function


    Public Function renewddl7()
        Dim ddl7 As Integer = 0
        Dim tempDiff As Integer = -1
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) = "A") Then
                tempDiff = DateDiff("d", Now, CDate(ddlArr(i, 0)))
                If (Now < CDate(ddlArr(i, 0)) And (tempDiff < 7)) Then
                    ddl7 += 1
                End If
            End If
        Next
        Label_ddl7.Text = ddl7

        renewFace()
    End Function

    Public Function renewAllDdl()
        CheckedListBox_fddl.Items.Clear()
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) = "Y") Then
                CheckedListBox_fddl.Items.Add(ddlArr(i, 0) & "    " & ddlArr(i, 1))
            End If
        Next

        CheckedListBox_deadddl.Items.Clear()
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) = "N") Then
                CheckedListBox_deadddl.Items.Add(ddlArr(i, 0) & "    " & ddlArr(i, 1) & "     DEAD!!!")
            End If
        Next

        CheckedListBox_unddl.Items.Clear()
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) = "A") Then
                CheckedListBox_unddl.Items.Add(ddlArr(i, 0) & "    " & ddlArr(i, 1))
            End If
        Next

    End Function

    Public Function renewChart()

        Dim tempSum As Integer
        Dim tempDif As Integer
        For i = 0 To 13
            tempSum = 0
            For j = 0 To ddlItems - 1
                If (ddlArr(j, 2) <> "D") Then
                    tempDif = DateDiff("d", CDate(ddlArr(j, 0)), endDate)
                    If ((tempDif >= 7 * i) And (tempDif < 7 * (i + 1))) Then
                        tempSum += 1
                    End If
                End If
            Next
            Chart1.Series(0).Points.AddXY(14 - i, tempSum)
        Next
    End Function





    Public Function renewEventDaily()
        CheckedListBox_event.Items.Clear()
        Dim selectedDate As Date = MonthCalendar1.SelectionStart

        For i = 0 To eventItems - 1
            If eventArr(i, 0).Substring(0, 9) = selectedDate.ToString.Substring(0, 9) And (eventArr(i, 3) = "A") Then
                CheckedListBox_event.Items.Add(eventArr(i, 0) & " - " & eventArr(i, 1) & "      " & eventArr(i, 2) & vbCrLf)
            End If

        Next
    End Function

    Public Function renewEventAll()
        CheckedListBox_event.Items.Clear()
        For i = 0 To eventItems - 1
            If (eventArr(i, 3) = "A") Then
                CheckedListBox_event.Items.Add(eventArr(i, 0) & " - " & eventArr(i, 1) & "     " & eventArr(i, 2))
            End If
        Next

    End Function

    Public Function renewCalendar()
        MonthCalendar1.RemoveAllBoldedDates()
        For i = 0 To eventItems - 1
            If (eventArr(i, 3) = "A") Then
                MonthCalendar1.AddBoldedDate(CDate(eventArr(i, 0)))
            End If
        Next
        MonthCalendar1.UpdateBoldedDates()
    End Function


    Public Function renewScore()
        Label_score.Text = score
        If score > 10 Then
            PictureBox_lock1.Image = My.Resources.star
            PictureBox_lock1.Enabled = True
            If score > 20 Then
                PictureBox_lock2.Image = My.Resources.star
                PictureBox_lock2.Enabled = True
                If score > 50 Then
                    PictureBox_lock3.Image = My.Resources.star
                    PictureBox_lock3.Enabled = True
                    If score > 100 Then
                        PictureBox_lock4.Image = My.Resources.star
                        PictureBox_lock4.Enabled = True
                        If score > 200 Then
                            PictureBox_lock5.Image = My.Resources.star
                            PictureBox_lock5.Enabled = True
                        End If
                    End If
                End If
            End If
        End If


    End Function





    Private Sub Button_list_Click(sender As Object, e As EventArgs) Handles Button_list.Click
        Panel_list.Visible = True
        Panel_graph.Visible = False
        Button_graph.Enabled = True
        Button_list.Enabled = False


        
        Button_graph.BackgroundImage = My.Resources.graph
        Button_graph.BackgroundImageLayout = ImageLayout.Zoom

    End Sub

    Private Sub Button_graph_Click(sender As Object, e As EventArgs) Handles Button_graph.Click
        Panel_list.Visible = False
        Panel_graph.Visible = True
        Button_list.Enabled = True
        Button_graph.Enabled = False
        renewChart()

        
        Button_list.BackgroundImage = My.Resources.list
        Button_list.BackgroundImageLayout = ImageLayout.Zoom

    End Sub


    Private Sub Button_addddl_Click(sender As Object, e As EventArgs) Handles Button_addddl.Click
        ddlArr(ddlItems, 0) = Format(dtpicker_addddl.Value, "yyyy-MM-dd HH:mm")
        ddlArr(ddlItems, 1) = TextBox_addddl.Text
        ddlArr(ddlItems, 2) = "A"
        ddlItems += 1

        renewDdlView()
        renewddl7()
        renewAllDdl()
        renewChart()

    End Sub

    Private Sub Button_ddldel_Click(sender As Object, e As EventArgs) Handles Button_ddldel.Click
        If MsgBox("Sure to delete?" & vbNewLine, vbExclamation + vbYesNo, "") = vbYes Then
            Dim arrIndex As Integer = 0
            Dim j As Integer = 0
            For i = 0 To CheckedListBox_unddl.Items.Count - 1
                If (CheckedListBox_unddl.GetItemChecked(i)) Then
                    While (j < ddlItems)

                        If (ddlArr(j, 2) = "A") Then
                            arrIndex += 1
                        End If
                        If (arrIndex = i + 1) Then
                            ddlArr(j, 2) = "D"

                            Exit While
                        End If
                        j += 1
                    End While
                End If
            Next
            renewDdlView()
            renewddl7()
            renewAllDdl()
            renewChart()
        End If
    End Sub

    Private Sub Button_ddlfinish_Click(sender As Object, e As EventArgs) Handles Button_ddlfinish.Click



        Dim arrIndex As Integer = 0
        Dim j As Integer = 0
        For i = 0 To CheckedListBox_unddl.Items.Count - 1
            If (CheckedListBox_unddl.GetItemChecked(i)) Then
                While (j < ddlItems)

                    If (ddlArr(j, 2) = "A") Then
                        arrIndex += 1
                    End If
                    If (arrIndex = i + 1) Then
                        ddlArr(j, 2) = "Y"
                        score += 5
                        Exit While
                    End If
                    j += 1
                End While
            End If
        Next
        renewDdlView()
        renewddl7()
        renewAllDdl()
        renewChart()
        renewScore()

    End Sub

    Private Sub Button_week1_Click(sender As Object, e As EventArgs) Handles Button_week1.Click
        Dim w As Integer = 13
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = DateDiff("d", CDate(ddlArr(i, 0)), endDate)
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next
        MsgBox(tempStr)
    End Sub

    Private Sub Button_week2_Click(sender As Object, e As EventArgs) Handles Button_week2.Click
        Dim w As Integer = 12
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week3_Click(sender As Object, e As EventArgs) Handles Button_week3.Click
        Dim w As Integer = 11
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week4_Click(sender As Object, e As EventArgs) Handles Button_week4.Click
        Dim w As Integer = 10
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week5_Click(sender As Object, e As EventArgs) Handles Button_week5.Click
        Dim w As Integer = 9
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week6_Click(sender As Object, e As EventArgs) Handles Button_week6.Click
        Dim w As Integer = 8
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week7_Click(sender As Object, e As EventArgs) Handles Button_week7.Click
        Dim w As Integer = 7
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week8_Click(sender As Object, e As EventArgs) Handles Button_week8.Click
        Dim w As Integer = 6
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week9_Click(sender As Object, e As EventArgs) Handles Button_week9.Click
        Dim w As Integer = 5
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week10_Click(sender As Object, e As EventArgs) Handles Button_week10.Click
        Dim w As Integer = 4
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week11_Click(sender As Object, e As EventArgs) Handles Button_week11.Click
        Dim w As Integer = 3
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week12_Click(sender As Object, e As EventArgs) Handles Button_week12.Click
        Dim w As Integer = 2
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week13_Click(sender As Object, e As EventArgs) Handles Button_week13.Click
        Dim w As Integer = 1
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub

    Private Sub Button_week14_Click(sender As Object, e As EventArgs) Handles Button_week14.Click
        Dim w As Integer = 0
        Dim tempDif As Integer
        Dim tempStr As String = "Deadlines in Week " & (14 - w) & " :" & vbCrLf
        For i = 0 To ddlItems - 1
            If (ddlArr(i, 2) <> "D") Then
                tempDif = CInt(DateDiff("d", CDate(ddlArr(i, 0)), endDate))
                If ((tempDif >= 7 * w) And (tempDif < 7 * (w + 1))) Then
                    tempStr += ddlArr(i, 0) & "    " & ddlArr(i, 1) & vbCrLf
                End If
            End If
        Next

        MsgBox(tempStr)
    End Sub





    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        Label_event.Text = "Daily Event"
        renewEventDaily()

    End Sub


    Private Sub Button_dailyevent_Click(sender As Object, e As EventArgs) Handles Button_dailyevent.Click
        Label_event.Text = "Daily Event"
        renewEventDaily()
    End Sub



    Private Sub Button_allevent_Click(sender As Object, e As EventArgs) Handles Button_allevent.Click
        Label_event.Text = "All Event"
        renewEventAll()
    End Sub

    Private Sub Button_addevent_Click(sender As Object, e As EventArgs) Handles Button_addevent.Click
        If (dtpicker_addevent1.Value >= dtpicker_addevent2.Value) Then
            MsgBox("Time Period is invalid!")
        Else

            eventArr(eventItems, 0) = dtpicker_addevent1.Value
            eventArr(eventItems, 1) = dtpicker_addevent2.Value
            eventArr(eventItems, 2) = TextBox_addevent.Text
            eventArr(eventItems, 3) = "A"
            eventItems += 1
            If Label_event.Text = "Daily Event" Then
                renewEventDaily()
            Else
                renewEventAll()
            End If
            renewTodoView()
            renewCalendar()

        End If
    End Sub


    Private Sub Button_eventfinish_Click(sender As Object, e As EventArgs) Handles Button_eventfinish.Click
        Dim arrIndex As Integer = 0
        Dim j As Integer = 0
        For i = 0 To CheckedListBox_event.Items.Count - 1
            If (CheckedListBox_event.GetItemChecked(i)) Then
                While (j < eventItems)

                    If ((eventArr(j, 3) = "A") And (checkDayMatch(eventArr(j, 0), MonthCalendar1.SelectionStart.ToString))) Then
                        arrIndex += 1
                    End If
                    If (arrIndex = i + 1) Then
                        eventArr(j, 3) = "D"

                        Exit While
                    End If
                    j += 1
                End While
            End If
        Next

        If Label_event.Text = "Daily Event" Then
            renewEventDaily()
        Else
            renewEventAll()
        End If

        renewTodoView()
        renewCalendar()

    End Sub

    Private Sub Button_eventdel_Click(sender As Object, e As EventArgs) Handles Button_eventdel.Click
        If MsgBox("Sure to delete?" & vbNewLine, vbExclamation + vbYesNo, "") = vbYes Then
            Dim arrIndex As Integer = 0
            Dim j As Integer = 0
            For i = 0 To CheckedListBox_event.Items.Count - 1
                If (CheckedListBox_event.GetItemChecked(i)) Then
                    While (j < eventItems)

                        If ((eventArr(j, 3) = "A") And (checkDayMatch(eventArr(j, 0), MonthCalendar1.SelectionStart.ToString))) Then
                            arrIndex += 1
                        End If
                        If (arrIndex = i + 1) Then
                            eventArr(j, 3) = "D"

                            Exit While
                        End If
                        j += 1
                    End While
                End If
            Next



            If Label_event.Text = "Daily Event" Then
                renewEventDaily()
            Else
                renewEventAll()
            End If


            renewTodoView()
            renewCalendar()
        End If
    End Sub

    Private Function checkDayMatch(date1 As String, date2 As String) As Boolean
        Dim result As Boolean = True
        If ((Label_event.Text = "Daily Event") And (date1.Substring(0, 9) <> date2.Substring(0, 9))) Then
            result = False
        End If
        Return result
    End Function

  
    Private Sub TabControl1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem
        Dim g As Graphics = e.Graphics
        Dim tp As TabPage = TabControl1.TabPages(e.Index)
        Dim br As Brush
        Dim sf As New StringFormat
        Dim r As New RectangleF(e.Bounds.X, e.Bounds.Y + 2, e.Bounds.Width, e.Bounds.Height - 2)

        sf.Alignment = StringAlignment.Center

        Dim strTitle As String = tp.Text
        'If the current index is the Selected Index, change the color
        If TabControl1.SelectedIndex = e.Index Then
            'this is the background color of the tabpage
            'you could make this a stndard color for the selected page
            br = New SolidBrush(tp.BackColor)
            'this is the background color of the tab page
            g.FillRectangle(br, e.Bounds)
            'this is the background color of the tab page
            'you could make this a stndard color for the selected page
            br = New SolidBrush(tp.ForeColor)
            g.DrawString(strTitle, TabControl1.Font, br, r, sf)
        Else
            'these are the standard colors for the unselected tab pages
            br = New SolidBrush(Color.FromArgb(93, 208, 192))
            g.FillRectangle(br, e.Bounds)
            br = New SolidBrush(Color.Gray)
            g.DrawString(strTitle, TabControl1.Font, br, r, sf)
        End If

    End Sub

    Private Sub Button_graph_MouseDown(sender As Object, e As MouseEventArgs) Handles Button_graph.MouseDown
        Button_graph.BackgroundImage = My.Resources.graph_pressed
    End Sub

    
    Private Sub Button_graph_MouseHover(sender As Object, e As EventArgs) Handles Button_graph.MouseHover
        Button_graph.BackgroundImage = My.Resources.graph_over
        Button_graph.BackgroundImageLayout = ImageLayout.Zoom
    End Sub

    Private Sub Button_graph_MouseLeave(sender As Object, e As EventArgs) Handles Button_graph.MouseLeave
        If (Button_graph.Enabled) Then
            Button_graph.BackgroundImage = My.Resources.graph
            Button_graph.BackgroundImageLayout = ImageLayout.Zoom
        End If
    End Sub

    Private Sub Button_list_MouseDown(sender As Object, e As MouseEventArgs) Handles Button_list.MouseDown
        Button_list.BackgroundImage = My.Resources.list_pressed
        Button_list.BackgroundImageLayout = ImageLayout.Zoom
    End Sub
    Private Sub Button_list_MouseHover(sender As Object, e As EventArgs) Handles Button_list.MouseHover
        Button_list.BackgroundImage = My.Resources.list_over
        Button_list.BackgroundImageLayout = ImageLayout.Zoom
    End Sub
    Private Sub Button_list_MouseLeave(sender As Object, e As EventArgs) Handles Button_list.MouseLeave
        If (Button_list.Enabled) Then
            Button_list.BackgroundImage = My.Resources.list
            Button_list.BackgroundImageLayout = ImageLayout.Zoom
        End If
    End Sub
   


    Private Sub Button_addddl_MouseEnter(sender As Object, e As EventArgs) Handles Button_addddl.MouseEnter
        Button_addddl.BackgroundImage = My.Resources.add_over
        Button_addddl.BackgroundImageLayout = ImageLayout.Zoom
    End Sub

    Private Sub Button_addddl_MouseLeave(sender As Object, e As EventArgs) Handles Button_addddl.MouseLeave
        Button_addddl.BackgroundImage = My.Resources.add
        Button_addddl.BackgroundImageLayout = ImageLayout.Zoom
    End Sub

    Private Sub Button_addevent_MouseEnter(sender As Object, e As EventArgs) Handles Button_addevent.MouseEnter
        Button_addevent.BackgroundImage = My.Resources.add_over
        Button_addevent.BackgroundImageLayout = ImageLayout.Zoom
    End Sub

    Private Sub Button_addevent_MouseLeave(sender As Object, e As EventArgs) Handles Button_addevent.MouseLeave
        Button_addevent.BackgroundImage = My.Resources.add
        Button_addevent.BackgroundImageLayout = ImageLayout.Zoom
    End Sub




    Private Sub BorderColor(ByRef _Control As Control, ByVal _Color As Color)

        ' Variables para metodos graficos

        Dim g As Graphics = Panel1.CreateGraphics

        Dim pen As New Pen(_Color, 2.0)
        ' Cambia el borde

        g.DrawRectangle(pen, New Rectangle(_Control.Location, _Control.Size))



        ' Libera los recursos

        pen.Dispose()

        g = Nothing

    End Sub



    Private Sub Panel1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

        ' Cuando se dibuja ( o redibuja ) la pantalla cambia el color de los objetos

        BorderColor(TextBox_ddl, Color.FromArgb(255, 173, 115))
        BorderColor(TextBox_todo, Color.FromArgb(255, 173, 115))

    End Sub


    Public Sub InitializeMyForm()
        Me.Opacity = 0.92
        Me.BackColor = Color.Red

        ' Make the background color of form display transparently.
        TransparencyKey = BackColor

    End Sub 'InitializeMyForm



    Private Sub Button_help_Click(sender As Object, e As EventArgs) Handles Button_help.Click
        Form_help.Show()

        If TabControl1.SelectedIndex = 0 Then
            '      Form_help.PictureBox1.Image = My.Resources.help_t1
            Form_help.PictureBox1.Visible = True
            Form_help.PictureBox7.Visible = True
        ElseIf TabControl1.SelectedIndex = 1 Then
            Form_help.PictureBox2.Visible = True
            If (Panel_list.Visible) Then
                Form_help.PictureBox2.Image = My.Resources.help_t2_2
            Else
                Form_help.PictureBox2.Image = My.Resources.help_t2_4
            End If


            ' Form_help.PictureBox3().Image = My.Resources.help_t2_1
            Form_help.PictureBox3.Visible = True
            Form_help.PictureBox4.Visible = True
        ElseIf TabControl1.SelectedIndex = 2 Then
            Form_help.PictureBox5.Visible = True
            Form_help.PictureBox6.Visible = True
        ElseIf TabControl1.SelectedIndex = 3 Then


        End If

    End Sub



    Private Sub PictureBox_lock1_Click(sender As Object, e As EventArgs) Handles PictureBox_lock1.Click
        If (PictureBox_lock1.Enabled = True) Then
            Form_star.Show()
        End If
    End Sub

    Private Sub PictureBox_lock2_Click(sender As Object, e As EventArgs) Handles PictureBox_lock2.Click
        If (PictureBox_lock1.Enabled = True) Then
            Form_star2.Show()
        End If
    End Sub


End Class


