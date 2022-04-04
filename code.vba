Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.count > 1 Then Exit Sub
    If Target.Style = "nb" Then
        Cells(12, 1).Value = 0
        Call deassemble
        Cells(4, 4).Value = Int((7 * Rnd) + 1)
        Cells(4, 5).Value = Int((4 * Rnd) + 1)
        Range(Cells(1, 17), Cells(22, 26)).Style = "field"
        If Cells(3, 1).Value = 0 Then
            Cells(3, 1).Value = 1
            Application.OnTime DateAdd("s", 1, Now), "check"
        End If
        Cells(10, 35).Select
    End If
    If Target.Style = "cb" Then
        Cells(3, 1).Value = 1
        Application.OnTime DateAdd("s", 1, Now), "check"
        Cells(10, 35).Select
    End If
    If Target.Style = "sb" Then
        Cells(3, 1).Value = 0
    End If
    If Target.Value = "U" And Cells(3, 1).Value = 1 Then
        Cells(10, 35).Select
        Call rotate_check
    End If
    If Target.Value = "D" And Cells(3, 1).Value = 1 Then
        Cells(10, 35).Select
        Call down_check
    End If
    If Target.Value = "L" And Cells(3, 1).Value = 1 Then
        Cells(10, 35).Select
        Call left_check
    End If
    If Target.Value = "R" And Cells(3, 1).Value = 1 Then
        Cells(10, 35).Select
        Call right_check
    End If
End Sub

Sub deassemble()
    Cells(22, 32).Value = Cells(12, 1).Value Mod 10
    Cells(22, 31).Value = Cells(12, 1).Value \ 10 Mod 10
    Cells(22, 30).Value = Cells(12, 1).Value \ 100 Mod 10
    Cells(22, 29).Value = Cells(12, 1).Value \ 1000 Mod 10
    Cells(22, 28).Value = Cells(12, 1).Value \ 10000 Mod 10
End Sub


Sub down_check()
    
    Dim check As Integer
    Dim field As Integer
    check = 0
    field = 0
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" And (Cells(i + 1, J).Style = "border" Or Cells(i + 1, J).Style = "sf") Then
                check = check + 1
            End If
            
            If Cells(i, J).Style = "ff" Then
                field = field + 1
            End If
        Next i
    Next J
    
    If check = 0 And field <> 0 Then
        Call go_max_down
    End If

End Sub

Sub go_max_down()
    Cells(4, 1).Value = Cells(4, 1).Value + 1
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" Then
                Cells(i, J).Style = "field"
                Cells(i + 1, J).Style = "ff"
            End If
        Next i
    Next J
End Sub


Sub left_check()
    
    Dim check As Integer
    Dim field As Integer
    check = 0
    field = 0
    For J = 17 To 26
        For i = 1 To 22
            If Cells(i, J).Style = "ff" And (Cells(i, J - 1).Style = "border" Or Cells(i, J - 1).Style = "sf") Then
                check = check + 1
            End If
            If Cells(i, J).Style = "ff" Then
                field = field + 1
            End If
        Next i
    Next J
    
    If check = 0 And field > 0 Then
        Call go_left
        Call copy_part
        Call figure_shadow
        Call piece_by_piece
    End If

End Sub

Sub go_left()
    Cells(4, 2).Value = Cells(4, 2).Value - 1
    For J = 17 To 26
        For i = 1 To 22
            If Cells(i, J).Style = "ff" Then
                Cells(i, J).Style = "field"
                Cells(i, J - 1).Style = "ff"
            End If
        Next i
    Next J
End Sub

Sub right_check()
    
    Dim check As Integer
    Dim field As Integer
    check = 0
    field = 0
    For J = 17 To 26
        For i = 1 To 22
            If Cells(i, J).Style = "ff" And (Cells(i, J + 1).Style = "border" Or Cells(i, J + 1).Style = "sf") Then
                check = check + 1
            End If
            If Cells(i, J).Style = "ff" Then
                field = field + 1
            End If
        Next i
    Next J
    
    If check = 0 And field > 0 Then
        Call go_right
        Call copy_part
        Call figure_shadow
        Call piece_by_piece
    End If
End Sub

Sub go_right()
    Cells(4, 2).Value = Cells(4, 2).Value + 1
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" Then
                Cells(i, J).Style = "field"
                Cells(i, J + 1).Style = "ff"
            End If
        Next i
    Next J
    
    
End Sub

Sub rotate_check()

    Dim field As Integer
    field = 0
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" Then
                field = field + 1
            End If
        Next i
    Next J

    If field = 0 Then Exit Sub

    Dim random As Integer
    Dim rotate As Integer
    Dim number As Integer
    Dim rangex1 As Integer
    Dim rangex2 As Integer
    random = Cells(5, 1).Value
    rotate = Cells(5, 2).Value
    If rotate = 4 Then
        Cells(5, 2).Value = 1
    Else
        Cells(5, 2).Value = rotate + 1
    End If
    number = 0
    Select Case rotate
        Case 1
            rangex1 = 57
            rangex2 = 57
        Case 2
            rangex1 = 62
            rangex2 = 62
        Case 3
            rangex1 = 69
            rangex2 = 69
        Case 4
            rangex1 = 50
            rangex2 = 50
    End Select
    For i = 1 To 40
        If Cells(i, rangex1).Style = "border" Then
            number = number + 1
            If number = random Then
                Dim rangey1 As Integer
                Dim rangey2 As Integer
                rangey1 = i + 1
                rangey2 = i + 1
                Do
                    rangex2 = rangex2 + 1
                Loop While Cells(i, rangex2).Style <> "border"
                
                Do
                    rangey2 = rangey2 + 1
                Loop While Cells(rangey2, rangex2).Style <> "border"
                
                
                For Each cel In Range(Cells(Cells(4, 1).Value, Cells(4, 2).Value), Cells(Cells(4, 1).Value + rangey2 - rangey1 - 1, Cells(4, 2).Value - 2 + rangex2 - rangex1))
                    If cel.Style = "border" Or cel.Style = "sf" Then
                        If Cells(5, 2).Value = 1 Then
                           Cells(5, 2).Value = 4
                        Else
                            Cells(5, 2).Value = Cells(5, 2).Value - 1
                        End If
                        If Cells(4, 2) = 16 Then
                            Call right_check
                            Call rotate_check
                            Exit Sub
                        Else
                            Dim move As Integer
                            move = Cells(4, 2).Value
                            Call left_check
                            If move = Cells(4, 2).Value Then Exit Sub
                            Call rotate_check
                            Exit Sub
                            End If
                        End If
                Next cel
                For Jjj = 26 To 17 Step -1
                    For iii = 22 To 1 Step -1
                        If Cells(iii, Jjj).Style = "ff" Then
                            Cells(iii, Jjj).Style = "field"
                        End If
                    Next iii
                Next Jjj
                Call ff_clear
                Dim order As Integer
                order = 0
                For Each cel In Range(Cells(rangey1, rangex1 + 1), Cells(rangey2 - 1, rangex2 - 1))
                    If cel.Style = "ff" Then
                        Cells(Cells(4, 1).Value + cel.Row - rangey1, Cells(4, 2).Value + cel.Column - rangex1 - 1).Style = "ff"
                    End If
                Next cel
                
                
            End If
        End If
    Next i
    interval = 1
    Call copy_part
    Call figure_shadow
    Call piece_by_piece
    Cells(10, 35).Select
End Sub

Sub ff_clear()
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" Then
                Cells(i, J).Style = "field"
            End If
        Next i
    Next J
End Sub

Sub copy_part()
    Dim index As Integer
    index = 0
    For J = Cells(4, 2).Value + 3 To Cells(4, 2).Value Step -1
        For i = Cells(4, 1).Value + 3 To Cells(4, 1).Value Step -1
            If Cells(i, J).Style = "ff" Then
                Cells(7 + index, 1) = i
                Cells(7 + index, 2) = J
                index = index + 1
            End If
        Next i
    Next J
    
End Sub

Sub piece_by_piece()
    Cells(Cells(7, 1).Value, Cells(7, 2).Value).Style = "ff"
    Cells(Cells(8, 1).Value, Cells(8, 2).Value).Style = "ff"
    Cells(Cells(9, 1).Value, Cells(9, 2).Value).Style = "ff"
    Cells(Cells(10, 1).Value, Cells(10, 2).Value).Style = "ff"
End Sub

Sub figure_shadow()
    Dim interval As Integer
    interval = 0
    Do
        interval = interval + 1
        Dim check As Integer
        check = 0
        Dim field As Integer
        field = 0
        For J = Cells(4, 2).Value + 3 To Cells(4, 2).Value Step -1
            For i = 22 To Cells(4, 1).Value Step -1
                If Cells(i, J).Style = "ff" Then
                    field = field + 1
                End If
                If Cells(i, J).Style = "ff" And (Cells(i + interval, J).Style = "border" Or Cells(i + interval, J).Style = "sf") Then
                    check = check + 1
                    Exit For
                End If
            Next i
        Next J
    Loop While check = 0 And field > 0
    
    Call clear_shadow
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" Then
                Cells(i, J).Style = "field"
                Cells(i + interval - 1, J).Style = "pf"
            End If
        Next i
    Next J
    
End Sub

Sub clear_shadow()
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "pf" Then
                Cells(i, J).Style = "field"
            End If
        Next i
    Next J
End Sub


Sub go_down()
    
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" Then
                Cells(i, J).Style = "field"
                Cells(i + 1, J).Style = "ff"
            End If
        Next i
    Next J
    
End Sub

