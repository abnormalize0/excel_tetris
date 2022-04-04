Dim lines As Integer
Sub check()
    If Cells(3, 1).Value = 0 Then Exit Sub
    Dim check As Integer
    Dim field As Integer
    check = 0
    field = 0
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" Then
                field = field + 1
            End If
            
            If Cells(i, J).Style = "ff" And (Cells(i + 1, J).Style = "border" Or Cells(i + 1, J).Style = "sf") Then
                check = check + 1
            End If
        
        Next i
    Next J
    
    If check = 0 And field <> 0 Then
        Call go_down
        Cells(4, 1).Value = Cells(4, 1).Value + 1
    ElseIf field <> 0 Then
        
        Call stand
    End If
    
    If field = 0 Then
        lines = 0
        Call erase_lines
        Call score
        Call spawn

        Call copy_part
        Call figure_shadow
        Call piece_by_piece
        
    End If
    
    If Cells(3, 1).Value = 1 Then
        Application.OnTime DateAdd("s", 1, Now), "check"
    End If
End Sub

Sub score()
    Select Case lines
        Case 1
            Cells(12, 1).Value = Cells(12, 1).Value + 10 * 0.5
        Case 2
            Cells(12, 1).Value = Cells(12, 1).Value + 20 * 1
        Case 3
            Cells(12, 1).Value = Cells(12, 1).Value + 30 * 1.5
        Case 4
            Cells(12, 1).Value = Cells(12, 1).Value + 40 * 2
    End Select
    Call deassemble
End Sub

Sub deassemble()
    Cells(22, 32).Value = Cells(12, 1).Value Mod 10
    Cells(22, 31).Value = Cells(12, 1).Value \ 10 Mod 10
    Cells(22, 30).Value = Cells(12, 1).Value \ 100 Mod 10
    Cells(22, 29).Value = Cells(12, 1).Value \ 1000 Mod 10
    Cells(22, 28).Value = Cells(12, 1).Value \ 10000 Mod 10
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

Sub stand()
    
    For J = 26 To 17 Step -1
        For i = 22 To 1 Step -1
            If Cells(i, J).Style = "ff" Then
                Cells(i, J).Style = "sf"
            End If
        Next i
    Next J
    
End Sub

Sub erase_lines()
    For i = 22 To 1 Step -1
        Dim filed As Integer
        filed = 0
        For J = 26 To 17 Step -1
            If Cells(i, J).Style = "sf" Then
                filed = filed + 1
            End If
        Next J
        If filed = 10 Then
            Range(Cells(1, 17), Cells(i - 1, 26)).copy
            Range(Cells(2, 17), Cells(i, 26)).PasteSpecial
            For Jjj = 26 To 17 Step -1
                Cells(1, Jjj).Style = "field"
            Next Jjj
            Cells(10, 35).Select
            lines = lines + 1
            Call erase_lines
        End If
    Next i
End Sub

Sub color_change()
    For J = 1 To 5
        For i = 40 To 45
            If Cells(J, i).Style = "ff" Then
                Cells(J, i).Style = "pf"
            End If
        Next i
    Next J
End Sub

Sub spawn()
    Dim random As Integer
    Dim rotate As Integer
    Dim nextrandom As Integer
    Dim nextrotate As Integer
    Dim number As Integer
    Dim rangex1 As Integer
    Dim rangex2 As Integer
    random = Cells(4, 4).Value
    rotate = Cells(4, 5).Value
    nextrandom = Int((7 * Rnd) + 1)
    nextrotate = Int((4 * Rnd) + 1)
    Cells(4, 4).Value = nextrandom
    Cells(4, 5).Value = nextrotate
    Cells(5, 1).Value = random
    Cells(5, 2).Value = rotate
    number = 0
    Select Case rotate
        Case 1
            rangex1 = 50
            rangex2 = 50
        Case 2
            rangex1 = 57
            rangex2 = 57
        Case 3
            rangex1 = 62
            rangex2 = 62
        Case 4
            rangex1 = 69
            rangex2 = 69
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
                
                Dim failure As Integer
                failure = 0
                
                For Each cel In Range(Cells(rangey1, rangex1 + 1), Cells(rangey2 - 1, rangex2 - 1))
                    If cel.Style = "ff" Then
                        If Cells(1 + cel.Row - rangey1, 20 + cel.Column - rangex1 - 1).Style = "sf" Then
                            failure = 1
                        End If
                        Cells(1 + cel.Row - rangey1, 20 + cel.Column - rangex1 - 1).Style = "ff"
                    End If
                Next cel
                    
                If failure = 1 Then
                    Cells(10, 35).Select
                    MsgBox ("Èãðà îêîí÷åíà")
                    Cells(3, 1).Value = 0
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    Range(Cells(14, 28), Cells(18, 32)).Style = "field"
    
    number = 0
    Dim rangex11 As Integer
    Dim rangex21 As Integer
    Select Case nextrotate
        Case 1
            rangex11 = 50
            rangex21 = 50
        Case 2
            rangex11 = 57
            rangex21 = 57
        Case 3
            rangex11 = 62
            rangex21 = 62
        Case 4
            rangex11 = 69
            rangex21 = 69
    End Select
    For i = 1 To 40
        If Cells(i, rangex11).Style = "border" Then
            number = number + 1
            If number = nextrandom Then
                Dim rangey11 As Integer
                Dim rangey21 As Integer
                rangey11 = i + 1
                rangey21 = i + 1
                Do
                    rangex21 = rangex21 + 1
                Loop While Cells(i, rangex21).Style <> "border"
                
                Do
                    rangey21 = rangey21 + 1
                Loop While Cells(rangey21, rangex21).Style <> "border"
                
                Range(Cells(rangey11, rangex11 + 1), Cells(rangey21 - 1, rangex21 - 1)).copy
                Cells(15, 29).PasteSpecial
            End If
        End If
    Next i
    
    
    
    Cells(10, 35).Select
    Cells(4, 1).Value = 1
    Cells(4, 2).Value = 20
End Sub

