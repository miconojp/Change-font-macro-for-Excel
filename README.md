# macro
Sub TEST7()
    Dim i As Long
    
    For i = 1 To 50
    '縦軸の第1軸の数値
    With ActiveSheet.ChartObjects(i).Chart.Axes(xlValue, 1).TickLabels
        .Font.Name = "Times New Roman"
    End With
    With ActiveSheet.ChartObjects(i).Chart.Axes(xlCategory, 1).TickLabels
        .Font.Name = "Times New Roman"
    End With
    
    Next i
End Sub

Sub SelectCelltoChangeFont()

Dim r As Range

Set r = Selection


    For Each c In r
    
        AllLength = Len(c.Value)
        ALLDATA = c.Value
        
        For i = 1 To AllLength
                
                OneWord = Mid(ALLDATA, i, 1)
                a = Len(OneWord)
                b = LenB(StrConv(OneWord, vbFromUnicode))
                
                If a <> b Then
                    With c.Characters(Start:=i, Length:=1).Font
                        .Name = "ＭＳ 明朝"
                        .Size = 9
                    End With
                ElseIf a = b Then
                    With c.Characters(Start:=i, Length:=1).Font
                          .Name = "Times New Roman"
                          .Size = 9
                    End With
                End If
        Next i
    
    
    Next c
    
    
End Sub
