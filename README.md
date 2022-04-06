# macro
Change_font.xlsm checks the data pasted into Excel one character at a time, and if it is a single byte character, the font will be Times New Roman. Convert to MS Mincho.
Since it is often the case that the font for one-byte alphanumeric characters is specified as Times New Roman and Japanese as MS Mincho in the document to be submitted, we have made it easy to change the font.
The code below has been modified to perform the conversion on a range of selected cells rather than on unique cells. Please put it in your personal macro book and use it.
However, for unknown reasons, the Excel file will be corrupted and cannot be saved when executed. So, copy and paste the converted result into another book and save it.
You will need to reopen the file to reflect the changes.

Translated with www.DeepL.com/Translator (free version)


Change_font.xlsmはExcelに貼り付けたデータを1文字ずつチェックし１バイト文字であればフォントをTimes New Roman、２バイト文字であればフォントをＭＳ 明朝に変換します。
提出文書で半角英数字のフォントはTimes New Roman、日本語はＭＳ 明朝と指定されていることがよくあるので簡単に変更できるように作りました。
下記のコードは固有セルではなく範囲選択したセルに対して変換を行うよう改良したものです。個人用マクロブックにいれて使ってください。
ただし、原因不明ですが実行するとExcelファイルが壊れてSaveできなくなります。ですので変換後の結果を別ブックにコピー＆ペーストして保存し
改めてファイルを開きなおして反映させるなどの工夫が必要です。


---------

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
