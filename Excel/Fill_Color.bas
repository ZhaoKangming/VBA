Sub FillColor()
    Dim rng As Range
    For each rng in [C2:L34]
        Select Case rng.value
            Case is = "市II类 5.0学分" : rng.Interior.Color = RGB(183, 222, 232)
            Case is = "省级II类 5.0学分" : rng.Interior.Color = RGB(204, 192, 218)
            Case is = "市II类5.0分(远程)" : rng.Interior.Color = RGB(184, 204, 228)
            Case is = "18年国I类 5.0学分" : rng.Interior.Color = RGB(252, 213, 180)
            Case is = "市I类5.0分(远程)" : rng.Interior.Color = RGB(220, 230, 241)
            Case is = "15年国I类 5.0学分" : rng.Interior.Color = RGB(230, 184, 183)
            Case is = "自治区级II类 5.0学分" : rng.Interior.Color = RGB(216, 228, 188)
        End Select
    Next
    ThisWorkbook.Save
    Msgbox "已经处理完成！"
End Sub


