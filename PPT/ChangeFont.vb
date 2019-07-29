Sub ChangeFont()
    Dim oShape As Shape 
    Dim oSlide As Slide 
    Dim oTxtRange As TextRange 
    On Error Resume Next 
    For Each oSlide In ActivePresentation.Slides    
        For Each oShape In oSlide.Shapes 
            Set oTxtRange = oShape.TextFrame.TextRange           
            If Not IsNull(oTxtRange) Then          
                With oTxtRange.Font 
                    .Name = "微软雅黑"       '改成你需要的字体              
                    ' .Size = 20       '改成你需要的文字大小 
                    ' .Color.RGB = RGB(Red:=255, Green:=0, Blue:=0) '改成你想要的文字颜色           
                End With           
            End If    
        Next    
    Next 
End Sub 
