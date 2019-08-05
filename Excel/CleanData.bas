' 数据清洗工具
Sub DepartData_Clean()
    Application.ScreenUpdating = False
    Dim ICU_Arr, Punctuation_Arr, NumbDepart_Arr, Hosp_Arr, i, rng As Range
    ICU_Arr = Array("icu","lcu","lCU","LCU","ＩＣＵ")
    Punctuation_Arr = Array(".","。",",","，","-","_","-","—","=","+","！")
    
    With Selection
        .Replace " ",""
        .Replace "其它","其他"
        .Replace "-请选择-","其他"
        .Replace "科科","科"
        .Replace "超生","超声"
        .Replace "终合","综合"
        .Replace "急診","急诊"
        .Replace "卫生服中心","卫生服务中心"
        .Replace "&","、"
        .Replace "neike","内科"
        .Replace "waike","外科"
        .Replace "guke","骨科"
        .Replace "jizhen","急诊"
        .Replace "ke","科"
        

        For i = 0 TO UBound(Punctuation_Arr)
            .Replace Punctuation_Arr(i),""
        Next

        For i = 0 TO UBound(ICU_Arr)
            .Replace ICU_Arr(i),"ICU"
        Next
    End With

    
    Hosp_Arr = Array("服务中心","服务站","医院","卫生院","卫生室","卫生所","卫生站","中心站","社区","诊所","工作室")
    For Each rng in Selection
        ' 去除科室中的医院信息
        For i = 0 TO UBound(Hosp_Arr)
            If rng.Value Like "*" & Hosp_Arr(i) & "*" Then 
                rng.Value = Right(rng.Value, Len(rng.Value) - InStr(rng.Value, Hosp_Arr(i)) - Len(Hosp_Arr(i))+1)
            End If
        Next

        If rng.Value = "科" Then rng.Value = "其他"


        ' 科室合并
        If rng.Value Like "*中西*" Then rng.Value = "中西医结合科"
        If rng.Value Like "*彩超*" Then rng.Value = "彩超科"
        If rng.Value Like "*住院*" Then rng.Value = "住院部"

    Next

    '单元格全部为数字的转化为其他

    '数字转中文数字
    If InStr("一二三四五六七八九十")

    NumbDepart_Arr = Array("十二科","十一科","一科","二科","三科","四科","五科","六科","七科","八科","九科","十科","","","")
    With Selection
        .Replace "","其他"       ' 补全空白单元格
    End With

    For Each rng in Selection
        If rng.Value = "科" Then rng.Value = "其他"
    Next

    '结尾处理
    '结尾为“医”的 卫生 人民 结合 
    

    Msgbox "Finished Data clean!"
    Application.ScreenUpdating = True
End Sub

Sub HospData_Clean()

End Sub