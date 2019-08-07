' 数据清洗工具
Sub DepartData_Clean()
    Application.ScreenUpdating = False
    Dim ICU_Arr, CCU_Arr, Punctuation_Arr, NumbDepart_Arr, IrrelevantStr_Arr, Hosp_Arr, NoDepart_Arr, NumbRegion_Arr
    Dim i&, rng As Range
    ICU_Arr = Array("icu","lcu","lCU","LCU","ＩＣＵ","Icu","iCu","icU","ICu","ICu","IcU","iCU")
    CCU_Arr = Array("ccu","Ccu","cCu","ccU","CCu","CcU","cCU")
    Punctuation_Arr = Array(".","。",",","，","-","_","-","—","=","+","！","(",")","（","）")
    IrrelevantStr_Arr = Array("（一）","（一）","（二）","（二）")

    With Selection
        .Replace " ",""
        .Replace "其它","其他"
        .Replace "-请选择-","其他"
        .Replace "不确定","其他"
        .Replace "科科","科"
        .Replace "&","、"

        ' 拼音问题
        .Replace "neike","内科"
        .Replace "waike","外科"
        .Replace "guke","骨科"
        .Replace "jizhen","急诊"
        .Replace "fuchan","妇产"
        .Replace "zhonghe","综合"
        .Replace "heci","核磁"
        .Replace "gangchang","肛肠"
        .Replace "zhuyuanbu","住院部"
        .Replace "ke","科"
        
        .Replace "x光","X光"
        .Replace "b超","B超"
        .Replace "ct","CT"
        .Replace "Ct","CT"
        .Replace "cT","CT"


        '缺字少字
        .Replace "眼耳鼻科","眼耳鼻喉科"
        .Replace "卫生服中心","卫生服务中心"
        
        ' 缩写的部分补充为全称
        .Replace "神内","神经内科"
        .Replace "神外","神经外科"
        .Replace "计生","计划生育"
        .Replace "计免","计划免疫"
        .Replace "公卫","公共卫生"

        '修改错别字部分
        .Replace "女姓","女性"
        .Replace "男姓","男性"
        .Replace "小二","小儿"
        .Replace "超生","超声"
        .Replace "终合","综合"
        .Replace "急診","急诊"

        For i = 0 TO UBound(Punctuation_Arr)
            .Replace Punctuation_Arr(i),""
        Next

        For i = 0 TO UBound(ICU_Arr)
            .Replace ICU_Arr(i),"ICU"
        Next

        For i = 0 TO UBound(CCU_Arr)
            .Replace CCU_Arr(i),"ICU"
        Next

        For i = 0 TO UBound(IrrelevantStr_Arr)
            .Replace IrrelevantStr_Arr(i),""
        Next
    End With

    
    Hosp_Arr = Array("服务中心","服务站","医院","卫生院","卫生室","卫生所","卫生站","中心站","社区","诊所","工作室", _
                    "居委会","医疗中心","小学","中学","大学")
    For Each rng in Selection
        ' 去除科室中的医院信息
        For i = 0 TO UBound(Hosp_Arr)
            If rng.Value Like "*" & Hosp_Arr(i) & "*" Then 
                rng.Value = Right(rng.Value, Len(rng.Value) - InStr(rng.Value, Hosp_Arr(i)) - Len(Hosp_Arr(i))+1)
            End If
        Next

        If Application.WorksheetFunction.IsNumber(rng) = True Then rng.Value = "其他"   ' 纯数字单元格变为其他
        If rng.Value <> "" And InStr("一二三四五六七八九十",Right(rng.Value,1))>0 Then rng.Value = Left(rng.Value,Len(rng.Value) -1) & "科"
    Next

    'TODO:数字转中文数字

    NumbDepart_Arr = Array("十二科","十一科","一科","二科","三科","四科","五科","六科","七科","八科","九科","十科", _
                            "12科","11科","10科","9科","8科","7科","6科","5科","4科","3科","2科","1科")
    NumbRegion_Arr = Array("十二区","十一区","一区","二区","三区","四区","五区","六区","七区","八区","九区","十区", _
                            "12区","11区","10区","9区","8区","7区","6区","5区","4区","3区","2区","1区")
    With Selection
        For i = 0 TO UBound(NumbDepart_Arr)
            .Replace NumbDepart_Arr(i),"科"
        Next
        For i = 0 TO UBound(NumbRegion_Arr)
            .Replace NumbRegion_Arr(i),"区"
        Next
        .Replace "科科","科"
        .Replace "","其他"       ' 补全空白单元格
    End With

    NoDepart_Arr = Array("内","外","皮肤","肿瘤","护理","辅助","肾脏","消化","乳腺","男","传染病","产","病理","保健","急诊", _
                        "急救","分泌","放射","风湿","妇产","妇保","肝胆","传染","骨","呼吸","介入","精神","康复","口腔", _
                        "耳鼻喉","老年","检验","结核","护理","防疫","儿","肺病")
    ' ,"","","","","",""
    For Each rng in Selection
        ' 结尾缺少科的，给补上
        For i = 0 To UBound(NoDepart_Arr)
            If Right(rng.Value,Len(NoDepart_Arr(i))) = NoDepart_Arr(i) Then rng.Value = rng.Value & "科"
        Next

        If rng.Value = "大内科" Then rng.Value = "内科"
        If rng.Value = "综合内科" Then rng.Value = "内科"
        If rng.Value = "大外科" Then rng.Value = "外科"
        If rng.Value = "综合外科" Then rng.Value = "外科"
        If rng.Value = "科" Then rng.Value = "其他"
        If rng.Value = "B超" Then rng.Value = "B超室"

        ' 科室合并
        If rng.Value Like "*药房*" Then rng.Value = "药房"
        If rng.Value Like "*ICU*" Then rng.Value = "ICU"
        If rng.Value Like "*人事*" Then rng.Value = "人事科"
        If rng.Value Like "*办事*" Then rng.Value = "办事处"
        If rng.Value Like "*客服*" Then rng.Value = "客服部"
        If rng.Value Like "*教务*" Then rng.Value = "教务部"
        If rng.Value Like "*教学*" Then rng.Value = "教务部"
        If rng.Value Like "*公共卫生*" Then rng.Value = "公共卫生部"
        If rng.Value Like "*行政*" Then rng.Value = "行政部"
        If rng.Value Like "*新生儿*" Then rng.Value = "新生儿科"
        If rng.Value Like "*中西*" Then rng.Value = "中西医结合科"
        If rng.Value Like "*彩超*" Then rng.Value = "彩超科"
        If rng.Value Like "*住院*" Then rng.Value = "住院部"
        If rng.Value Like "*门诊*" Then rng.Value = "门诊部"
        If rng.Value Like "*急诊*" Then rng.Value = "急诊科"
        If rng.Value Like "*产前*" Then rng.Value = "产科"
        If rng.Value Like "*产卡*" Then rng.Value = "产科"
        If rng.Value Like "*病区*" Then rng.Value = "病区"
        If rng.Value Like "*病房*" Then rng.Value = "病区"
        If rng.Value Like "*高压氧*" Then rng.Value = "高压氧科"
        If rng.Value Like "*病案*" Then rng.Value = "病案室"
        
        
    Next

    Selection.Replace "科科","科"
    Selection.Replace "科区","科"
    
    'TODO:结尾为科和室这些要统一，比如B超科与B超室应该是一致的

    Msgbox "Finished Data clean!"
    Application.ScreenUpdating = True
End Sub

Sub HospData_Clean()

End Sub