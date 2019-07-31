'Author：ZhaoKangming
'================================CONTENTS======================================='
【1】DelPunctuation —— 用于清除单元格内的所有标点符号
【2】IdValid —— 用于验证身份证号码的有效性
【3】GetDigit —— 用于提取字符串中的数字
【4】FilesCount —— 用于返回文件夹路径（FolderPath）子一级文件和文件夹的数目



'=============================================================================='

'【DelPunctuation】用于清除单元格内的所有标点符号
Public Function DelPunctuation(Txt As String) As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "[^\u4e00-\u9fBF\u0041-\u005a\u0061-\u007\u2150-\u218faa-zA-Z\d]"
        .IgnoreCase = True
        .Global = True
        DelPunctuation = .Replace(Txt, "")
    End With
End Function

'=============================================================================='

'【IdValid】用于验证身份证号码的有效性
Public Function IdValid(IdNumb)
    Dim b As Variant
    a = Len(IdNumb)
    If a = 0 Then IdValid = "空": GoTo 100  '先判断是否为空
    
    b = Right(IdNumb, 1)
    If IsNumeric(b) Then
      b = b * 1
    Else: b = "X"
    End If
    
    C = Left(Right(IdNumb, 18), 1) * 1
    d = Left(Right(IdNumb, 17), 1) * 1
    e = Left(Right(IdNumb, 16), 1) * 1
    f = Left(Right(IdNumb, 15), 1) * 1
    g = Left(Right(IdNumb, 14), 1) * 1
    h = Left(Right(IdNumb, 13), 1) * 1
    i = Left(Right(IdNumb, 12), 1) * 1
    j = Left(Right(IdNumb, 11), 1) * 1
    k = Left(Right(IdNumb, 10), 1) * 1
    l = Left(Right(IdNumb, 9), 1) * 1
    m = Left(Right(IdNumb, 8), 1) * 1
    n = Left(Right(IdNumb, 7), 1) * 1
    o = Left(Right(IdNumb, 6), 1) * 1
    p = Left(Right(IdNumb, 5), 1) * 1
    q = Left(Right(IdNumb, 4), 1) * 1
    r = Left(Right(IdNumb, 3), 1) * 1
    s = Left(Right(IdNumb, 2), 1) * 1
    u = C * 7 + d * 9 + e * 10 + f * 5 + g * 8 + h * 4 + i * 2 + j * 1 + k * 6 + l * 3 + m * 7 + n * 9 + o * 10 + p * 5 + q * 8 + r * 4 + s * 2
    v = u Mod 11
    
    If a = 15 Then IdValid = "老号，请认真核对！": GoTo 100
    
    If a = 18 Then
        If v = 0 And b = 1 Then IdValid = "正确": GoTo 100
        If v = 1 And b = 0 Then IdValid = "正确": GoTo 100
        If v = 2 And b = "X" Then IdValid = "正确": GoTo 100
        If v = 3 And b = 9 Then IdValid = "正确": GoTo 100
        If v = 4 And b = 8 Then IdValid = "正确": GoTo 100
        If v = 5 And b = 7 Then IdValid = "正确": GoTo 100
        If v = 6 And b = 6 Then IdValid = "正确": GoTo 100
        If v = 7 And b = 5 Then IdValid = "正确": GoTo 100
        If v = 8 And b = 4 Then IdValid = "正确": GoTo 100
        If v = 9 And b = 3 Then IdValid = "正确": GoTo 100
        If v = 10 And b = 2 Then IdValid = "正确": GoTo 100
    End If
    
    If (a <> 0) And (a <> 15) And (a <> 18) Then IdValid = "位数不对"
    
    If (IdValid <> "正确") And (IdValid <> "老号，请认真核对！") And (IdValid <> "空") And (IdValid <> "位数不对") Then IdValid = "出错啦"
100:
End Function

'=============================================================================='

'【GetDigit】用于提取字符串中的数字
'Bug 451石巧Y110015 只能提取到451
Public Function GetDigit(strValue As String) As Variant
    Dim objReg As Object, objMatchs As Object, objMatch As Object
    Dim strPat$, intIndex%
    Dim strResult() As Long  '字符型String可以得到[00001],如果想直接得到数字，可以定义为长整型Long
    
    strPat = "\d+\.*\d*"
    
    Set objReg = CreateObject("VBScript.RegExp")
    With objReg
        .Global = True
        .Pattern = strPat
        Set objMatchs = .Execute(strValue)
        intIndex = objMatchs.Count
        If intIndex > 0 Then
            ReDim strResult(1 To objMatchs.Count)
            intIndex = 0
            For Each objMatch In objMatchs
                    intIndex = intIndex + 1
                    strResult(intIndex) = objMatch
            Next
            GetDigit = strResult
        End If
    End With
End Function

'=============================================================================='

'【FilesCount】用于返回文件夹路径（FolderPath）子一级文件和文件夹的数目
Public Function FilesCount(FolderName) As Integer
    Dim MyFile$
    FilesCount = 0
    MyFile = Dir(IIf(Right(FolderName, 1) = "\", FolderName & "*.*", FolderName & "\*.*"), vbNormal + vbDirectory)
    Do While MyFile <> ""
        FilesCount = FilesCount + 1
        MyFile = Dir()
    Loop
    If FilesCount > 0 Then FilesCount = FilesCount - 2  '删除 "." 与 ".." 两个隐藏空文件夹
End Function

'=============================================================================='
' TODO: 获取选中区域的唯一值
Public Function UniqueValue(ValueRange As Range)
    Dim rng As Range, arr, d As Object
    Set d = CreateObject("scripting.dictionary")
    For Each rn In ValueRange
        If rng <> "" And Not d.exists(rng.Value) Then d(rng.Value)= rng.Value
    Next
    arr = d.items
    For i = 0 To d.Count - 1
        Cells(i + 1, 3) = arr(i)
    Next
End Function

' 另一种思路
  Dim lRow As Long
  Dim i As Long
  Dim str As Variant
  Dim strKey As String
 
  lRow = Range("A65536").End(xlUp).Row
' lRow = Cells(Rows.Count,1).End(xlUp).Row
  str = Range("A1:A" & lRow)
  For i = 1 To lRow
    strKey = CStr(str(i, 1))
     If Not d.exists(strKey) Then
        d.Add strKey, strKey
     End If
  Next i
  Range("D1").Resize(UBound(d.keys) + 1, 1) = Application.Transpose(d.keys)


'=============================================================================='
' TODO:【ListNumbCount】用于加强版counttif，自动去重并数数量
横向纵向
Public Function GetFrequency(ValueRange As Range)
    Application.ScreenUpdating = False
    Dim rng As Range, arr, d As Object
    Set d = CreateObject("scripting.dictionary")
    For Each rn In ValueRange
        If rng <> "" And Not d.exists(rng.Value) Then d(rng.Value)= rng.Value
    Next
    arr = d.items
    For i = 0 To d.Count - 1
        Cells(i + 1, 3) = arr(i)
    Next
    
    Set d = Nothing
End Function

