'Author：ZhaoKangming
'====================================== CONTENTS ============================================'




'============================================================================================'

'宏作用:将选中区域（NameRng）的每个单元格内容作为文件夹名在指定路径下（FolderPath）建立文件夹
Public Sub NewFolders()
    Dim NameRng as Range, CellRng as Range, FolderPath$, ErrorOccurred As Boolean
    ErrorOccurred = FALSE
    FolderPath = inputbox("请输入新建文件夹的路径，不以'\'结尾","输入地址")
    'TODO: 如果没有输入地址，Msgbox 并且叫重新输入
    If FolderPath = "" Then FolderPath = ActiveWorkbook.Path
    If len(Dir(FolderPath,16)) = 0 Then MkDir FolderPath  '若路径不存在则新建此文件夹
    Set NameRng = Selection
    For each CellRng in NameRng
        If len(Trim(CellRng)) > 0 Then                    'Trim函数来去除单元格首尾的空格
            If len(Dir(FolderPath & "\" & CellRng,16)) <> 0 Then
                ErrorOccurred = True
                CellRng.Interior.Color = RGB(64, 224, 208) '填充为宝石绿
            Else
                MkDir FolderPath & "\" & CellRng
            End If
        End If 
    Next
    set NameRng = nothing
    If ErrorOccurred = FALSE Then 
        Msgbox "All folders have been created!"
    Else
        Msgbox "部分文件夹已存在，在表中填充为宝石绿以标注！"
    End if
End Sub

'============================================================================================'
'宏作用：将指定路径（FolderPath）下的子一级文件及文件夹列表输出至当前单元格起始的列区域
Public Sub GetFilesList()
    Dim FolderPath$, MyFile, ActiveCol%, ActiveRow%
    ActiveCol = ActiveCell.Column
    ActiveRow = ActiveCell.Row
    'If WorksheetFunction.CountA(Columns(ActiveCol)) > 0 Then
    '    MsgBox "请选择空白列的单元格作为输出起始区域"
    '   Exit Sub
    'End If
    FolderPath = InputBox("请输入新建文件夹的路径，不以'\'结尾", "输入地址") & "\"
    MyFile = Dir(FolderPath, 16)
    Do While MyFile <> ""
        If MyFile <> "." And MyFile <> ".." Then
            ActiveSheet.Cells(ActiveRow, ActiveCol) = MyFile
            ActiveRow = ActiveRow + 1
        End If
        MyFile = Dir
    Loop
    ActiveWorkbook.Save
    MsgBox "已生成子一级文件及文件夹列表！"
End Sub

'============================================================================================'
'TODO:右列表中有重复名字，左侧列表中没有该文件？
'宏作用：将指定路径（FolderPath）下的名字为选区(Selection)中左列的文件重命名为选区中右列的文件名
Public Sub ReNameFiles()
    Dim FolderPath$, MyFile, ReNamedFile, StartCol%, EndCol%, StartRow%, EndRow%, CellNumb%, i%
    'TODO:检查选区的列数，是否空列，是否两列，是否左列有值而右列无值,右列是否有重复值
    CellNumb = Selection.Cells.Count
    StartCol = Selection.Cells(1).Column
    EndCol = Selection.Cells(CellNumb).Column
    StartRow = Selection.Cells(1).Row
    EndRow = Selection.Cells(CellNumb).Row

    FolderPath = InputBox("请输入文件夹的路径，不以'\'结尾", "输入地址") & "\"
    For i = StartRow To EndRow
        If Trim(Cells(i, StartCol)) <> "" Then
            MyFile = FolderPath & Cells(i, StartCol)
            ReNamedFile = FolderPath & Cells(i, EndCol)
            If Len(Dir(MyFile)) > 0 Then
                Name MyFile As ReNamedFile
            Else
                MsgBox "不存在文件 " & Cells(i, StartCol)
            End If
        End If
    Next
    ActiveWorkbook.Save
    MsgBox "All files have been renamed!"
End Sub


'============================================================================================'
'宏作用：将指定路径（FolderPath）下的名字为选区(Selection)中左列的文件夹重命名为选区中右列的文件夹名
Public Sub ReNameFolders()
    Dim FolderPath$, MyFile, ReNamedFile, StartCol%, EndCol%, StartRow%, EndRow%, CellNumb%, i%
    '【TODO】检查选区的列数，是否空列，是否两列，是否左列有值而右列无值,右列是否有重复值
    CellNumb = Selection.Cells.Count
    StartCol = Selection.Cells(1).Column
    EndCol = Selection.Cells(CellNumb).Column
    StartRow = Selection.Cells(1).Row
    EndRow = Selection.Cells(CellNumb).Row

    FolderPath = InputBox("请输入文件夹的路径，不以'\'结尾", "输入地址") & "\"
    For i = StartRow To EndRow
        If Trim(Cells(i, StartCol)) <> "" Then
            MyFile = FolderPath & Cells(i, StartCol)
            ReNamedFile = FolderPath & Cells(i, EndCol)
            If Len(Dir(MyFile, 16)) > 0 Then
                If Len(Dir(ReNamedFile, 16)) > 0 And MyFile <> ReNamedFile Then
                    Cells(i, StartCol).Interior.ColorIndex = 44    '新添加，先添加
                Else
                    Name MyFile As ReNamedFile
                End If
            'Else
                'MsgBox "不存在文件夹 " & Cells(i, StartCol)
            End If
        End If
    Next
    ActiveWorkbook.Save
    MsgBox "All files have been renamed!"
End Sub


'============================================================================================'
'宏作用：将指定路径（FolderPath）下的子一级空文件夹删除
Public Sub DelBlankFolders()
    Dim FolderPath$, MyFile
    FolderPath = InputBox("请输入新建文件夹的路径，不以'\'结尾", "输入地址") & "\"
    MyFile = Dir(FolderPath, 16)
    Do While MyFile <> ""
        If InStr(MyFile, ".") = 0 Then
            On Error Resume Next
            RmDir FolderPath & MyFile
        End If
        MyFile = Dir
    Loop
    ActiveWorkbook.Save
    MsgBox "所有空文件夹已删除！"
End Sub


'============================================================================================'
' 宏作用:将选中区域（NameRng）的每个单元格内容作为文件夹名在指定路径下（FolderPath）删除文件夹
Public Sub DelFolders()
    Dim NameRng As Range, CellRng As Range, FolderPath$, ErrorOccurred As Boolean
    Dim fso As Object
    ErrorOccurred = False
    FolderPath = InputBox("请输入要删除文件夹的路径，不以'\'结尾", "输入地址")
    If FolderPath = "" Then
        MsgBox "请输入路径！"
        FolderPath = InputBox("请输入要删除文件夹的路径，不以'\'结尾", "输入地址")
    End If

    Set NameRng = Selection
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each CellRng In NameRng
        myfolder = FolderPath & "\" & CellRng.Value
        If fso.FolderExists(myfolder) Then
            fso.DeleteFolder myfolder
        Else
            MsgBox "没有文件夹 --" & CellRng.Value
            ErrorOccurred = True
            CellRng.Interior.Color = RGB(64, 224, 208) '填充为宝石绿
        End If
    Next
    
    Set NameRng = Nothing
    Set fso = Nothing
    If ErrorOccurred = False Then
        MsgBox "All folders have been Deleted!"
    Else
        MsgBox "部分文件夹不存在，在表中填充为宝石绿以标注！"
    End If
End Sub


'============================================================================================'
' 宏作用:将选中区域的空行删除
Sub DeleteBlankRows()
    ' 【注意】：如果某一行为空行但是有单元格中含有空格的话是无法检测出并删除的
    ' 【警告】：千万不要选中整个表格（wholesheet），否则遍历会浪费超级超级多时间，请选用（usedrange）
    Dim i As Long 
    Application.ScreenUpdating = False
    '由于要删除行，所以要FOR NEXT要倒着来
    For i = Selection.Rows.Count To 1 Step -1
        If WorksheetFunction.CountA(Selection.Rows(i)) = 0 Then
            Selection.Rows(i).EntireRow.Delete
        End If
    Next i
    Application.ScreenUpdating = True                 
End Sub


'============================================================================================'
' 宏作用:美化表格，设置全边框，首行填充蓝色，字体为白色，关闭网格
Sub BeautifySheet()
    ActiveWorkSheet.UsedRange.Select
    ' 设置字体样式以及居中
    With Selection
        .Font.Name = "微软雅黑"
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter    
    End With
    ActiveWindow.DisplayGridlines = False   ' 隐藏网格线

    ' 设置全边框
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 设置首行标题

    With Selection.Rows(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12611584
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Rows(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

End Sub

'============================================================================================'
'TODO:宏作用: 对表格选定区域的内容进行隔行颜色填充
Sub FillColor()

End Sub
'TODO:可选参数，一，从区域的第几行开始填充，填充什么颜色,中间用空格分隔，如果某参数不填则有一个默认值
'TODO:测试给选中区域的第n行填充颜色，填的是整行还是区域内的行


'============================================================================================'
'TODO:生成表格内容层级结构


'============================================================================================'
'TODO:合并单元格
Sub MergeCells()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim StartCol%, EndCol%, StartRow%, EndRow%, CellNumb%, i%, j%
    'TODO:检查选区的列数，是否空列，是否两列，是否左列有值而右列无值,右列是否有重复值
    CellNumb = Selection.Cells.Count
    StartCol = Selection.Cells(1).Column
    EndCol = Selection.Cells(CellNumb).Column
    StartRow = Selection.Cells(1).Row
    EndRow = Selection.Cells(CellNumb).Row

    For i = EndRow To StartRow Step -1
        If Cells(i,StartCol) = Cells(i - 1,StartCol) Then
            For j = StartCol To EndCol
                Range(Cells(i - 1,j),Cells(i,j)).Merge
            Next
        End If
    Next

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True 
    Msgbox "Finished Cells Merge!"
End Sub

'=============================================================================='
Public Declare PtrSafe Function MsgBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long 'AutoClose
Public Sub AddDif()
    Application.ScreenUpdating = False
    Dim rng As Range, ValueRange As Range
    Set ValueRange = Selection
    Dim New_Max%, Orig_Sum%, New_Sum%, New_Dif%, DifNumb
    Orig_Sum = Application.WorksheetFunction.Sum(ValueRange)
    DifNumb = InputBox("请输入要增添的数值或者以#开头的目标值", "数据差值")
    If Left(DifNumb,1)="#" Then DifNumb = Right(DifNumb, Len(DifNumb)-1) - Orig_Sum
    For Each rng In ValueRange
        If rng <> "" Then rng.Value = CInt(DifNumb * rng.Value / Orig_Sum) + rng.Value
    Next
    New_Sum = Application.WorksheetFunction.Sum(ValueRange)
    New_Max = Application.WorksheetFunction.Max(ValueRange)
    New_Dif = DifNumb + Orig_Sum - New_Sum
    For Each rng In ValueRange
        If rng = New_Max Then
            rng.Value = rng.Value + New_Dif
            Exit For
        End If
    Next
    MsgBoxTimeOut 0,"增加数据成功！", "提示", 64, 0, 300
End Sub

'=============================================================================='
'此宏用于统计选区内的医生职称
Public Sub DocTitle()
    Application.ScreenUpdating = False
    Dim ZR_Numb&, FZR_Numb&, ZZ_Numb&, YS_Numb&, rng As Range, Parameters, DstRow&, DstCol%
    ' 医生职称的统计
    Parameters = InputBox("请输入:行号&列号#有无表头", "输入参数")
    For Each rng in Selection
        If rng Like "*副*" Then
            FZR_Numb = FZR_Numb + 1
        ElseIf rng Like "*主任*" Then
            ZR_Numb = ZR_Numb + 1
        ElseIf rng Like "*主治*" Then
            ZZ_Numb = ZZ_Numb + 1
        Else
            YS_Numb = YS_Numb + 1
        End If
    Next
    DstRow = Left(Parameters,InStr(Parameters,"&")-1)
    DstCol = Mid(Parameters,InStr(Parameters,"&")+1,Len(Parameters)-InStr(Parameters,"&")-2)
    If Right(Parameters,1) = 1 Then
        Cells(DstRow, DstCol) = "主任医师"
        Cells(DstRow + 1, DstCol) = "副主任医师"
        Cells(DstRow + 2, DstCol) = "主治医师"
        Cells(DstRow + 3, DstCol) = "医师"
        Cells(DstRow + 4, DstCol) = "总计"
        
        Cells(DstRow, DstCol + 1)  = ZR_Numb
        Cells(DstRow + 1, DstCol + 1) = FZR_Numb
        Cells(DstRow + 2, DstCol + 1) = ZZ_Numb
        Cells(DstRow + 3, DstCol + 1) = YS_Numb
        Cells(DstRow + 4, DstCol + 1) = ZR_Numb + FZR_Numb + ZZ_Numb + YS_Numb
    Elseif Right(Parameters,1) = 0 Then
        Cells(DstRow, DstCol) = FZR_Numb
        Cells(DstRow + 1, DstCol) = FZR_Numb
        Cells(DstRow + 2, DstCol) = ZZ_Numb
        Cells(DstRow + 3, DstCol) = YS_Numb
        Cells(DstRow + 4, DstCol) = ZR_Numb + FZR_Numb + ZZ_Numb + YS_Numb
    End If
    Msgbox "Finished!"
    Application.ScreenUpdating = True
End Sub

'=============================================================================='
'此宏用于在选定单元格所在的空列生成上一列的职称信息的转化
'TODO:如果所在的列存在数据那么提示
'TODO:用超多、较少的行数老分别测试，可能有问题
'TODO:DstCol and DstRow更正过来
Public Sub DocTitle_Trans()
    Application.ScreenUpdating = False
    Dim i&, LastRow&, DstCol%

    DstRow = Selection.Column
    LastRow = ActiveSheet.Cells(1048576, DstRow - 1).End(xlUp).Row
    For i = Selection.Row To LastRow
        If Cells(i, DstRow - 1) Like "*副*" Then
                Cells(i, DstRow) = "副主任医师"
        ElseIf Cells(i, DstRow - 1) Like "*主任*" Then
            Cells(i, DstRow) = "主任医师"
        ElseIf Cells(i, DstRow - 1) Like "*主治*" Then
            Cells(i, DstRow) = "主治医师"
        Else
            Cells(i, DstRow) = "医师"
        End If
    Next i
    MsgBox "Finished!"
    Application.ScreenUpdating = True
End Sub
'=============================================================================='
'此宏用于从选区内读取要替换和替换词，实现批量替换
Public Sub BatchReplace()

End Sub
'=============================================================================='
'【说明】此程序用于按照已经定好的比例给用户随机分配一个医院级别
Sub Test()
    Application.ScreenUpdating = False
    Randomize
    For i = 2 To [A1048576].End(xlUp).Row
        If Cells(i, 5) = "其他" Then
            If Rnd() < 0.4418 Then
                Cells(i, 5) = "二甲"
            ElseIf Rnd() < 0.6701 Then
                Cells(i, 5) = "三甲"
            ElseIf Rnd() < 0.8455 Then
                Cells(i, 5) = "一甲"
            ElseIf Rnd() < 0.9203 Then
                Cells(i, 5) = "三乙"
            ElseIf Rnd() < 0.9797 Then
                Cells(i, 5) = "二乙"
            Else
                Cells(i, 5) = "一乙"
            End If
        End If
    Next i
    MsgBox "已经处理完成！"
End Sub

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
' TODO:【GetFrequency】自动去重并数数量
' 横向纵向
Public Declare PtrSafe Function MsgBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long 'AutoClose
Public Sub GetFrequency()
    Application.ScreenUpdating = False
    Dim rng As Range, input_Rng As Range, output_Range, arr, dict As Object

    'Part1:将不重复项及出现次数写入到词典
    Set input_Rng = Selection
    Set dict = CreateObject("scripting.dictionary")
    For Each rng In input_Rng
        If rng <> "" And Not dict.exists(rng.Value) Then 
            dict(rng.Value)= Application.WorksheetFunction.Countif(input_Rng,rng.Value)
        End If
    Next
    
    'Part2:将词典内容输出到指定单元格
    Set output_Range = Application.InputBox(prompt:="请选择输出单元格：", Type:=8)
    output_Range.Resize(UBound(dict.keys) + 1, 1) = Application.Transpose(dict.keys)
    Cells(output_Range.Row, output_Range.Column + 1).Resize(UBound(dict.items) + 1, 1) = Application.Transpose(dict.items)
    
    'Part3：对输出进行排序
    ActiveSheet.Range(output_Range,Cells(output_Range.Row + dict.count - 1,output_Range.Column + 1)).Sort key1:=Cells(output_Range.Row,output_Range.Column + 1), _
                order1:=xlDescending, Header:=xlNo, MatchCase:=True

    Set dict = Nothing
    Application.ScreenUpdating = False
    MsgBoxTimeOut 0,"统计数据成功！", "提示", 64, 0, 300
End Sub

'=============================================================================='
'TODO:数据清洗工具；医院、职称、科室、医院级别等等
