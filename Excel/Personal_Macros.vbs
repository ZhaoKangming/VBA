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




'============================================================================================'
'TODO:生成表格内容层级结构
