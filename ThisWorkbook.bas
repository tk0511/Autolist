Public EXTRA_CODE_OK As Boolean

Private Sub Workbook_Open()
    If Mid(ThisWorkbook.Name, 2, 2) = "备份" Then
        MsgBox ("该表格是备份表格，使用前需要去除文件名中的[备份]字样以保证表格正常工作！")
    Else
        If isNewMonth Then
            Call myStatusBar("正在备份，不要关闭工作簿！...")
            Dim backupPath As String
            backupPath = ThisWorkbook.PATH & "\备份\" & Year(Date) & "\" & MonthName(Month(Date) - 1)
            Call createPath(backupPath)
            Call ThisWorkbook.SaveCopyAs(backupPath & "\[备份]" & ThisWorkbook.Name)
            
            Call createMonthlyRecord(backupPath)
            Dim deletedSheetCounter As Integer
            PRICECOUNTER_V = True
            Call deleteExpiredData(deletedSheetCounter)
            PRICECOUNTER_V = False
            
            Call chgValue("上次备份日期", Date)
            ThisWorkbook.Save
            Call myStatusBar("备份完毕！" & "共删除过期清单" & deletedSheetCounter & "张，保存目录：" & backupPath, 15)
        End If
    End If

    Call update
    
End Sub

Sub update()
    On Error GoTo reverse
    Dim pkgName As String
    Dim waitUntill
    Dim fileName() As String
    
    pkgName = Dir(ThisWorkbook.PATH + "\update\*.txt")
    If Len(pkgName) <= 0 Then Exit Sub
    If Not newerVersion(pkgName) Then Exit Sub
    If MsgBox("发现新版本 v" + Left(pkgName, Len(pkgName) - 4) + "，是否更新？", vbYesNo) = vbNo Then Exit Sub
    
    fileName = Split(Dir(ThisWorkbook.PATH + "\update\*"), ".")
    Do While Not emptyArr(fileName)
        If fileName(UBound(fileName)) <> "txt" And fileName(UBound(fileName)) <> "frx" Then
            If VBExist(fileName(0)) Then
                ThisWorkbook.VBProject.VBComponents(fileName(0)).Name = fileName(0) + "_DEL"
            End If
            ThisWorkbook.VBProject.VBComponents.Import (ThisWorkbook.PATH + "\update\" + Join(fileName, "."))
            Debug.Print fileName(0) + " - updated"
        End If
        fileName = Split(Dir, ".")
    Loop
    
    For Each component In ThisWorkbook.VBProject.VBComponents
        If Right(component.Name, 4) = "_DEL" Then ThisWorkbook.VBProject.VBComponents.Remove component
    Next
    

    Application.OnTime Now, "extraCode.runExtraCode"

    
    Exit Sub
reverse:
    Call MsgBox("升级失败，退回至上个版本。（请勿保存）")
    Application.DisplayAlerts = False
    ThisWorkbook.Close
End Sub

Function newerVersion(ByRef pkgName As String) As Boolean
    newerVersion = True
    pkgversion = Split(pkgName, ".")
    thisversion = Split(getValue("v"), ".")
    
    If pkgversion(0) <= thisversion(0) And pkgversion(1) <= thisversion(1) And pkgversion(2) <= thisversion(2) Then
        newerVersion = False
    End If
End Function

Function VBExist(ByRef Name As String) As Boolean
    VBExist = True
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Name = Name Then Exit Function
    Next
    VBExist = False
End Function

