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

    Call checkUpdate
    
End Sub

Sub checkUpdate()
    On Error GoTo subEnd
    Dim VersionList() As String
    VersionList = Split(httpGET("https://raw.githubusercontent.com/tk0511/Autolist/master/VersionList").responseText, Chr(10))
    If newerVersion(VersionList(UBound(VersionList))) Then
        For i = 0 To UBound(VersionList)
            If newerVersion(VersionList(i)) Then
                updateTo VersionList(i)
                Exit Sub
            End If
        Next
    End If
subEnd:
End Sub

Sub updateTo(ByVal version As String)
    On Error GoTo reverse
    Dim modeName As String
    Dim baseURL As String
    Dim fileList() As String
    
    baseURL = "https://raw.githubusercontent.com/tk0511/Autolist/master/updates/" & version & "/"
    fileList = Split(Replace(httpGET(baseURL & "__List__").responseText, Chr(13), ""), Chr(10))
    
    For Each fileName In fileList
        Call myDump(CStr(fileName), baseURL)
        If Split(fileName, ".")(1) = "frm" Then Call myDumpByte(Split(fileName, ".")(0) & ".frx", baseURL)
    Next

    For Each fileName In fileList
        modeName = Split(fileName, ".")(0)
        If VBExist(modeName) Then
            ThisWorkbook.VBProject.VBComponents(modeName).Name = modeName & "_DEL"
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(modeName & "_DEL")
        End If
        ThisWorkbook.VBProject.VBComponents.import ThisWorkbook.PATH & "/" & fileName
    Next
    
    For Each fileName In fileList
        Kill ThisWorkbook.PATH & "/" & fileName
        If Split(fileName, ".")(1) = "frm" Then Kill ThisWorkbook.PATH & "/" & Split(fileName, ".")(0) & ".frx"
    Next
    
    Application.OnTime Now, "extraCode.runExtraCode"
    Exit Sub
reverse:
    Call MsgBox("升级失败，退回至上个版本。（请勿保存）")
    Application.DisplayAlerts = False
    ThisWorkbook.Close
End Sub

Function httpGET(ByRef url As String)
    Dim httpREQ As Object
    Set httpREQ = CreateObject("MSXML2.XMLHTTP.3.0")
    httpREQ.Open "GET", url, False
    httpREQ.send
    If httpREQ.Status <> 200 Then Call Err.Raise(vbObjectError + 513, "httpGET", "http GET status code <> 200" & Chr(13) & "URL = " & url)
    Set httpGET = httpREQ
End Function

Sub myDumpByte(ByRef fileName As String, ByRef baseURL As String)
    Set S = CreateObject("ADODB.Stream")
    S.Type = 1
    S.Open
    S.Write httpGET(baseURL & fileName).responseBody
    S.SaveToFile ThisWorkbook.PATH & "/" & fileName, 2
    S.Close
End Sub

Sub myDump(ByRef fileName As String, ByRef baseURL As String)
    Dim text() As String
    text = Split(httpGET(baseURL & fileName).responseText, Chr(10))
    Set S = CreateObject("ADODB.Stream")
    S.Type = 2
    S.Charset = "gbk"
    S.Open
    For Each Line In text
        S.WriteText Line, 1
    Next
    S.SaveToFile ThisWorkbook.PATH & "/" & fileName, 2
    S.Close
End Sub

Function VBExist(ByRef Name As String) As Boolean
    VBExist = True
    For Each Component In ThisWorkbook.VBProject.VBComponents
        If Component.Name = Name Then Exit Function
    Next
    VBExist = False
End Function

Function newerVersion(ByRef version As String) As Boolean
    newerVersion = True
    pkgversion = Split(version, ".")
    thisversion = Split(getValue("v"), ".")
    
    If pkgversion(0) <= thisversion(0) And pkgversion(1) <= thisversion(1) And pkgversion(2) <= thisversion(2) Then
        newerVersion = False
    End If
End Function


