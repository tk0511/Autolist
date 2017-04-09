Attribute VB_Name = "tools"
Sub myTimer(Optional ByVal msg As String = "")
    Static time
    If time = 0 Then
        time = Timer
    Else
        time = Timer - time
        Debug.Print (msg & " : " & Format(time, "0.000"))
        time = 0
    End If
End Sub

Sub myStatusBar(Optional ByVal tip As String = "", Optional ByVal duration As Integer = 0)
    With ThisWorkbook
        If Len(tip) = 0 Then
        Application.statusbar = False
        Else
        Application.DisplayStatusBar = True
        Application.statusbar = tip
        If duration > 0 Then Application.OnTime Now + TimeValue("00:00:" & duration), "myStatusBar"
        End If
    End With
End Sub

Function isProtected(ByVal sheetName As String) As Boolean
    isProtected = Sheets(sheetName).ProtectContents
End Function

Function isInt(ByVal text As String) As Boolean
    If IsNumeric(text) Then
        If Fix(text) - text = 0 Then
            isInt = True
        End If
    Else
        isInt = False
    End If
End Function

Sub errPrint(ByVal code As String, ByVal msg As String)
    Debug.Print (Chr(10) & "[" & Now & "] - " & code & " :")
    Debug.Print (Chr(9) & ">" & msg & "<")
End Sub

Function toInt(ByVal str As String) As Integer
    On Error Resume Next
    toInt = Int(str)
    If toInt <> 0 Then GoTo funcEnd
    toInt = Int(Val(str))
    If toInt <> 0 Then GoTo funcEnd
    toInt = 0
funcEnd:
End Function

Function pathExist(ByRef PATH As String) As Boolean
    pathExist = (Dir(PATH, vbDirectory) <> "")
End Function

Sub createPath(ByRef PATH As String)
    If Not pathExist(PATH) Then
        Call createPath(Left(PATH, InStrRev(PATH, "\") - 1))
        Call MkDir(PATH)
    End If
End Sub

Function emptyArr(ByVal sArray As Variant) As Boolean
    Dim i As Long
    emptyArr = False
    On Error GoTo lerr:
    If UBound(sArray) > 0 Then Exit Function
lerr:
    emptyArr = True
End Function

Sub getAllRef()
    Dim n As Integer
    On Error Resume Next
    For n = 1 To ThisWorkbook.VBProject.References.count
        Cells(n, 1) = ThisWorkbook.VBProject.References.item(n).Name
        Cells(n, 2) = ThisWorkbook.VBProject.References.item(n).Description
        Cells(n, 3) = ThisWorkbook.VBProject.References.item(n).GUID
        Cells(n, 4) = ThisWorkbook.VBProject.References.item(n).Major
        Cells(n, 5) = ThisWorkbook.VBProject.References.item(n).Minor
        Cells(n, 6) = ThisWorkbook.VBProject.References.item(n).fullpath
    Next n
End Sub

Function test()
Dim ww(1 To 20) As Integer
Dim result As New PageSummery
Set test = result
    'records = Range(.Cells(5, 1), .Cells(5, pageWidth))
    'For row = 1 To UBound(records)

    'Next
    Call setVAL_D
End Function
