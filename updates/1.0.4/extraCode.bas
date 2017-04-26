Attribute VB_Name = "extraCode"
'updated file tools code this workbookOpen


Sub runExtraCode()
    On Error GoTo reverse
    
    Call code.chgValue("v", "1.0.4")
    Call editVBA
    Call checkImport
    
    Application.OnTime Now, "ThisWorkbook.checkUpdate"
    Exit Sub
reverse:
    Call MsgBox("升级失败，退回至上个版本。（正在关闭工作簿，请勿保存）")
    Application.DisplayAlerts = False
    ThisWorkbook.Close
End Sub

Sub checkImport()
    Dim import()
    import = Array(Array("{000204EF-0000-0000-C000-000000000046}", "4", "2"), Array("{00020813-0000-0000-C000-000000000046}", "1", "9"), Array("{00020430-0000-0000-C000-000000000046}", "2", "0"), Array("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", "2", "8"), Array("{0002E157-0000-0000-C000-000000000046}", "5", "3"), Array("{2A75196C-D9EB-4129-B803-931327F72D5C}", "2", "8"), Array("{0D452EE1-E08F-101A-852E-02608C4D0BB4}", "2", "0"))
    
    For Each pkg In import
        For i = 1 To ThisWorkbook.VBProject.References.count
            If ThisWorkbook.VBProject.References.item(i).GUID = pkg(0) Then pkg(0) = ""
        Next
        If Len(pkg(0)) > 0 Then ThisWorkbook.VBProject.References.AddFromGuid pkg(0), pkg(1), pkg(2)
    Next
End Sub

Sub editVBA()
    For Each Sheet In ThisWorkbook.Sheets
        If Sheet.Cells(1, 1) = getValue("清单头") And Sheet.Visible Then
            i = Sheet.CodeName
            Call ThisWorkbook.VBProject.VBComponents(i).CodeModule.DeleteLines(25, 44)
            Call ThisWorkbook.VBProject.VBComponents(i).CodeModule.AddFromString(newOnChange())
        End If
    Next
    i = ThisWorkbook.Sheets("样本").CodeName
    Call ThisWorkbook.VBProject.VBComponents(i).CodeModule.DeleteLines(25, 44)
    Call ThisWorkbook.VBProject.VBComponents(i).CodeModule.AddFromString(newOnChange())
End Sub

Function newOnChange() As String
    str1 = "Private Sub Worksheet_Change(ByVal target As Range)" & Chr(10) & _
        "    On Error GoTo exitsub" & Chr(10) & _
        "    If Name = ""样本"" Then Exit Sub" & Chr(10) & _
        "    If target.rows.count < Int(Cells(1, Int(getValue(""清单长度列""))) - 10) And target.row >= 5 And target.row <= Int(Cells(1, Int(getValue(""清单长度列""))) - 7) Then" & Chr(10) & _
        "        Application.EnableEvents = False" & Chr(10) & _
        "        Dim emptyRecord As Boolean" & Chr(10) & _
        "        Dim targetRow As Integer" & Chr(10) & _
        "        For targetRow = target.row To target.row + target.rows.count - 1" & Chr(10) & _
        "            emptyRecord = True" & Chr(10) & _
        "            For Each c In Range(Cells(targetRow, 2), Cells(targetRow, Int(getValue(""清单宽度""))))" & Chr(10) & _
        "                If c <> 0 And Len(c) > 0 And c.Column <> Int(getValue(""物流收货日期列"")) Then" & Chr(10) & _
        "                    emptyRecord = False" & Chr(10) & _
        "                End If" & Chr(10) & _
        "            Next" & Chr(10) & _
        "            If emptyRecord Then" & Chr(10) & _
        "                Cells(targetRow, Int(getValue(""物流收货日期列""))).ClearContents" & Chr(10) & _
        "            Else" & Chr(10) & _
        "                If Len(Cells(targetRow, Int(getValue(""物流收货日期列"")))) = 0 Then" & Chr(10) & _
        "                    Cells(targetRow, Int(getValue(""物流收货日期列""))) = Now" & Chr(10) & _
        "                End If" & Chr(10)


    str2 = "            End If" & Chr(10) & _
        "        Next" & Chr(10) & _
        "        Application.EnableEvents = True" & Chr(10) & _
        "    End If" & Chr(10) & _
        "    " & Chr(10) & _
        "    Dim i, fstrow, lstrow As Integer" & Chr(10) & _
        "    Dim cel As Range" & Chr(10) & _
        "    Dim listwidth As Integer" & Chr(10) & _
        "    listwidth = Val(getValue(""清单宽度""))" & Chr(10) & _
        "    i = Cells(1, Val(getValue(""清单长度列""))) - 7" & Chr(10) & _
        "    fstrow = target.row" & Chr(10) & _
        "    lstrow = fstrow - 1 + target.rows.count" & Chr(10) & _
        "    If i >= fstrow And i <= lstrow Then" & Chr(10) & _
        "        For Each cel In Range(Cells(i, 2), Cells(i, listwidth))" & Chr(10) & _
        "            On Error GoTo nextloop" & Chr(10) & _
        "            If (((cel.Column > 1 And cel.Column < 7) Or (cel.Column > 11 And cel.Column < 17)) And Len(cel.value) > 0) Or ((cel.Column > 6 And cel.Column < 12) And Len(cel.value) > 0 And (Val(cel.value) <> 0 Or Not isInt(cel.value))) Then" & Chr(10) & _
        "                Call editOn" & Chr(10) & _
        "                new_line (15)" & Chr(10) & _
        "                Call editOff" & Chr(10) & _
        "                Exit Sub" & Chr(10)


    str3 = "            End If" & Chr(10) & _
        "nextloop: Next" & Chr(10) & _
        "    End If" & Chr(10) & _
        "exitsub:" & Chr(10) & _
        "    Application.EnableEvents = True" & Chr(10) & _
        "End Sub" & Chr(10)

    newOnChange = str1 & str2 & str3
End Function
