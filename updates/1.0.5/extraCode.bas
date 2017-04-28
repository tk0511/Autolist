Attribute VB_Name = "extraCode"
Sub runExtraCode()
    On Error GoTo reverse
    
    Call code.chgValue("v", "1.0.5")
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
