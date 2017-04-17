Attribute VB_Name = "extraCode"
'updated file tools code this workbookOpen


Sub runExtraCode()
    On Error GoTo reverse
    

    Call editOn("值")
    ThisWorkbook.Sheets("值").Cells(45, 1) = "DBADD"
    ThisWorkbook.Sheets("值").Cells(45, 2) = "mysql.rdsmk7l09ertw86.rds.bj.baidubce.com"
    ThisWorkbook.Sheets("值").Cells(46, 1) = "DB"
    ThisWorkbook.Sheets("值").Cells(46, 2) = "kangtai"
    Call editOff("值")
    Call editOn("样本")
    ThisWorkbook.Sheets("样本").Cells(41, 1).Formula = "=IF((SUMIF(L5:L39,""外付"",H5:H39)-SUM(I5:I39)-SUM(J5:J39))=SUM(K5:K39),""合计外收 ""&SUM(K5:K39)&"" 元 - ""&P41&"" ""& Q41 &"" 元""&"" = "" &IF(SUM(K5:K39)-Q41>0,""退运费 "",""付运费 "")&ABS(SUM(K5:K39)-Q41)&"" 元"","""")"
    ThisWorkbook.Sheets("样本").Columns("S:Z").Interior.ThemeColor = xlThemeColorDark1
    Call editOff("样本")
    Dim colCounter As Integer
    colCounter = 1
    While Len(ThisWorkbook.Sheets("价格").Cells(1, colCounter)) > 0
        Call code.editOn(ThisWorkbook.Sheets("价格").Cells(1, colCounter).text)
        ThisWorkbook.Sheets(ThisWorkbook.Sheets("价格").Cells(1, colCounter).text).Cells(ThisWorkbook.Sheets(ThisWorkbook.Sheets("价格").Cells(1, colCounter).text).Cells(1, toInt(getValue("清单长度列"))) - 4, 1).Formula = "=IF((SUMIF(L5:L39,""外付"",H5:H39)-SUM(I5:I39)-SUM(J5:J39))=SUM(K5:K39),""合计外收 ""&SUM(K5:K39)&"" 元 - ""&P41&"" ""& Q41 &"" 元""&"" = "" &IF(SUM(K5:K39)-Q41>0,""退运费 "",""付运费 "")&ABS(SUM(K5:K39)-Q41)&"" 元"","""")"
        ThisWorkbook.Sheets(ThisWorkbook.Sheets("价格").Cells(1, colCounter).text).Columns("S:Z").Interior.ThemeColor = xlThemeColorDark1
        Call code.editOff(ThisWorkbook.Sheets("价格").Cells(1, colCounter).text)
        colCounter = colCounter + getValue("价格单宽度")
    Wend


    Call code.chgValue("v", "1.0.3")
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



