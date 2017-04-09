'updated file tools code this workbookOpen


Sub runExtraCode()
    On Error GoTo reverse

    Call code.editOn("值")
    ThisWorkbook.Sheets("值").Cells(39, 1) = "件数列"
    ThisWorkbook.Sheets("值").Cells(39, 2) = 7
    ThisWorkbook.Sheets("值").Cells(40, 1) = "备注列"
    ThisWorkbook.Sheets("值").Cells(40, 2) = 14
    ThisWorkbook.Sheets("值").Cells(41, 1) = "杂费列"
    ThisWorkbook.Sheets("值").Cells(41, 2) = 17
    Call code.editOff("值")
    
    Call code.editOn("样本")
    ThisWorkbook.Sheets("样本").Range("N42:Q45").Merge
    ThisWorkbook.Sheets("样本").Range("K5").Formula = "=IF(L5<>""外付"",-I5-J5,H5-I5-J5)"
    ThisWorkbook.Sheets("样本").Range("K5").AutoFill destination:=ThisWorkbook.Sheets("样本").Range("K5:K39"), Type:=xlFillDefault
    With ThisWorkbook.Sheets("样本").Range("K5:L39").Validation
        .Delete
        .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="内付,外付,内欠,外欠"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    ThisWorkbook.Sheets("样本").Columns("L:L").FormatConditions.add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""外欠"""
    ThisWorkbook.Sheets("样本").Columns("L:L").FormatConditions(ThisWorkbook.Sheets("样本").Columns("L:L").FormatConditions.count).SetFirstPriority
    With ThisWorkbook.Sheets("样本").Columns("L:L").FormatConditions(1).Font
        .Bold = True
    End With
    With ThisWorkbook.Sheets("样本").Columns("L:L").FormatConditions(1).Interior
        .Pattern = xlGray25
        .PatternThemeColor = xlThemeColorAccent3
        .ColorIndex = xlAutomatic
        .PatternTintAndShade = 0
    End With
    ThisWorkbook.Sheets("样本").Columns("L:L").FormatConditions(1).StopIfTrue = False
    ThisWorkbook.Sheets("样本").Cells(1, 1) = getValue("清单头")
    Call code.editOff("样本")
    
    Call code.chgValue("v", "1.0.2")
    Call code.setVAL_D
    Dim colCounter As Integer
    colCounter = 1
    While Len(ThisWorkbook.Sheets("价格").Cells(1, colCounter)) > 0
        code.editOn (ThisWorkbook.Sheets("价格").Cells(1, colCounter).text)
        ThisWorkbook.Sheets(ThisWorkbook.Sheets("价格").Cells(1, colCounter).text).Cells(1, 1) = getValue("清单头")
        code.editOff (ThisWorkbook.Sheets("价格").Cells(1, colCounter).text)
        colCounter = colCounter + getValue("价格单宽度")
    Wend

    Debug.Print "extra code runed"
    ThisWorkbook.VBProject.References.AddFromGuid GUID:="{2A75196C-D9EB-4129-B803-931327F72D5C}", Major:=2, Minor:=8
    
    Exit Sub
reverse:
    Call MsgBox("升级失败，退回至上个版本。（正在关闭工作簿，请勿保存）")
    Application.DisplayAlerts = False
    ThisWorkbook.Close
End Sub
