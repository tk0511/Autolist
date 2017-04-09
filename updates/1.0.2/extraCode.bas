'updated file tools code this workbookOpen


Sub runExtraCode()
    On Error GoTo reverse

    Call code.editOn("ֵ")
    ThisWorkbook.Sheets("ֵ").Cells(39, 1) = "������"
    ThisWorkbook.Sheets("ֵ").Cells(39, 2) = 7
    ThisWorkbook.Sheets("ֵ").Cells(40, 1) = "��ע��"
    ThisWorkbook.Sheets("ֵ").Cells(40, 2) = 14
    ThisWorkbook.Sheets("ֵ").Cells(41, 1) = "�ӷ���"
    ThisWorkbook.Sheets("ֵ").Cells(41, 2) = 17
    Call code.editOff("ֵ")
    
    Call code.editOn("����")
    ThisWorkbook.Sheets("����").Range("N42:Q45").Merge
    ThisWorkbook.Sheets("����").Range("K5").Formula = "=IF(L5<>""�⸶"",-I5-J5,H5-I5-J5)"
    ThisWorkbook.Sheets("����").Range("K5").AutoFill destination:=ThisWorkbook.Sheets("����").Range("K5:K39"), Type:=xlFillDefault
    With ThisWorkbook.Sheets("����").Range("K5:L39").Validation
        .Delete
        .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="�ڸ�,�⸶,��Ƿ,��Ƿ"
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
    ThisWorkbook.Sheets("����").Columns("L:L").FormatConditions.add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""��Ƿ"""
    ThisWorkbook.Sheets("����").Columns("L:L").FormatConditions(ThisWorkbook.Sheets("����").Columns("L:L").FormatConditions.count).SetFirstPriority
    With ThisWorkbook.Sheets("����").Columns("L:L").FormatConditions(1).Font
        .Bold = True
    End With
    With ThisWorkbook.Sheets("����").Columns("L:L").FormatConditions(1).Interior
        .Pattern = xlGray25
        .PatternThemeColor = xlThemeColorAccent3
        .ColorIndex = xlAutomatic
        .PatternTintAndShade = 0
    End With
    ThisWorkbook.Sheets("����").Columns("L:L").FormatConditions(1).StopIfTrue = False
    ThisWorkbook.Sheets("����").Cells(1, 1) = getValue("�嵥ͷ")
    Call code.editOff("����")
    
    Call code.chgValue("v", "1.0.2")
    Call code.setVAL_D
    Dim colCounter As Integer
    colCounter = 1
    While Len(ThisWorkbook.Sheets("�۸�").Cells(1, colCounter)) > 0
        code.editOn (ThisWorkbook.Sheets("�۸�").Cells(1, colCounter).text)
        ThisWorkbook.Sheets(ThisWorkbook.Sheets("�۸�").Cells(1, colCounter).text).Cells(1, 1) = getValue("�嵥ͷ")
        code.editOff (ThisWorkbook.Sheets("�۸�").Cells(1, colCounter).text)
        colCounter = colCounter + getValue("�۸񵥿���")
    Wend

    Debug.Print "extra code runed"
    ThisWorkbook.VBProject.References.AddFromGuid GUID:="{2A75196C-D9EB-4129-B803-931327F72D5C}", Major:=2, Minor:=8
    
    Exit Sub
reverse:
    Call MsgBox("����ʧ�ܣ��˻����ϸ��汾�������ڹرչ����������𱣴棩")
    Application.DisplayAlerts = False
    ThisWorkbook.Close
End Sub
