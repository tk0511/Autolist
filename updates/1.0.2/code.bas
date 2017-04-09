Public Type People
    Name As String
    cel As String
    add As String
    id As String
    lst_linker As String
    lst_item As String
    lst_pkg As String
End Type
Public Type PriceRecord
    receiver As String
    address As String
    item As String
    pkg As String
    sender As String
    price As Double
End Type
Public VAL_D As Object
Public PRICECOUNTER_V As Boolean

Sub new_page(Optional ByVal sheetName As String = "", Optional ByVal pageHead_Row As Integer = 1)
    Dim pagehead As String
    pagehead = Cells(1, 1)
    If Len(pagehead) > 4 And pagehead <> getValue("�嵥ͷ") Then GoTo error_process
    If Len(sheetName) <= 0 Then sheetName = ActiveSheet.Name
    If sheetName = "����" Then Exit Sub
    Dim destination As String
    
    On Error GoTo error_process
    Call editOn(sheetName)
    Dim i As Integer
    Dim r As Integer
    Dim pageSize As Integer
    Dim pageWidth As Integer
    Dim sender As People
    Dim receiver As People
    Dim driver As People
    pageSize = Cells(pageHead_Row, toInt(getValue("�嵥������")))
    pageWidth = Int(getValue("�嵥����"))
    
    
    Call myStatusBar("���ڱ�������...")
    If pageSize < 45 Then pageSize = 45
    With Sheets(sheetName)
        destination = .Cells(1, toInt(getValue("�嵥Ŀ�ĵ���")))
        If Len(destination) <= 0 Then destination = sheetName
        .Calculate
        .Range(.Cells(1, 1), .Cells(pageSize, pageWidth)) = .Range(.Cells(1, 1), .Cells(pageSize, pageWidth)).value
        .Range(.Cells(1, 1), .Cells(pageSize - 4, pageWidth)).Locked = True
        For Each cel In .Range(.Cells(1, 1), .Cells(pageSize, pageWidth))
            If IsError(cel.value) Then cel.value = ""
        Next
        
        records = .Range(.Cells(5, 1), .Cells(pageHead_Row + pageSize - 7, pageWidth))

        For row = 1 To UBound(records)
            If records(row, Int(getValue("������"))) > 0 Then
                receiver.Name = records(row, Int(getValue("�ջ���������")))
                receiver.cel = records(row, Int(getValue("�ջ��˵绰��")))
                receiver.add = records(row, Int(getValue("�ջ��˵�ַ��")))
                receiver.id = ""
                receiver.lst_linker = records(row, Int(getValue("������������")))
                
                sender.Name = records(row, Int(getValue("������������")))
                sender.cel = records(row, Int(getValue("�����˵绰��")))
                sender.add = records(row, Int(getValue("�����˵�ַ��")))
                sender.id = records(row, Int(getValue("����������֤����")))
                sender.lst_linker = records(row, Int(getValue("�ջ���������")))
                sender.lst_item = records(row, Int(getValue("����������")))
                sender.lst_pkg = records(row, Int(getValue("��װ��")))
                If (receiver.cel <> "---" Or receiver.add <> destination) Then Call update_people(destination & "�ջ�����Ϣ", receiver)
                Call update_people(destination & "��������Ϣ", sender)
            End If
        Next

        driver.Name = .Cells(3, Int(getValue("��ʻԱ������"))).value
        driver.cel = "---"
        driver.add = .Cells(3, Int(getValue("��ʻԱ������"))).value
        Call update_people(destination & "��ʻԱ��Ϣ", driver)
        .Cells(2, Int(getValue("������"))) = getId

        'Data Base
        Call easyTmpPageUploader(records, .Cells(2, Int(getValue("������"))), .Cells(3, Int(getValue("�嵥������"))), driver.Name, driver.add, destination, .Cells(pageHead_Row + pageSize - 5, Int(getValue("�ӷ���"))), .Cells(pageHead_Row + pageSize - 4, Int(getValue("��ע��"))))

        If destination = sheetName Or Len(destination) <= 0 Then
            .rows("1:45").Insert
            Sheet4.Range(Sheet4.Cells(1, 1), Sheet4.Cells(45, pageWidth)).Copy .Range(.Cells(1, 1), .Cells(45, pageWidth))
            .rows(1).RowHeight = 26.25
            .rows(4).RowHeight = 37.5
            .rows("41:44").RowHeight = 15
            .Cells(1, Val(getValue("�嵥Ŀ�ĵ���"))) = sheetName
            Call addFormular(sheetName, sheetName)
        Else
            If Not sheet_exist(destination) Then
                If Not sheet_exist("�ӵ�") Then
                    Sheets("����").Copy Before:=Sheets(Sheets.count)
                    With ActiveSheet
                        .Name = "�ӵ�"
                        .Unprotect
                        .rows("1:46").Delete
                        .button_input.Enabled = False
                        .button_delete_page.Enabled = False
                        .button_new_page.Enabled = False
                    End With
                End If
                destination = "�ӵ�"
            End If
            Dim insert_row As Integer
            Call editOn(destination)
            With Sheets(destination)
                insert_row = toInt(.Cells(1, toInt(getValue("�嵥������")))) + 1
                .rows(insert_row & ":" & insert_row + pageSize - 1).Insert
                Sheets(sheetName).Range(Sheets(sheetName).Cells(1, 1), Sheets(sheetName).Cells(pageSize, pageWidth)).Copy .Range(.Cells(insert_row, 1), .Cells(insert_row + pageSize - 1, pageWidth))
                .rows(insert_row).RowHeight = 26.25
                .rows(insert_row + 3).RowHeight = 37.5
                .rows(insert_row + pageSize - 5 & ":" & insert_row + pageSize - 2).RowHeight = 15
                .Activate
                .Range(.Cells(insert_row, 1), .Cells(insert_row + pageSize - 1, pageWidth)).Select
            End With
            Call editOff(destination)
            Call editOff(sheetName)
            Application.DisplayAlerts = False
            If temporary_sheet(sheetName) Then Sheets(sheetName).Delete
            Application.DisplayAlerts = True
            Exit Sub
        End If
        
        .Range(.Cells(.Cells(1, Int(getValue("�嵥������"))) + 1, 1), .Cells(.Cells(1, Int(getValue("�嵥������"))) + pageSize, Int(getValue("��ӡ��������")))).Select
    End With
    
    Call editOff(sheetName)
    Call myStatusBar("������ҳ����", 5)
    Exit Sub
error_process: Call editOff(sheetName)
    Application.statusbar = False
    MsgBox ("���󣡣���")
End Sub

Sub editOn(Optional ByVal sheetName As String = "")
    If Len(sheetName) = 0 Then sheetName = ActiveSheet.Name
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ThisWorkbook.Sheets(sheetName).Unprotect Password:="0511/*"
End Sub

Sub editOff(Optional ByVal sheetName As String = "")
    If Len(sheetName) = 0 Then sheetName = ActiveSheet.Name
    ThisWorkbook.Sheets(sheetName).Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ThisWorkbook.Sheets(sheetName).Protect Password:="0511/*", DrawingObjects:=False, Contents:=True, Scenarios:=False
End Sub

Sub delete_page(Optional ByVal sheetName As String = "", Optional ByVal pageHead_Row As Integer = 1, Optional ByVal verify As Boolean = True)
    Dim pageSize As Integer
    If Len(sheetneme) <= 0 Then sheetName = ActiveSheet.Name
    If temporary_sheet Then Exit Sub
    Call editOn(sheetName)
    With Sheets(sheetName)
        pageSize = Val(.Cells(1, Val(getValue("�嵥������"))))
        If pageSize <= 0 Then
            Call errPrint("delete_page 001", "Ŀ���к�С�ڻ�������")
            GoTo subEnd
        End If
        If verify Then
            If (Len(.Cells(1, 1)) >= 0 And (MsgBox("ȷ��ɾ����", vbYesNo) = vbYes)) Then
                .rows(pageHead_Row & ":" & pageSize).Delete
            End If
        Else
            .rows(pageHead_Row & ":" & pageSize).Delete
        End If
    End With
    Call editOff(sheetName)
    Call myStatusBar("�ɹ�ɾ����", 5)
subEnd:
End Sub

Sub export(ByVal pagehead As Integer)
    On Error GoTo error_process
    Dim filepath As String
    Dim filestream As Object
    Dim pageSize As Integer
    Dim pageWidth As Integer
    Dim i As Integer
    Dim str As String
    
    Call myStatusBar("��������...")
    filepath = Application.ActiveWorkbook.PATH & "\" & "[" & ActiveSheet.Name & "]" & Format(Now, "YYYY_MM_DD_HH_MM_SS") & "_" & ActiveSheet.Cells(pagehead + 1, 11).text & "_" & ActiveSheet.Cells(pagehead, 15).value - 10
    Set filestream = CreateObject("ADODB.Stream")
    pageWidth = getValue("�嵥����")
    pageSize = Cells(pagehead, pageWidth - 1)
    
    With filestream
        .Type = 2
        .Charset = "UTF-8"
        .Open

        .WriteText "  ���ţ�" & Cells(pagehead + 1, 11), 1
        .WriteText Cells(pagehead + pageSize - 6, 1), 1
        .WriteText "", 1
        
        For i = pagehead + 4 To pagehead + pageSize - 7
            For Each cel In Range(Cells(i, 1), Cells(i, pageWidth))
                If IsError(cel.value) Then
                    str = str & "[]"
                Else
                    str = str & "[" & cel.value & "]"
                End If
            Next
            .WriteText str, 1
            str = ""
        Next
        .SaveToFile filepath, 1
        .Flush
        .Close
    End With
    Call myStatusBar("�������ļ���" & filepath, 5)
    Exit Sub
error_process: MsgBox ("�ļ����ɴ��󣡣���")
    Application.statusbar = False
End Sub


Sub update_people(ByVal sheetName As String, ByRef data As People)
    If Not sheet_exist(sheetName) Then
        Call errPrint("update_people 201", "������:" & sheetName & "������")
        Exit Sub
    End If
    Dim Record As People
    Dim person As People
    person = data
    If Len(person.Name) <> 0 And (Len(person.cel) <> 0 Or Len(person.add) <> 0 Or Len(person.id) <> 0) Then
        r = getRow(person.Name, 1, sheetName, 3)
        With Sheets(sheetName)
            If r > 0 Then
                Record.cel = .Cells(r, 4).value
                Record.add = .Cells(r, 5).value
                Record.id = .Cells(r, 6).value
                If Record.cel <> person.cel Or Record.add <> person.add Or Record.id <> person.id Then
                    .Cells(r, 1) = person.lst_linker
                    .Range(.Cells(r, 4), .Cells(r, 12)) = Array(person.cel, person.add, person.id, person.lst_item, person.lst_pkg, Now, Record.cel, Record.add, Record.id)
                End If
            Else
                r = .Columns(3).Find("").row
                .Range(.Cells(r, 1), .Cells(r, 8)) = Array(person.lst_linker, "", person.Name, person.cel, person.add, person.id, person.lst_item, person.lst_pkg)
            End If
        End With
    End If
End Sub

Sub new_line(ByVal rows As Integer)
    If ActiveSheet.Name = "����" Then Exit Sub
    If rows < 1 Then Exit Sub
    On Error GoTo error_process
    Dim lstrow As Integer
    Dim listwidth As Integer
    Dim arr() As Integer
    
    listwidth = getValue("�嵥����")
    lstrow = Cells(1, Val(getValue("�嵥������"))) - 6
    PRICECOUNTER_V = True
    ActiveSheet.rows(lstrow & ":" & rows + lstrow - 1).Insert
    PRICECOUNTER_V = False
    Cells(lstrow + rows, 1) = lstrow + rows - 4
    Range("B" & lstrow + rows & ":Q" & lstrow + rows).AutoFill destination:=Range("B" & lstrow & ":Q" & lstrow + rows), Type:=xlFillCopy
    Range("B" & lstrow & ":Q" & lstrow + rows - 1).Locked = False
    ReDim arr(lstrow To lstrow + rows)
    For i = 0 To rows
        arr(lstrow + i) = lstrow - 4 + i
    Next
    Range("A" & lstrow & ":A" & lstrow + rows) = Application.Transpose(arr)
    Cells(1, Val(getValue("�嵥������"))) = lstrow + 6 + rows
    
    Exit Sub
error_process:
    MsgBox ("���󣡣���")
End Sub

Function getRow(ByVal value As String, ByVal row As Integer, ByVal sheetName As String, Optional ByVal col As Integer = 1) As Integer
    Dim target As Range
    With Sheets(sheetName)
        If value = .Cells(row, col).value Then
            getRow = row
            Exit Function
        End If
        Set target = .Columns(col).Find(value, After:=.Cells(row, col), lookat:=xlPart)
    End With
    
    If target Is Nothing Then
        getRow = 0
    Else
        If target.row < row Then
            getRow = 0
        Else
            getRow = target.row
        End If
    End If
End Function

Function PRICECOUNTER(ByVal receiver As String, ByVal address As String, ByVal item As String, ByVal sender As String, ByVal pkg As String, ByVal quantity As Integer, ByVal sheetName As String, Optional ByVal Shift As Integer = 0) As Double
    If sheetName = Left(ActiveSheet.Name, Len(sheetName)) And PRICECOUNTER_V = True Then Application.Volatile
    Dim i As Integer
    Dim listhead As Range
    Dim pl As Integer
    Dim target_row As Integer
    Dim Record As PriceRecord
    If quantity = 0 Then
        PRICECOUNTER = 0
        Exit Function
    End If
    
    With Sheets("�۸�")
        i = 0
        Set listhead = .rows(1).Find(sheetName)
        If listhead Is Nothing Then
            target_row = 0
            GoTo output
        Else
            pl = listhead.Column + (Shift * 6)
        End If
        
        target_row = 4
        Do
            If Len(receiver) = 0 Then Exit Do
            target_row = target_row + 1
            target_row = getRow(receiver, target_row, "�۸�", pl)
            If target_row > 0 Then
                Record.address = .Cells(target_row, pl + 1)
                Record.sender = .Cells(target_row, pl + 2)
                Record.item = .Cells(target_row, pl + 3)
                Record.pkg = .Cells(target_row, pl + 4)
                If (Len(Record.address) = 0 Or Record.address = address) And (Len(Record.item) = 0 Or Record.item = item) And (Len(Record.pkg) = 0 Or Record.pkg = pkg) And (Len(Record.sender) = 0 Or Record.sender = sender) Then
                    Record.price = .Cells(target_row, pl + 5).value
                    GoTo output
                End If
            End If
        Loop While target_row > 0

        target_row = 4
        Do
            If Len(address) = 0 Then Exit Do
            target_row = target_row + 1
            target_row = getRow(address, target_row, "�۸�", pl + 1)
            If target_row > 0 Then
                Record.receiver = .Cells(target_row, pl)
                Record.sender = .Cells(target_row, pl + 2)
                Record.item = .Cells(target_row, pl + 3)
                Record.pkg = .Cells(target_row, pl + 4)
                If Len(Record.receiver) = 0 And (Len(Record.item) = 0 Or Record.item = item) And (Len(Record.pkg) = 0 Or Record.pkg = pkg) And (Len(Record.sender) = 0 Or Record.sender = sender) Then
                    Record.price = .Cells(target_row, pl + 5).value
                    GoTo output
                End If
            End If
        Loop While target_row > 0

        target_row = 4
        Do
            If Len(sender) = 0 Then Exit Do
            target_row = target_row + 1
            target_row = getRow(sender, target_row, "�۸�", pl + 2)
            If target_row > 0 Then
                Record.receiver = .Cells(target_row, pl)
                Record.address = .Cells(target_row, pl + 1)
                Record.item = .Cells(target_row, pl + 3)
                Record.pkg = .Cells(target_row, pl + 4)
                If Len(Record.receiver) = 0 And Len(Record.address) = 0 And (Len(Record.pkg) = 0 Or Record.pkg = pkg) And (Len(Record.item) = 0 Or Record.item = item) Then
                    Record.price = .Cells(target_row, pl + 5).value
                    GoTo output
                End If
            End If
        Loop While target_row > 0
        
        target_row = 4
        Do
            If Len(item) = 0 Then Exit Do
            target_row = target_row + 1
            target_row = getRow(item, target_row, "�۸�", pl + 3)
            If target_row > 0 Then
                Record.receiver = .Cells(target_row, pl)
                Record.address = .Cells(target_row, pl + 1)
                Record.sender = .Cells(target_row, pl + 2)
                Record.pkg = .Cells(target_row, pl + 4)
                If Len(Record.receiver) = 0 And Len(Record.address) = 0 And Len(Record.sender) = 0 And (Len(Record.pkg) = 0 Or Record.pkg = pkg) Then
                    Record.price = .Cells(target_row, pl + 5).value
                    GoTo output
                End If
            End If
        Loop While target_row > 0
        
        target_row = 4
        Do
            If Len(pkg) = 0 Then Exit Do
            target_row = target_row + 1
            target_row = getRow(pkg, target_row, "�۸�", pl + 4)
            If target_row > 0 Then
                Record.receiver = .Cells(target_row, pl)
                Record.address = .Cells(target_row, pl + 1)
                Record.item = .Cells(target_row, pl + 3)
                Record.sender = .Cells(target_row, pl + 2)
                If Len(Record.receiver) = 0 And Len(Record.address) = 0 And Len(Record.item) = 0 And Len(Record.sender) = 0 Then
                    Record.price = .Cells(target_row, pl + 5).value
                    GoTo output
                End If
            End If
        Loop While target_row > 0
    End With
    
output:
    If target_row > 0 Then
        PRICECOUNTER = Round(quantity * Record.price)
    Else
        PRICECOUNTER = 0
    End If
End Function

Function STRBOX(ByVal str As String, ByVal strlen As Integer) As String
    Dim i As Integer
    Dim space As String
    i = strlen - Len(str)
    While i > 0
        space = space & " "
        i = i - 1
    Wend
    STRBOX = str & space
End Function

Function priceListExist(Optional ByVal sheetName As String = "") As Boolean
    If Len(sheetName) = 0 Then sheetName = ActiveSheet.Name
    Dim i As Integer
    Dim listwidth As Integer
    listwidth = getValue("�۸񵥿���")
    i = 0
    With Sheets("�۸�")
        While .Cells(1, i * listwidth + 1) <> sheetName And Len(.Cells(1, i * listwidth + 1)) > 0
            i = i + 1
        Wend
        If Len(.Cells(1, i * listwidth + 1)) <= 0 Then
            priceListExist = False
        Else
            priceListExist = True
        End If
    End With
End Function

Function sheet_exist(ByVal sheetName As String) As Boolean
    Dim i As Integer
    Dim sh
    For Each sh In Worksheets
        If sh.Name = sheetName Then
            sheet_exist = True
            Exit Function
        End If
    Next
    sheet_exist = False
End Function

Function getValue(ByVal valName As String) As String
    If Not setVAL_D Then GoTo funcEnd
    getValue = VAL_D(valName)
    Exit Function
funcEnd:
    Call Err.Raise(vbObjectError + 513, "code.getValue", "")
End Function

Function setVAL_D() As Boolean
    setVAL_D = False
    Dim valSize As Integer
    valSize = Val(Sheets("ֵ").Cells(1, 3))

    If VAL_D Is Nothing Then
        Set VAL_D = CreateObject("Scripting.Dictionary")
        Dim i As Integer
        i = 1
        With Sheets("ֵ")
            .Calculate
            If valSize <= 0 Then
                Call Err.Raise(vbObjectError + 513, "code.setVAL_D", "Unable to read value sheet")
            End If
            arr = .Range("A1:B" & valSize)
            Do
                If Len(arr(i, 1)) > 0 And Len(arr(i, 2)) > 0 Then
                    VAL_D(arr(i, 1)) = arr(i, 2)
                Else
                    Call Err.Raise(vbObjectError + 513, "code.setVAL_D", "Empty value or key on value sheet")
                End If
                i = i + 1
            Loop While i <= valSize
        End With
    End If
    If valSize <> VAL_D.count Then
        Set VAL_D = Nothing
        Call setVAL_D
    End If
    setVAL_D = True
End Function

Function chgValue(ByVal valName As String, ByVal value As String) As Boolean
    chgValue = False
    If setVAL_D And tools.isInt(VAL_D(valName & "row")) Then
        Dim valRow As Integer
        valRow = VAL_D(valName & "row")
        If valRow <= 0 Then GoTo funcEnd
        With Sheets("ֵ")
            Call editOn("ֵ")
            .Cells(valRow, 2) = value
            VAL_D(valName) = value
        End With
    Else
        Dim i As Integer
        'Call errPrint("chgValue 001", "unoptimized item  " & valName)
        With ThisWorkbook.Sheets("ֵ")
            i = getRow(valName, 1, "ֵ")
            If i > 0 Then
                Call editOn("ֵ")
                .Cells(i, 2) = value
                VAL_D(valName) = value
            Else
                Call tools.errPrint("chgValue 001", "looking for unknow val " & valName)
                GoTo funcEnd
            End If
        End With
    End If
    chgValue = True
funcEnd:
    Call editOff("ֵ")
End Function

Sub info_sheet_sort(ByVal sheetName As String)
    On Error GoTo subEnd
    With Sheets(sheetName).Sort
        .SetRange Range("A:F")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
subEnd:
End Sub

Function getId() As String
    Dim id As String
    Dim counter As String
    id = getValue("����ͷ")
    id = id & Format(Date, "#")
    
    If Date <> getValue("�ϴ�ˢ������") Then
        Call chgValue("�ϴ�ˢ������", Date)
        Call chgValue("�嵥������", 1)
        id = id & "001"
    Else
        counter = getValue("�嵥������") + 1
        Call chgValue("�嵥������", counter)
        id = id & Format(counter, "000")
    End If
    getId = id
End Function

Function temporary_sheet(Optional ByVal sheetName As String = "") As Boolean
    If sheetName = "" Then sheetName = ActiveSheet.Name
    With Sheets(sheetName)
        If .Cells(Int(.Cells(1, Int(getValue("�嵥������")))) + 1, 1) = "��ʱ�嵥" Or .Name = "�ӵ�" Then
            temporary_sheet = True
        Else
            temporary_sheet = False
        End If
    End With
End Function

Sub createPriceList(ByVal listName As String)
    If Len(listName) <= 0 Then
        Call errPrint("createPriceList 001", "Empty list name")
        GoTo subEnd
    End If
    With Sheets("�۸�")
        Dim i As Integer
        Dim listwidth As Integer
        listwidth = getValue("�۸񵥿���")
        i = 0
        Call editOn("�۸�")
        While .Cells(1, i * listwidth + 1) <> ""
            i = i + 1
        Wend
        Sheets("�۸�����").Range(Sheets("�۸�����").Columns(1), Sheets("�۸�����").Columns(listwidth)).Copy .Cells(1, i * listwidth + 1)
        .Cells(1, i * listwidth + 1) = listName
        Call editOff("�۸�")
    End With
subEnd:
End Sub

Function isNewMonth() As Boolean
    isNewMonth = False
    If Year(Now) > Year(getValue("�ϴα�������")) Or Month(Now) > Month(getValue("�ϴα�������")) Then
        isNewMonth = True
    End If
End Function

Sub createMonthlyRecord(ByRef PATH As String)
    Workbooks.add
    Dim newWorkbookName As String
    Dim colCounter As Integer
    Dim sheetsCounter As Integer
    Dim specialSheet As Boolean
    newWorkbookName = ActiveWorkbook.Name
    colCounter = 1
    sheetsCounter = 1
    specialSheet = sheet_exist("�ӵ�")
    While Len(ThisWorkbook.Sheets("�۸�").Cells(1, colCounter)) > 0 Or specialSheet
        Dim sheetName As String
        Dim copyRowStart As Integer
        Dim copyRowEnd As Integer
        If specialSheet Then
                sheetName = "�ӵ�"
                specialSheet = False
                colCounter = 1 - Int(getValue("�۸񵥿���"))
        Else
            sheetName = ThisWorkbook.Sheets("�۸�").Cells(1, colCounter)
        End If
        copyRowStart = 1
        With ThisWorkbook.Sheets(sheetName)
        
            Dim pageSize As Integer
            pageSize = .Cells(copyRowStart, Int(getValue("�嵥������")))
            If pageSize > 0 Then
                While Int(.Cells(copyRowStart, Int(getValue("�嵥������")))) > 0 And .Cells(copyRowStart + 2, Int(getValue("�嵥������"))) >= DateSerial(Year(Date), Month(Date), 1)
                    copyRowStart = copyRowStart + Int(.Cells(copyRowStart, Int(getValue("�嵥������"))))
                Wend
                copyRowEnd = copyRowStart
                While Int(.Cells(copyRowEnd, Int(getValue("�嵥������")))) > 0 And .Cells(copyRowEnd + 2, Int(getValue("�嵥������"))) >= DateSerial(Year(Date), Month(Date) - 1, 1)
                    copyRowEnd = copyRowEnd + Int(.Cells(copyRowEnd, Int(getValue("�嵥������"))))
                Wend
                copyRowEnd = copyRowEnd - 1
            End If
            
            If Workbooks(newWorkbookName).Sheets.count < sheetsCounter Then Workbooks(newWorkbookName).Sheets.add After:=Workbooks(newWorkbookName).Sheets(Workbooks(newWorkbookName).Sheets.count)
            If copyRowEnd - copyRowStart > 1 Then
                Dim counter As Integer
                counter = Int(getValue("�嵥����"))
                .rows(copyRowStart & ":" & copyRowEnd).Copy
                Workbooks(newWorkbookName).Sheets(sheetsCounter).Paste
                While counter > 0
                    Workbooks(newWorkbookName).Sheets(sheetsCounter).Columns(counter).ColumnWidth = .Columns(counter).ColumnWidth
                    counter = counter - 1
                Wend
            End If
            
        End With
        Workbooks(newWorkbookName).Sheets(sheetsCounter).Name = sheetName
        sheetsCounter = sheetsCounter + 1
        colCounter = colCounter + Int(getValue("�۸񵥿���"))
    Wend
    Workbooks(newWorkbookName).SaveAs fileName:=PATH & "\(" & MonthName(Month(Date) - 1) & ")" & "�����˵�.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks("(" & MonthName(Month(Date) - 1) & ")" & "�����˵�.xlsx").Close
End Sub

Sub deleteExpiredData(ByRef deletedSheetCounter As Integer)
    Dim specialSheet As Boolean
    Dim colCounter As Integer
    colCounter = 1
    specialSheet = sheet_exist("�ӵ�")
    deletedSheetCounter = 0
    While Len(ThisWorkbook.Sheets("�۸�").Cells(1, colCounter)) > 0 Or specialSheet
        If specialSheet Then
                sheetName = "�ӵ�"
                specialSheet = False
                colCounter = 1 - Int(getValue("�۸񵥿���"))
        Else
            sheetName = ThisWorkbook.Sheets("�۸�").Cells(1, colCounter)
        End If
        With ThisWorkbook.Sheets(sheetName)
        
            Dim delRowStart As Integer
            Dim delRowEnd As Integer
            Dim pageSize As Integer
            delRowStart = 1
            pageSize = .Cells(delRowStart, Int(getValue("�嵥������")))
            
            While Int(.Cells(delRowStart, Int(getValue("�嵥������")))) > 0 And .Cells(delRowStart + 2, Int(getValue("�嵥������"))) >= DateSerial(Year(Date), Month(Date) - Int(getValue("�嵥��������")), 1)
                delRowStart = delRowStart + Int(.Cells(delRowStart, Int(getValue("�嵥������"))))
            Wend
            delRowEnd = delRowStart
            While Int(.Cells(delRowEnd, Int(getValue("�嵥������")))) > 0
                delRowEnd = delRowEnd + Int(.Cells(delRowEnd, Int(getValue("�嵥������"))))
                deletedSheetCounter = deletedSheetCounter + 1
            Wend
            delRowEnd = delRowEnd - 1
            
            If delRowEnd - delRowStart > 0 Then
                'Debug.Print (sheetName & " : " & deletedSheetCounter & " : " & delRowStart & " : " & delRowEnd)
                Call editOn(sheetName)
                .rows(delRowStart & ":" & delRowEnd).Delete
                Call editOff(sheetName)
            End If
            
        End With
        sheetsCounter = sheetsCounter + 1
        colCounter = colCounter + Int(getValue("�۸񵥿���"))
    Wend
    Application.DisplayAlerts = False
    If sheet_exist("�ӵ�") Then If Len(ThisWorkbook.Sheets("�ӵ�").Cells(1, Int(getValue("�嵥������")))) = 0 Then ThisWorkbook.Sheets("�ӵ�").Delete
    Application.DisplayAlerts = True
End Sub


Sub addFormular(ByVal sheetName As String, ByVal destination As String)
With Sheets(sheetName)
    .Cells(3, Val(getValue("��ʻԱ������"))).Formula = Replace(getValue("��ʻԱ���ƹ�ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("�ջ��˵绰��"))).Formula = Replace(getValue("�ջ��˵绰��ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("�ջ��˵绰��"))).AutoFill destination:=.Range(.Cells(5, Val(getValue("�ջ��˵绰��"))), .Cells(39, Val(getValue("�ջ��˵绰��")))), Type:=xlFillValues
    .Cells(5, Val(getValue("�ջ��˵�ַ��"))).Formula = Replace(getValue("�ջ��˵�ַ��ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("�ջ��˵�ַ��"))).AutoFill destination:=.Range(.Cells(5, Val(getValue("�ջ��˵�ַ��"))), .Cells(39, Val(getValue("�ջ��˵�ַ��")))), Type:=xlFillValues
    
    .Cells(5, Val(getValue("����������"))).Formula = Replace(getValue("�������ƹ�ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("����������"))).AutoFill destination:=.Range(.Cells(5, Val(getValue("����������"))), .Cells(39, Val(getValue("����������")))), Type:=xlFillValues
    .Cells(5, Val(getValue("��װ��"))).Formula = Replace(getValue("��װ��ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("��װ��"))).AutoFill destination:=.Range(.Cells(5, Val(getValue("��װ��"))), .Cells(39, Val(getValue("��װ��")))), Type:=xlFillValues
    
    .Cells(5, Val(getValue("������������"))).Formula = Replace(getValue("������������ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("������������"))).AutoFill destination:=.Range(.Cells(5, Val(getValue("������������"))), .Cells(39, Val(getValue("������������")))), Type:=xlFillValues
    .Cells(5, Val(getValue("�����˵绰��"))).Formula = Replace(getValue("�����˵绰��ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("�����˵绰��"))).AutoFill destination:=.Range(.Cells(5, Val(getValue("�����˵绰��"))), .Cells(39, Val(getValue("�����˵绰��")))), Type:=xlFillValues
    .Cells(5, Val(getValue("�����˵�ַ��"))).Formula = Replace(getValue("�����˵�ַ��ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("�����˵�ַ��"))).AutoFill destination:=.Range(.Cells(5, Val(getValue("�����˵�ַ��"))), .Cells(39, Val(getValue("�����˵�ַ��")))), Type:=xlFillValues
    .Cells(5, Val(getValue("����������֤����"))).Formula = Replace(getValue("����������֤�Ź�ʽ"), "__DESTINATION__", destination)
    .Cells(5, Val(getValue("����������֤����"))).AutoFill destination:=.Range(.Cells(5, Val(getValue("����������֤����"))), .Cells(39, Val(getValue("����������֤����")))), Type:=xlFillValues
End With
End Sub

Sub updateIdList(ByRef id_dict As Object)
    Dim idRow As Long
    Dim id As String
    Dim maxId As Long
    maxId = 31
    
    If Sheets("�۸�").rows("1").Find(ActiveSheet.Name) Is Nothing Then
        Call errPrint("000", "updateIdList")
        Exit Sub
    End If
    
    If Not id_dict Is Nothing Then id_dict.RemoveAll
    Set id_dict = CreateObject("Scripting.Dictionary")
    If Len(Cells(2, Int(getValue("������")))) = 0 Then
        idRow = 2 + Val(Cells(1, Int(getValue("�嵥������"))))
    Else
        idRow = 2
    End If
    
    Do While maxId > 0
        id = Cells(idRow, Int(getValue("������")))
        If Len(id) > 0 Then
            If Cells(idRow - 1, 1) = getValue("�嵥ͷ") Then id_dict(Cells(idRow, Int(getValue("������"))).text) = idRow - 1
            idRow = idRow + Val(Cells(idRow - 1, Int(getValue("�嵥������")))) 'format
        Else
            Exit Do
        End If
        maxId = maxId - 1
    Loop
End Sub

Sub hidePage(Optional ByVal sheetName As String = "", Optional ByVal pageHead_Row As Integer = 1)
    Dim pageSize As Integer
    If Len(sheetneme) <= 0 Then sheetName = ActiveSheet.Name
    If temporary_sheet Or pageHead_Row = 1 Then Exit Sub
    Call editOn(sheetName)
    With Sheets(sheetName)
        pageSize = Val(.Cells(pageHead_Row, Val(getValue("�嵥������"))))
        If pageSize <= 0 Then
            Call errPrint("delete_page 001", "Ŀ���к�С�ڻ�������")
            GoTo subEnd
        End If
        
        With .Range(.Cells(pageHead_Row, 1), .Cells(pageHead_Row + pageSize - 1, Int(getValue("�嵥����")))).Interior
            .Pattern = xlLightDown
            .PatternColor = 255
            .ColorIndex = 2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        .Cells(pageHead_Row, 1) = "ɾ���ڣ�" & Now
        .rows(pageHead_Row & ":" & pageHead_Row + pageSize - 1).Hidden = True
        
        Call easyTmpPageDeleter(.Cells(pageHead_Row + 1, Int(getValue("������"))))
    End With
    Call editOff(sheetName)
    Call myStatusBar("�ɹ�ɾ����", 5)
subEnd:
    Call editOff(sheetName)
End Sub

Sub easyTmpPageDeleter(ByVal id As String)
    On Error GoTo errorProcess
    Dim db As New DataBase
    Call db.connect("kangtai")
    Call db.openTrans
    Call db.deleteDataById("tmp_detailed_record", id)
    Call db.deleteDataById("tmp_general_record", id)
    Call db.commit
    Call db.disconnect
    Exit Sub
errorProcess:
    Call errPrint("easyTmpPageDeleter", "delete failed")
End Sub

Sub easyTmpPageUploader(ByRef datas As Variant, ByRef pageId As String, ByRef pageDate As String, ByRef driverName As String, ByRef driverCarNumber As String, ByRef destination As String, ByRef cost As Currency, ByRef note As String)
    On Error GoTo errorProcess
    Dim db As New DataBase
    Dim records() As Record
    records = db.toRecordSet(datas, pageId, pageDate, driverName, driverCarNumber, destination)
    Call db.connect("kangtai")
    Call db.openTrans
    
    
    For row = 1 To UBound(records)
        Call db.uploadRecord("tmp_detailed_record", records(row))
    Next
    
    Call db.uploadPageSummery("tmp_general_record", db.sumRecords(records, pageId, pageDate, driverName, driverCarNumber, destination, cost, note))
    
    
    Call db.commit
    Call db.disconnect
    Exit Sub
errorProcess:
    Call errPrint("easyTmpPageUploader", "upload failed")
End Sub

