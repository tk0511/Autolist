Attribute VB_Name = "code"
Function version()
    version = "1.0.0"
End Function

Function getConnection()
    Set connection = New ADODB.connection
    connection.ConnectionString = ThisWorkbook.Sheets("values").Cells(1, 2).text
    connection.Open
    Set getConnection = connection
End Function

Sub uploadPage()
    On Error GoTo errorProcess
    rb = False
    id = ThisWorkbook.Sheets("运单").Cells(2, 14).text
    Set connection = getConnection()
    Set rs = connection.Execute("SELECT verified FROM tmp_general_record WHERE id = """ & id & """")
    If rs.EOF() Then GoTo errorProcess
    connection.BeginTrans
    rb = True
    
    If Not CBool(rs.Fields(0)) Then
        Call connection.Execute("UPDATE `tmp_general_record` SET verified = 1 WHERE id = """ & id & """")
        Call connection.Execute("INSERT INTO general_record SELECT * FROM tmp_general_record WHERE id = """ & id & """")
        Call connection.Execute("INSERT INTO detailed_record SELECT * FROM tmp_detailed_record WHERE id = """ & id & """")
    End If
    
    r = 4
    With ThisWorkbook.Sheets("运单")
        freightAtDestination = 0
        freightAtBase = 0
        freightAtBaseUnpaid = 0
        freightAtDestinationUnpaid = 0
        unloadFee = 0
        totalQty = 0
        
        While Len(.Cells(r, 4)) > 0
            If .Cells(r, 9) = "外付" Then
                freightAtDestination = freightAtDestination + .Cells(r, 5).Value
            ElseIf .Cells(r, 9) = "内付" Then
                freightAtBase = freightAtBase + .Cells(r, 5).Value
            ElseIf .Cells(r, 9) = "内欠" Then
                freightAtBaseUnpaid = freightAtBaseUnpaid + .Cells(r, 5).Value
            ElseIf .Cells(r, 9) = "外欠" Then
                freightAtDestinationUnpaid = freightAtDestinationUnpaid + .Cells(r, 5).Value
            Else
                GoTo errorProcess
            End If
            unloadFee = unloadFee + .Cells(r, 6).Value
            transferFee = transferFee + .Cells(r, 7).Value
            totalQty = totalQty + .Cells(r, 4).Value
            
            connection.Execute ("UPDATE `detailed_record` SET `item` = """ & .Cells(r, 2) & """, `pkg` = """ & .Cells(r, 3) & """, `qty` = " & .Cells(r, 4).Value & ", `freight` = " & .Cells(r, 5).Value & ", `unloadingFee` = " & .Cells(r, 6).Value & ", `transferFee` = " & .Cells(r, 7).Value & ", `sum` = " & .Cells(r, 8).Value & ", `payment` = """ & .Cells(r, 9) & """, `comment` = """ & .Cells(r, 10) & """, `receverName` = """ & .Cells(r, 11) & """, `receverTel` = """ & .Cells(r, 12) & """, `senderName` = """ & .Cells(r, 13) & """, `senderTel` = """ & .Cells(r, 14) & """ WHERE id = """ & id & """ AND count = " & .Cells(r, 1))
            r = r + 1
        Wend
        
        cost = .Cells(2, 12).Value
        extraCost = .Cells(2, 9).Value
        totalFreight = freightAtDestination + freightAtBase + freightAtBaseUnpaid + freightAtDestinationUnpaid
        payAtDestination = freightAtDestination - unloadFee + transferFee
        profit = totalFreight - cost - extraCost
        
        connection.Execute ("UPDATE `general_record` SET `extraCostDesc`=""" & .Cells(2, 7).Value & """,`freightAtDestination`=" & freightAtDestination & ",`freightAtBase`=" & freightAtBase & ",`freightAtBaseUnpaid`=" & freightAtBaseUnpaid & ",`totalFreight`=" & totalFreight & ",`cost`=" & cost & ",`profit`=" & profit & ",`freightAtDestinationUnpaid`=" & freightAtDestinationUnpaid & ",`unloadFee`=" & unloadFee & ",`transferFee`=" & transferFee & ",`payAtDestination`=" & payAtDestination & ",`totalQty`=" & totalQty & ",`extraCost`=" & extraCost & " WHERE `id` = """ & id & """")
    End With
    
    connection.CommitTrans
    Call editOn("运单")
    ThisWorkbook.Sheets("运单").Cells(2, 14).Interior.PatternColor = 5287936
    Call editOff("运单")
    Exit Sub
errorProcess:
    If rb Then connection.RollbackTrans
    MsgBox "上传失败!"
End Sub

Sub editOn(Optional ByVal sheetName As String = "")
    If Len(sheetName) = 0 Then sheetName = ActiveSheet.name
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ThisWorkbook.Sheets(sheetName).Unprotect Password:=ThisWorkbook.Sheets("values").Cells(2, 2).text
End Sub

Sub editOff(Optional ByVal sheetName As String = "")
    If Len(sheetName) = 0 Then sheetName = ActiveSheet.name
    ThisWorkbook.Sheets(sheetName).Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ThisWorkbook.Sheets(sheetName).Protect Password:=ThisWorkbook.Sheets("values").Cells(2, 2).text, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
End Sub

Sub clearPage()
    Call editOn("运单")
    With ThisWorkbook.Sheets("运单")
        .Range("B4:N303").ClearContents
        .Cells(1, 1).Formula = ThisWorkbook.Sheets("values").Cells(3, 2).text
        .Cells(2, 9) = 0
        .Cells(2, 12) = 0
    End With
    Call editOff("运单")
End Sub

Function downloadPage(ByRef id As String) As Boolean
    On Error GoTo errorProcess
    downloadPage = False
    
    Set connection = getConnection()
    Set rs = connection.Execute("SELECT id, destination, pageDate, driverName, driverCarNumber, note, cost, extraCost, verified, extraCostDesc FROM tmp_general_record WHERE id = """ & id & """")
    If rs.EOF() Then
        GoTo errorProcess
    Else
        Call clearPage
        Call editOn("运单")
        With ThisWorkbook.Sheets("运单")
            If CBool(rs.Fields(8)) Then
                targetTable = "detailed_record"
                Set rs = connection.Execute("SELECT id, destination, pageDate, driverName, driverCarNumber, note, cost, extraCost, verified, extraCostDesc FROM general_record WHERE `id` = """ & id & """")
            Else
                targetTable = "tmp_detailed_record"
            End If
            .Cells(1, 1).Formula = ThisWorkbook.Sheets("values").Cells(3, 2).text & "& """ & rs.Fields(5) & """"
            .Cells(2, 1) = "[" & rs.Fields(1) & "] " & rs.Fields(2) & " - " & rs.Fields(3) & " " & rs.Fields(4)
            .Cells(2, 9) = rs.Fields(7)
            .Cells(2, 12) = rs.Fields(6)
            .Cells(2, 7) = rs.Fields(9)
            .Cells(2, 14) = id
            
            
            Set rs = connection.Execute("SELECT `item`, `pkg`, `qty`, `freight`, `unloadingFee`, `transferFee`, `sum`, `payment`, `comment`, `receverName`, `receverTel`, `senderName`, `senderTel` FROM `" & targetTable & "` WHERE id = " & id & " ORDER BY count")
            r = 4
            While Not rs.EOF()
                If r > 303 Then
                    MsgBox "运单过长，载入失败！"
                    GoTo errorProcess
                End If
                
                For c = 2 To rs.Fields.Count + 1 Step 1
                    If c <> 8 Then
                        .Cells(r, c) = rs.Fields(c - 2)
                    Else
                        .Cells(r, c).Formula = "=E" & r & "-F" & r & "+G" & r
                    End If
                Next
                r = r + 1
                rs.MoveNext
            Wend
        End With
        Call editOff("运单")
    End If
    downloadPage = True
    Exit Function
errorProcess:
    Call clearPage
    Call editOff("运单")
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

Sub showUnverified()
    Set connection = getConnection()
    Set rs = connection.Execute("SELECT destination, pageDate, id FROM tmp_general_record WHERE verified = 0 limit 15")
    c = 0
    msg = ""
    While Not rs.EOF()
      msg = msg & "[" & rs.Fields(0) & "] " & rs.Fields(1) & " : " & rs.Fields(2) & Chr(10)
      c = c + 1
      rs.MoveNext
    Wend
    If c = 15 Then msg = msg + "....."
    MsgBox msg
End Sub

Sub printSummery(ByRef beging As Date, ByRef ending As Date)
    On Error GoTo errorProcess
    Set connection = getConnection()
    Set result = CreateObject("Scripting.Dictionary")
    
    Set rs = connection.Execute("SELECT `qty`, `sum`, `payment`, `receverName`, `receverTel`, `senderName`, `senderTel` FROM `detailed_record` WHERE date >= """ & beging & """ and date < """ & ending & """")
    
    While Not rs.EOF()
        RID = rs.Fields(3) & "|" & rs.Fields(4)
        SID = rs.Fields(5) & "|" & rs.Fields(6)
        If IsEmpty(result(RID)) Then
            Set result(RID) = New People
            result(RID).name = rs.Fields(3)
            result(RID).tel = rs.Fields(4)
            result(RID).role = "收货人"
        End If
    
        If IsEmpty(result(SID)) Then
            Set result(SID) = New People
            result(SID).name = "" & rs.Fields(5)
            result(SID).tel = "" & rs.Fields(6)
            result(SID).role = "发货人"
        End If
        
        payment = rs.Fields(2).Value
        If payment = "外付" Then
            result(RID).freightAtDestination = result(RID).freightAtDestination + rs.Fields(1).Value
            result(SID).freightAtDestination = result(SID).freightAtDestination + rs.Fields(1).Value
            result(RID).qtyAtDestination = result(RID).qtyAtDestination + rs.Fields(0).Value
            result(SID).qtyAtDestination = result(SID).qtyAtDestination + rs.Fields(0).Value
        ElseIf payment = "内付" Then
            result(RID).freightAtBase = result(RID).freightAtBase + rs.Fields(1).Value
            result(SID).freightAtBase = result(SID).freightAtBase + rs.Fields(1).Value
            result(RID).qtyAtBase = result(RID).qtyAtBase + rs.Fields(0).Value
            result(SID).qtyAtBase = result(SID).qtyAtBase + rs.Fields(0).Value
        ElseIf payment = "内欠" Then
            result(SID).freightAtBaseUnpaid = result(SID).freightAtBaseUnpaid + rs.Fields(1).Value
            result(SID).qtyAtBaseUnpaid = result(SID).qtyAtBaseUnpaid + rs.Fields(0).Value
        ElseIf payment = "外欠" Then
            result(RID).freightAtDestinationUnpaid = result(RID).freightAtDestinationUnpaid + rs.Fields(1).Value
            result(RID).qtyAtDestinationUnpaid = result(RID).qtyAtDestinationUnpaid + rs.Fields(0).Value
        Else
            GoTo errorProcess
        End If
        rs.MoveNext
    Wend
    
    r = 4
    Call clearSummeryPage
    With ThisWorkbook.Sheets("统计")
        For Each p In result.keys
            .Range("A" & r & ":K" & r) = Array(result(p).role, result(p).name, result(p).tel, result(p).qtyAtBaseUnpaid, result(p).freightAtBaseUnpaid, result(p).qtyAtDestinationUnpaid, result(p).freightAtDestinationUnpaid, result(p).qtyAtBase, result(p).freightAtBase, result(p).qtyAtDestination, result(p).freightAtDestination)
            r = r + 1
        Next
    End With
    Call sortSummeryPage
    Exit Sub
errorProcess:
    Call clearSummeryPage
    MsgBox "统计失败!"
End Sub

Sub clearSummeryPage()
    r = 4
    With ThisWorkbook.Sheets("统计")
        While Len(.Cells(r, 1)) > 0
            r = r + 1
        Wend
        .Range("A4:K" & r).ClearContents
    End With
End Sub

Sub sortSummeryPage()
    ThisWorkbook.Worksheets("统计").AutoFilter.sort.SortFields.Clear
    ThisWorkbook.Worksheets("统计").AutoFilter.sort.SortFields.Add Key:=Range("E3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("统计").AutoFilter.sort.SortFields.Add Key:=Range("G3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("统计").AutoFilter.sort.SortFields.Add Key:=Range("D3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("统计").AutoFilter.sort.SortFields.Add Key:=Range("F3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ThisWorkbook.Worksheets("统计").AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub test()
    Debug.Print DateSerial(2017, 6, 1) + 0.25
End Sub
