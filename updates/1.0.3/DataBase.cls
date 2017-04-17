VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private connection

Public Sub connect()
    Set connection = New ADODB.connection
    connection.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" + getValue("DBADD") + "; DATABASE=" + getValue("DB") + "; UID=" + getValue("DBID") + "; PWD=" + getValue("DBPW") + "; OPTION=3"
    connection.Open
End Sub

Public Sub disconnect()
    If connection.state <> 0 Then
        connection.Close
    End If
End Sub

Public Sub openTrans()
    connection.BeginTrans
End Sub

Public Sub commit()
    connection.CommitTrans
End Sub

Public Sub rollback()
    connection.RollbackTrans
End Sub

Public Sub uploadPageSummery(ByRef tableName As String, ByRef r As PageSummery)
    connection.Execute "INSERT INTO `" & tableName & "`(`uploadTime`, `pageDate`, `id`, `destination`, `driverName`, `driverCarNumber`, `freightAtDestination`, `freightAtBase`, `freightAtBaseUnpaid`, `totalFreight`, `cost`, `profit`, `freightAtDestinationUnpaid`, `note`, `unloadFee`, `transferFee`, `payAtDestination`, `totalQty`, `extraCost`, `extraCostDesc`) VALUES (""" & r.uploadTime & """, """ & r.pageDate & """, """ & r.id & """ ,""" & r.destination & """,""" & r.driverName & """,""" & r.driverCarNumber & """,""" & r.freightAtDestination & """,""" & r.freightAtBase & """,""" & r.freightAtBaseUnpaid & """,""" & r.totalFreight & """,""" & r.cost & """,""" & r.profit & """,""" & r.freightAtDestinationUnpaid & """,""" & r.note & """,""" & r.unloadFee & """,""" & r.transferFee & """,""" & r.payAtDestination & """,""" & r.totalQty & """,""" & r.extraCost & """,""" & r.extraCostDesc & """)"
End Sub

Public Sub uploadRecord(ByRef tableName As String, ByRef r)
    connection.Execute "INSERT INTO `" & tableName & "`(`collectdDate`, `date`, `id`, `count`, `item`, `pkg`, `qty`, `freight`, `unloadingFee`, `transferFee`, `sum`, `payment`, `comment`, `receverName`, `receverTel`, `receverAdd`, `senderName`, `senderTel`, `senderAdd`, `senderId`, `driverName`, `driverCarNumber`, `destination`) VALUES (""" & r.collectdDate & """,""" & r.pageDate & """,""" & r.id & """," & r.count & ",""" & r.item & """,""" & r.pkg & """," & r.qty & "," & r.freight & "," & r.unloadingFee & "," & r.transferFee & "," & r.sum & ",""" & r.payment & """,""" & r.comment & """,""" & r.receverName & """,""" & r.receverTel & """,""" & r.receverAdd & """,""" & r.senderName & """,""" & r.senderTel & """,""" & r.senderAdd & """,""" & r.senderId & """,""" & r.driverName & """,""" & r.driverCarNumber & """,""" & r.destination & """)"
End Sub

Sub deleteDataById(ByRef tableName As String, ByRef id As String)
    connection.Execute "DELETE FROM `" & tableName & "` WHERE `id`=""" & id & """;"
End Sub

Function errors()
    Set errors = connection.errors
End Function

Function state()
    state = connection.state
End Function

Function toRecordSet(ByRef datas As Variant, ByVal pageId As String, ByVal pageDate As String, ByVal driverName As String, ByVal driverCarNumber As String, ByVal destination As String) As Record()
    Dim records() As Record
    Dim i As Integer
    ReDim records(1 To UBound(datas))
    For row = 1 To UBound(records)
        Set records(row) = New Record
        records(row).id = pageId
        records(row).count = datas(row, 1)
        records(row).receverName = datas(row, 2)
        records(row).receverTel = datas(row, 3)
        records(row).receverAdd = datas(row, 4)
        records(row).item = datas(row, 5)
        records(row).pkg = datas(row, 6)
        records(row).qty = datas(row, 7)
        records(row).freight = datas(row, 8)
        records(row).unloadingFee = datas(row, 9)
        records(row).transferFee = datas(row, 10)
        records(row).sum = datas(row, 11)
        records(row).payment = datas(row, 12)
        records(row).comment = datas(row, 13)
        records(row).senderName = datas(row, 14)
        records(row).senderTel = datas(row, 15)
        records(row).senderAdd = datas(row, 16)
        records(row).senderId = datas(row, 17)
        records(row).collectdDate = datas(row, 18)
        records(row).pageDate = pageDate
        records(row).driverName = driverName
        records(row).driverCarNumber = driverCarNumber
        records(row).destination = destination
    Next
    For Each r In records
        If r.qty > 0 Then i = i + 1
    Next
    If i = 0 Then
        ReDim records(0)
    Else
        ReDim Preserve records(1 To i)
    End If
    toRecordSet = records
End Function

Function sumRecords(ByRef records As Variant, ByRef pageId As String, ByRef pageDate As String, ByRef driverName As String, ByRef driverCarNumber As String, ByRef destination As String, ByRef cost As Currency, ByRef note As String, ByRef uploadTime As String, ByRef extraCost As Currency, ByRef extraCostDesc As String) As PageSummery
    Dim result As New PageSummery
    result.pageDate = pageDate
    result.id = pageId
    result.destination = destination
    result.driverName = driverName
    result.driverCarNumber = drivercarname
    result.cost = cost
    result.extraCost = extraCost
    result.extraCostDesc = extraCostDesc
    result.note = note
    
    If UBound(records) > 0 Then
        For Each r In records
            If r.payment = "外付" Then
                result.freightAtDestination = result.freightAtDestination + r.freight
            ElseIf r.payment = "内付" Then
                result.freightAtBase = result.freightAtBase + r.freight
            ElseIf r.payment = "内欠" Then
                result.freightAtBaseUnpaid = result.freightAtBaseUnpaid + r.freight
            ElseIf r.payment = "外欠" Then
                result.freightAtDestinationUnpaid = result.freightAtDestinationUnpaid + r.freight
            Else
                Call MsgBox("第 " & r.count & " 条上传统计数据失败！未知付款方式： " & r.payment)
            End If
            result.unloadFee = result.unloadFee + r.unloadingFee
            result.transferFee = result.transferFee + r.transferFee
            result.totalFreight = result.totalFreight + r.freight
            result.totalQty = result.totalQty + r.qty
            result.payAtDestination = result.payAtDestination + r.sum
        Next
    End If
    
    result.profit = result.totalFreight - cost
    result.uploadTime = uploadTime
    Set sumRecords = result
End Function

Private Sub Class_Initialize()
    Set connection = New ADODB.connection
End Sub

Private Sub Class_Terminate()
    If connection.state <> 0 Then Call disconnect
End Sub


