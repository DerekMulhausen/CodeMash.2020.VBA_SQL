Public Sub Worksheet_Activate()
    PopCustNames
    PopItemNames
    btnWriteToDb.Enabled = False
    ProcCustCode
End Sub
Private Sub PopCustNames()
    Dim CustArr As Variant
    Dim Qry As String
    Dim i As Variant
    Qry = "Select CustomerName from Customers Order By CustomerName"
    CustArr = SQLProc.GetRecSet(Qry)
    cboCustomerName.Clear
    For Each i In CustArr
        cboCustomerName.AddItem (i)
    Next i
End Sub
Private Sub cboCustomerName_Change()
    ProcCustCode
End Sub
Private Function GetCustomerID() As Integer
    Dim Qry As String
    Qry = "Select CustomerID from Customers where CustomerName='" & cboCustomerName.Value & "'"
    GetCustomerID = SQLProc.GetRecSet(Qry)(0, 0)
End Function

Private Sub PopItemNames()
    Dim ItemArr As Variant
    Dim Qry As String
    Dim i As Variant
    Qry = "Select ItemName from Items Order By ItemName"
    ItemArr = SQLProc.GetRecSet(Qry)
    cboItemName.Clear
    For Each i In ItemArr
        cboItemName.AddItem (i)
    Next i
End Sub

Private Sub cboItemName_Change()
    ProcCustCode
End Sub

Private Function GetItemID() As Integer
    Dim Qry As String
    Qry = "Select ItemID from Items where ItemName='" & cboItemName.Value & "'"
    GetItemID = SQLProc.GetRecSet(Qry)(0, 0)
End Function

Private Sub ProcCustCode()
    Dim SelQry As String
    If ((Len(cboCustomerName.Value) > 0) And (Len(cboItemName.Value) > 0)) Then
        'Check if code exists?
        Dim ItemID As Integer
        Dim CustomerID As Integer
        ItemID = GetItemID
        CustomerID = GetCustomerID
        SelQry = "Select CustomerCodeID From CustomerCodes where ItemID=" & ItemID & " And CustomerID=" & CustomerID
        If (SQLProc.GetRecSet(SelQry)(0, 0) > 0) Then
            PopCustCode SelQry, ItemID, CustomerID
        Else
            If (MsgBox("CustomerCode does not exist. Would you like to create one?", vbYesNo) = vbYes) Then
                Dim InsQry As String
                InsQry = "Insert into CustomerCodes(CustomerID, ItemID, CustomerCode) Values(" & CustomerID & ", " & ItemID & ", '"
                InsQry = InsQry & InputBox("Please enter " & cboCustomerName.Value & "'s code for " & cboItemName.Value & " if known") & "')" 'Result String is added to InsQry?
                SQLProc.ProcQry (InsQry)
                PopCustCode SelQry, ItemID, CustomerID
            End If
        End If
    Else
        Range("R11").Value = 0
        Range("R10").Value = ""
        btnWriteToDb.Enabled = False
    End If
End Sub

Private Sub PopCustCode(SelQry As String, ItemID As Integer, CustomerID As Integer)
    Range("R11").Value = SQLProc.GetStr(SelQry)
    SelQry = "Select CustomerCode From CustomerCodes where ItemID=" & ItemID & " And CustomerID=" & CustomerID
    Range("R10").Value = GetStr(SelQry)
    btnWriteToDb.Enabled = True
End Sub

Private Sub BtnWriteToDb_Click()
    'get max qty
    Dim Row As Integer
    Dim maxRow As Integer
    If (Len(Range("b5").Text) > 0) Then
        maxRow = Range("b4").End(xlDown).Row
    Else
        MsgBox ("No Data is detected to insert.  Please verify a quantity is present in Cell B5 of PriceList")
        Exit Sub
    End If
 'build query
    Dim Qry As String
    Qry = "Insert Into CustomerCodeDetails (CustomerCodeID, QtyPriced, UnitPrice, StartDate, FinishDate, DiscountPct) Values "
    Row = 5
    Do While (Row <= maxRow)
        Qry = Qry & "(" & Range("r11").Text & ", "
        Qry = Qry & Cells(Row, 2).Text & ", "
        If Len(Cells(Row, 3).Text) > 0 Then
            Qry = Qry & Cells(Row, 3).Text & ", "
        Else
            Qry = Qry & "NULL, "
        End If
        Qry = Qry & "'" & Cells(Row, 4).Text & "', "
        Qry = Qry & "'" & Cells(Row, 5).Text & "', "
        If Cells(Row, 6).Value > 0 Then
            Qry = Qry & Cells(Row, 6).Text & ")"
        Else
            Qry = Qry & "NULL)"
        End If
        If (Row < maxRow) Then
            Qry = Qry & ","
        End If
        Row = Row + 1
    Loop
     'process query
    SQLProc.ProcQry (Qry)
    'clear data from new price breaks area
    Range("b5:f30").ClearContents
    'update current price breaks
    UpdateData True
End Sub

Private Sub btnRefresh_Click()
    UpdateData (True)
End Sub

Private Sub btnUpdate_Click()
    UpdateData (False)
End Sub

Private Sub UpdateData(VerifyUpdate As Boolean)
    Dim Results As Variant
    If (VerifyUpdate = True) Then
        CompareAndUpdate ThisWorkbook.Worksheets("Original Data"), ThisWorkbook.Worksheets("PriceList")
    End If
'query for current data
    Dim Qry As String
    Qry = "Use Contacts SELECT CustomerCodeDetailsID, QtyPriced, UnitPrice, StartDate, FinishDate, DiscountPct FROM CustomerCodeDetails WHERE CustomerCodeID = " & Range("R11").Text
'write qry data to pricelist and original data
    ThisWorkbook.Connections("OriginalData").ODBCConnection.CommandText = Qry
    ThisWorkbook.Connections("PriceListData").ODBCConnection.CommandText = Qry
    On Error Resume Next
    ThisWorkbook.Connections("OriginalData").Refresh
    ThisWorkbook.Connections("PriceListData").Refresh
    Dim bottomID As Integer
    Dim bottomDel As Integer
    If (Len(Range("h5").Value) = 0) Then
        bottomID = 4
    Else
        bottomID = Range("h4").End(xlDown).Row
    End If
    bottomDel = Range("n3").Value
    Dim i As Integer
    If (bottomDel > bottomID) Then  'destroy buttons from bottomid+1 to bottomdel?
        i = bottomID + 1
        Do While (i <= bottomDel)
            DestroyDelBtn i
            i = i + 1
        Loop
    Else  'create buttons from bottomdel+1 to bottomid?
        i = bottomDel + 1
        Do While (i <= bottomID)
            CreateDelBtn i
            i = i + 1
        Loop
    End If
    If (bottomID = 4) Then
        'recolor from row 6 to row 30
        Range("I6:M30").Interior.Color = RGB(183, 222, 232)
    Else
        'recolor from bottomid+1 to row 30
        Range(Cells(bottomID + 1, 9), Cells(30, 13)).Interior.Color = RGB(183, 222, 232)
    End If
End Sub

Sub CompareAndUpdate(ods As Worksheet, nds As Worksheet)
    'compares current data to displayed data and updates db with changes if approved by user on msgbox?
    Dim Row As Integer
    Dim Col As Integer
    Dim UpdateApproved As Boolean
    UpdateApproved = False
    Row = 5
    Col = 8
    Do While (Len(ods.Cells(Row, Col).Text) > 0)
        Col = 8
        Do While (Col <= 13)
            If (ods.Cells(Row, Col).Value <> nds.Cells(Row, Col).Value) Then
                If (UpdateApproved = False) Then
                    If (MsgBox("Changes exist in values.  Would you like to update these?", vbYesNo) = vbYes) Then
                        UpdateApproved = True
                    Else
                        Exit Sub
                    End If
                End If
                'build update query
                Dim Qry As String
                Qry = "Update CustomerCodeDetails Set QtyPriced = "
                Qry = Qry & nds.Cells(Row, 9).Value & ", UnitPrice = "
                Qry = Qry & nds.Cells(Row, 10).Value & ", StartDate = "
                Qry = Qry & nds.Cells(Row, 11).Value & ", FinishDate = "
                Qry = Qry & nds.Cells(Row, 12).Value & ", DiscountPct = "
                Qry = Qry & nds.Cells(Row, 13).Value & " Where CustomerDetailsID = " & nds.Cells(Row, 8).Value
                SQLProc.ProcQry (Qry)
                Col = 13
            End If
            Col = Col + 1
        Loop
        Row = Row + 1
    Loop
End Sub

Private Sub CreateDelBtn(pos As Integer)
    Dim btn As Button
    Application.ScreenUpdating = False
    Dim t As Range
    Set t = Range(Cells(pos, 14), Cells(pos, 14))
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
        .OnAction = "PriceList.DelBtnPressed"
        .Caption = "Delete"
        .Name = "del" + CStr(pos)
    End With
    t.Value = pos
    Application.ScreenUpdating = True
End Sub
Private Sub DestroyDelBtn(pos As Integer)
'Deletes a Delete Button at the provided range value?
    Cells(pos, 14).ClearContents
    Shapes("del" + CStr(pos)).Delete
End Sub
Private Sub DelBtnPressed()
    Dim Qry As String
    Dim pos As Integer
    pos = Mid(Application.Caller, 4)
    Qry = "DELETE FROM dbo.CustomerCodeDetails WHERE dbo.CustomerCodeDetails.CustomerCodeDetailsID = "
    Qry = Qry + CStr(Cells(pos, 8).Value)
    SQLProc.ProcQry (Qry)
    DestroyDelBtn Range("n3").Value
    UpdateData (True)
End Sub


