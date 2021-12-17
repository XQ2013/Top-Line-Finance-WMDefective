Attribute VB_Name = "Module1"

Function getHeaderRowsArray(ws As Worksheet, col As Integer, fixedValue As String) As Variant

    Dim n As Integer: n = 1
    Dim i As Integer: i = 0
    Dim upcCount As Integer
    upcCount = WorksheetFunction.CountIf(ws.Columns(col), "=" & fixedValue)

    Debug.Print "upcCount" & upcCount

    Dim tempArr As Variant
    ReDim tempArr(1 To upcCount)

    For n = 1 To ws.Range("A" & ws.Rows.Count).End(xlUp).Row
        If ws.Cells(n, col).Value = fixedValue Then
        i = i + 1
        tempArr(i) = n
    End If
    Next n

    getHeaderRowsArray = tempArr
End Function



Function copyEasy(src As Worksheet, _
qty, rate, itemType As String, description As String, _
externalID As Integer, _
ws As Worksheet, inputDate As String, ckNumber As String, xlsFile)

    Dim nRow As Integer
    nRow = ws.Range("C" & ws.Rows.Count).End(xlUp).Row + 1
    ws.Range("A" & nRow).Value = "CR" & WorksheetFunction.Text(externalID, "0000")
    ws.Range("B" & nRow).Value = externalID + 20
    If (Left(xlsFile, 1) = 9) Then
        ws.Range("c" & nRow).Value = _
        "Wal-Mart Stores Inc (Dot Com) : Wal-Mart.com (DSV)"
    Else
        ws.Range("c" & nRow).Value = _
        "Wal-Mart Stores Inc (Dot Com) : Sam's Club.Com"
    End If
    ws.Range("D" & nRow).Value = inputDate
    ws.Range("F" & nRow).Value = "Dot Com"
    ws.Range("G" & nRow).Value = "IL-S"
    ws.Range("H" & nRow).Value = "USD"
    ws.Range("I" & nRow).Value = "1"
    ws.Range("J" & nRow).Value = "FALSE"
    ws.Range("K" & nRow).Value = "FALSE"
    ws.Range("L" & nRow).Value = "FALSE"
    ws.Range("M" & nRow).Value = "Defective Return CK# " & ckNumber
    ws.Range("N" & nRow).Value = "Mdse. Return>" & Left(xlsFile, 10)

    If (itemType = "MERCHANDISE RETURN - DEFECTIVE MERCHANDISE") Or _
    (itemType = "DEFECTIVE MDSE") Then
        ws.Range("O" & nRow).Value = "Ad-Hoc Defective"
    ElseIf (itemType = "HANDLING CHARGE APPLIED") Then
        ws.Range("O" & nRow).Value = "Handling Fee"
    ElseIf (itemType = "FREIGHT CHARGE APPLIED") Then
        ws.Range("O" & nRow).Value = "Freight prepaid"
    End If
    ws.Range("P" & nRow).Value = "1"
    ws.Range("Q" & nRow).Value = "Custom"
    ws.Range("R" & nRow).Value = -rate
    ws.Range("S" & nRow).Value = -rate

    ws.Range("T" & nRow).Value = description

    Dim j As Integer
    If qty > 1 Then
          For j = 1 To qty - 1
        nRow = ws.Range("C" & ws.Rows.Count).End(xlUp).Row + 1
        ws.Rows(nRow - 1).Copy
        ws.Rows(nRow).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
      Next j
    End If




End Function



Sub wm()


''get fileNamesCol Collection(excel lists)
filepath = Application.ActiveWorkbook.Path & "\"

Dim xlsNamesCol As New Collection
Dim xlsFile As Variant 'Strings and primitive data types aren't allowed with collection

xlsFile = Dir$(filepath & "*.xls") 'xls works for xls, xlsx, xlsm
Do While xlsFile <> ""
    xlsNamesCol.Add (xlsFile)
    xlsFile = Dir$
Loop



''get input value from WM Defective xlsm

Dim selfinput As Worksheet
Set selfinput = ThisWorkbook.Sheets("Input")
ThisWorkbook.Sheets.Add(After:=Sheets("Input")).name = "Sheet1"
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim inputDate As String
Dim ckNumber As String
Dim checkFile As String
inputDate = selfinput.Cells(2, 2).Value
ckNumber = selfinput.Cells(3, 2).Value
checkFile = selfinput.Cells(4, 2).Value

'open Item Check
Dim checkWb As Workbook
Dim check As Worksheet
Set checkWb = Workbooks.Open(filepath & checkFile & ".xlsx")
Set check = checkWb.Worksheets(1)



Dim header As Variant
header = Array("External ID", _
"Credit #", "Customer", "Date", "Posting Period", "Department", "Location", _
"Currency", "Exchange Rate", "To Be Printed", "To Be E-mailed", "To Be Faxed", _
"Memo", "PO #", "Item", "Quantity", "Price Level", "Rate", "Sale Amnt", _
"Description", "Taxable", "PO details", "Apply_Applied", "Apply_payment")
ws.Range("A1:X1").Value = header


Dim srcWb As Workbook
Dim src As Worksheet

Dim upcRowsArray As Variant
Dim qty
Dim rate
Dim itemType As String
Dim upc
Dim dUpc
Dim lName As Integer
Dim description As String
Dim externalID As Integer: externalID = 0
Dim i As Integer



For Each xlsFile In xlsNamesCol


    Set srcWb = Workbooks.Open(filepath & xlsFile)
    Set src = srcWb.Worksheets(1)
    externalID = externalID + 1

    If (Left(xlsFile, 1) = 9) Then 'starts with 9, xls file, from Walmart


        upcRowsArray = getHeaderRowsArray(src, 3, "ITEM #")

        For Each r In upcRowsArray
            If (Right(xlsFile, 3) = "xls") Then
                qty = src.Cells(r + 1, 8).Value
                rate = src.Cells(r + 1, 7).Value
                itemType = src.Cells(r + 1, 2).Value
                upc = src.Cells(r + 1, 5).Value
            Else
                qty = src.Cells(r + 2, 8).Value
                rate = src.Cells(r + 2, 7).Value
                itemType = src.Cells(r + 2, 2).Value
                upc = src.Cells(r + 2, 5).Value
            End If

            'get Description
            lName = WorksheetFunction.Match(upc, check.Range("a:a"), 0) 'string won't work
            description = check.Cells(lName, 2).Value


            Call copyEasy(src, qty, rate, itemType, description, externalID, _
            ws, inputDate, ckNumber, xlsFile)

        Next r
        Application.DisplayAlerts = False
        srcWb.Close
        Application.DisplayAlerts = True

    ElseIf (Left(xlsFile, 1) = 1) Then

        upcRowsArray = getHeaderRowsArray(src, 4, "UNIT COST")
        dUpc = -1

        For Each r In upcRowsArray

            qty = src.Cells(r + 1, 6).Value
            rate = src.Cells(r + 1, 4).Value
            itemType = src.Cells(r - 1, 1).Value
            upc = src.Cells(r + 1, 1).Value
            If IsEmpty(upc) = True Then
                upc = dUpc
            End If
            dUpc = upc

            'get Description
            lName = WorksheetFunction.Match(upc, check.Range("a:a"), 0)
            description = check.Cells(lName, 2).Value


            Call copyEasy(src, qty, rate, itemType, description, externalID, _
            ws, inputDate, ckNumber, xlsFile)

        Next r
        Application.DisplayAlerts = False
        srcWb.Close
        Application.DisplayAlerts = True
    End If



Next xlsFile

checkWb.Close
ws.name = WorksheetFunction.Text(inputDate, "mmddyy") & " WM Defective"
ws.Copy

ActiveWorkbook.SaveAs Filename:=filepath & ws.name, _
FileFormat:=xlCSV, CreateBackup:=True


Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True



End Sub














