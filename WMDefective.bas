Sub wm()


'' Get all excel lists
Dim fileNamesCol As New Collection
Dim MyFile As Variant  'Strings and primitive data types aren't allowed with collection

filepath = "C:\Users\katy.ouyang\Downloads\WM 20210917\"
MyFile = Dir$(filepath & "*.xlsx")
Do While MyFile <> ""
    fileNamesCol.Add (Replace(MyFile, ".xlsx", ""))
    MyFile = Dir$
Loop


Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim header As Variant
header = Array("External ID", _
"Credit #", "Customer", "Date", "Posting Period", "Department", "Location", _
"Currency", "Exchange Rate", "To Be Printed", "To Be E-mailed", "To Be Faxed", _
"Memo", "PO #", "Item", "Quantity", "Price Level", "Rate", "Sale Amnt", _
"Description", "Taxable", "PO details", "Apply_Applied", "Apply_payment")
ws.Range("A1:X1").Value = header

'open Item Check
Dim checkWb As Workbook
Dim check As Worksheet
Set checkWb = Workbooks.Open("Walmart Item Check 2021.xlsx")
Set check = checkWb.Sheets("ItemBasicInfoWalmartDSVReportR")


Dim i As Integer: i = 1
Dim srcWb As Workbook
Dim src As Worksheet
Dim nRow As Integer
Dim lDefective As Integer
Dim lHandling
Dim lFreight
Dim rate
Dim defectiveQty As Integer
Dim lName As Integer
Dim name As String


For Each MyFile In fileNamesCol

  If (Left(fileNamesCol(i), 1) = 9) And (Len(fileNamesCol(i)) = 10) Then




  '' Sam's Club
  ElseIf (Left(fileNamesCol(i), 1) = 1) And (Len(fileNamesCol(i)) = 10) Then
    Set srcWb = Workbooks.Open(fileNamesCol(i) & ".xlsx")
    Set src = srcWb.Sheets("Sheet1")

    nRow = ws.Range("C" & ws.Rows.Count).End(xlUp).Row + 1
    ws.Range("c" & nRow).Value = _
    "Wal-Mart Stores Inc (Dot Com) : Sam's Club.Com"
    ws.Range("D" & nRow).Value = "4/8/2021"
    ws.Range("F" & nRow).Value = "Dot Com"
    ws.Range("G" & nRow).Value = "IL-S"
    ws.Range("H" & nRow).Value = "USD"
    ws.Range("I" & nRow).Value = "1"
    ws.Range("J" & nRow).Value = "FALSE"
    ws.Range("K" & nRow).Value = "FALSE"
    ws.Range("L" & nRow).Value = "FALSE"
    ws.Range("M" & nRow).Value = "Defective Return CK# "
    ws.Range("N" & nRow).Value = "Mdse. Return>" & fileNamesCol(i)
    ws.Range("O" & nRow).Value = "Ad-Hoc Defective"
    ws.Range("P" & nRow).Value = "1"
    ws.Range("Q" & nRow).Value = "Custom"

    ' get DEFECTIVE MDSE info
    lDefective = Application.Match("DEFECTIVE MDSE", src.Range("a:a"), 0)
    defectiveQty = src.Cells(lDefective + 2, 6).Value
    rate = -src.Cells(lDefective + 2, 4).Value
    ws.Range("R" & nRow).Value = rate
    ws.Range("S" & nRow).Value = rate
    'get Description
    lName = Application.Match(src.Cells(lDefective + 2, 1), check.Range("a:a"), 0)
    name = check.Cells(lName, 2).Value
    ws.Range("T" & nRow).Value = name

    ws.Range("U" & nRow).Value = "FALSE"

    'if qty > 1, multple lines
    If lDefective > 1 Then
      For j = 1 To lDefective - 1
        nRow = ws.Range("C" & ws.Rows.Count).End(xlUp).Row + 1
        ws.Rows(nRow - 1).Copy
        ws.Rows(nRow).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
      Next j
    End If

    'if handling applied
    lHandling = Application.Match("HANDLING CHARGE APPLIED", src.Range("a:a"), 0)
    If Not IsError(lHandling) Then '<<< test for no match here
      nRow = ws.Range("C" & ws.Rows.Count).End(xlUp).Row + 1
      ws.Rows(nRow - 1).Copy
      ws.Rows(nRow).PasteSpecial xlPasteAll
      Application.CutCopyMode = False

      ws.Cells(nRow, 15).Value = "Handling Fee"
      ws.Cells(nRow, 18).Value = -src.Cells(lHandling + 2, 4).Value
      ws.Cells(nRow, 19).Value = -src.Cells(lHandling + 2, 4).Value
    End If

      'if freight applied
    lFreight = Application.Match("FREIGHT CHARGE APPLIED", src.Range("a:a"), 0)
    If Not IsError(lFreight) Then '<<< test for no match here
      nRow = ws.Range("C" & ws.Rows.Count).End(xlUp).Row + 1
      ws.Rows(nRow - 1).Copy
      ws.Rows(nRow).PasteSpecial xlPasteAll
      Application.CutCopyMode = False

      ws.Cells(nRow, 15).Value = "Freight prepaid"
      ws.Cells(nRow, 18).Value = -src.Cells(lFreight + 2, 4).Value
      ws.Cells(nRow, 19).Value = -src.Cells(lFreight + 2, 4).Value
    End If

    srcWb.Close

  End If

  i = i + 1

Next MyFile

End Sub


