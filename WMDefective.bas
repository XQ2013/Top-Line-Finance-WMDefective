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
Set ws = Sheets("Sheet1")

Dim header As Variant
header = Array("External ID", _
"Credit #", "Customer", "Date", "Posting Period", "Department", "Location", _
"Currency", "Exchange Rate", "To Be Printed", "To Be E-mailed", "To Be Faxed", _
"Memo", "PO #", "Item", "Quantity", "Price Level", "Rate", "Sale Amnt", _
"Description", "Taxable", "PO details", "Apply_Applied", "Apply_payment")
ws.Range("A1:X1").Value = header


Dim i As Integer: i = 1
Dim srcWb As Workbook
Dim src As Worksheet
Dim nRow as Integer
Dim lDefective as integer
Dim lHandling
Dim lFreight
Dim rate


For Each MyFile In fileNamesCol

  If (Left(fileNamesCol(i), 1) = 1) And (Len(fileNamesCol(i)) = 10) Then

  Set srcWb = Workbooks.Open(fileNamesCol(i) &".xlsx")
  Set src = srcWb.Sheets("Sheet1")



  ws.Range("A" & i+1).Value = fileNamesCol(i)
  nRow = ws.Range("A" & Ws.Rows.Count).End(xlUp).Row+1
  ws.Range("c" & nRow).Value = _
  "Wal-Mart Stores Inc (Dot Com) : Wal-Mart.com (DSV)"
  ws.Range("D" & nRow).Value = "4/8/2021"
  ws.Range("E" & nRow).Value = "Dot Com"
  ws.Range("f" & nRow).Value = "4/8/2021"
  ws.Range("G" & nRow).Value = "IL-S"
  ws.Range("H" & nRow).Value = "USD"
  ws.Range("I" & nRow).Value = "1"
  ws.Range("J:L" & nRow).Value = "FALSE"
  ws.Range("M" & nRow).Value = "Defective Return CK# "
  ws.Range("N" & nRow).Value = "Mdse. Return>" & fileNamesCol(i)
  ws.Range("O" & nRow).Value = "Ad-Hoc Defective"
  ws.Range("P" & nRow).Value = "1"
  ws.Range("Q" & nRow).Value = "Custom"


  ' get DEFECTIVE MDSE info
  Set lDefective  = Application.Match("DEFECTIVE MDSE", src.Range("a:a"), 0)
  rate = -src.Cells(m + 2, 7).Value / src.Cells(m + 2, 6).Value









  srcWb.Close



  ElseIf (Left(fileNamesCol(i), 1) = 9) And (Len(fileNamesCol(i)) = 10) Then
  ws.Range("A" & i+1).Value = fileNamesCol(i)
  nRow = Ws.Range("A" & Ws.Rows.Count).End(xlUp).Row
  ws.Range("D" & nRow).Value = "4/8/2021"

  End If
  i = i + 1

Next MyFile

End Sub
