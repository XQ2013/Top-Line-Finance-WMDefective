Attribute VB_Name = "Module1"

Sub test()


Dim fileNamesCol As New Collection
Dim MyFile As Variant  'Strings and primitive data types aren't allowed with collection

filepath = "C:\Users\katy.ouyang\Downloads\WM 20210917\"
MyFile = Dir$(filepath & "*.xlsx")
Do While MyFile <> ""
    fileNamesCol.Add (Replace(MyFile, ".xlsx", ""))
    MyFile = Dir$
Loop

Dim myWs As Worksheet
Set myWs = Sheets("Sheet1")
Dim ic As Integer: ic = 1

For Each MyFile In fileNamesCol
    myWs.Range("A" & ic).Value = fileNamesCol(ic)
    ic = ic + 1
Next MyFile


End Sub



