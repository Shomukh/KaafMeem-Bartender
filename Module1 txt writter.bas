Attribute VB_Name = "Module1"
Sub ExportToNotepad()
Dim wsData As Variant
Dim myFileName As String
Dim FN As Integer
Dim p As Integer, q As Integer
Dim path As String
Dim myString As String
Dim lastrow As Long, lastcolumn As Long

lastrow = Sheets("sheet1").Range("A" & Rows.Count).End(xlUp).Row
lastcolumn = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
path = "C:\Users\Shomu\Desktop\T4"

For p = 1 To lastcolumn
wsData = ActiveSheet.Cells(1, p).Value
If wsData = "" Then Exit Sub
myFileName = wsData
myFileName = myFileName & ".txt"
myFileName = path & myFileName
'MsgBox myFileName
For q = 2 To lastrow
myString = myString & vbCrLf & Cells(q, p)

FN = FreeFile
Open myFileName For Output As #FN
Print #FN, myString
Close #FN
Next q
myString = ""
Next p

End Sub
