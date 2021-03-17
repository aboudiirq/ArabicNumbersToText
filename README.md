# ArabicNumbersToText
Arabic Numbers To Text in microsoft word or excel


example Ms Word Code
Dim Num As String
On Error Resume Next
    Num = NumberToText(Selection, "دينار", "")
  Selection.TypeText Text:=Num
