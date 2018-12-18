# FSO
Sub tonghop()
Dim arr, arr1, arr2
Dim lr As Long, a As Long, b As Long, i As Long, j As Long, d As Long, c As Long
Dim dk As String
Dim sh As Worksheet
Dim dic As Object
Dim tong As Worksheet
Set dic = CreateObject("scripting.dictionary")
Set tong = Sheets("tong hop")
b = tong.Range("B" & Rows.Count).End(xlUp).Row
If b > 2 Then
   arr = tong.Range("B2:K" & b).Value
   For i = 1 To UBound(arr, 1)
     dk = Empty
      For j = 1 To 10
          If dk = Empty Then
             dk = arr(i, j)
          Else
             dk = dk & "#" & arr(i, j)
          End If
      Next j
      If dic.exists(dk) = 0 Then
         dic.Item(dk) = "KK"
      End If
  Next i
End If
For Each sh In ThisWorkbook.Worksheets
    If sh.Name <> "Tong hop" Then
        c = sh.Range("B" & Rows.Count).End(xlUp).Row
        d = d + c
    End If
Next
ReDim arr1(1 To d, 1 To 11)
For Each sh In ThisWorkbook.Worksheets
    If sh.Name <> "Tong hop" Then
       lr = sh.Range("B" & Rows.Count).End(xlUp).Row
       If lr > 1 Then
          arr2 = sh.Range("B2:K" & lr).Value
         For i = 1 To UBound(arr2, 1)
                dk = Empty
                For j = 1 To 10
                  If dk = Empty Then
                     dk = arr2(i, j)
                  Else
                     dk = dk & "#" & arr2(i, j)
                  End If
                Next j
                If dic.exists(dk) = 0 Then
                dic.Item(dk) = "KK"
                a = a + 1
                arr1(a, 1) = a + b - 1
                  For j = 2 To 11
                    arr1(a, j) = arr2(i, j - 1)
                  Next j
                End If
        Next i
      End If
   End If
Next
    If a Then tong.Range("A" & b + 1).Resize(a, 11).Value = arr1

MsgBox "Da cap nhap duoc :" & a & " dong"
End Sub
