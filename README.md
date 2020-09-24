<div align="center">

## multiplication of extra large integers


</div>

### Description

Calculation of the exact value 1000 factorial is nearly impossible by using traditional variables. This code multiplies extra large integers by using strings. 1000! is calculated in 10 seconds
 
### More Info
 
two integers in string form

multiplication result as a string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[zafer ketizmen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/zafer-ketizmen.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/zafer-ketizmen-multiplication-of-extra-large-integers__1-64839/archive/master.zip)





### Source Code

```
' sample of usage calculation of 1000!
' msgbox may truncate the entire string
' so it is set into clipboard
'Dim Ret As String
'Dim k As Integer
'
'Ret = "1"
'For k = 1 To 1000
' Ret = Multiply(CStr(k), Ret)
'Next
'Clipboard.Clear
'Clipboard.SetText (Ret)
'
'MsgBox Ret
Private Function Multiply(a_num1 As String, a_num2 As String) As String
Dim ls_line() As String
Dim ls1 As String
Dim ls2 As String
Dim ls_mul As String
Dim li_num As Integer
Dim li_mul As Integer
Dim li_elde As Integer
Dim li_sum As Integer
Dim li_maxlen As Integer
Dim li_linecount As Integer
Dim li_up As Integer
Dim k As Long
Dim j As Long
' select larger one
Select Case True
 Case Len(a_num1) >= Len(a_num2)
  ls1 = a_num1
  ls2 = a_num2
 Case Len(a_num1) < Len(a_num2)
  ls1 = a_num2
  ls2 = a_num1
End Select
' start multiplication
li_maxlen = -1
For j = Len(ls2) To 1 Step -1
 li_elde = 0
 ls_mul = ""
 li_num = CInt(Mid(ls2, j, 1)) ' number from right
 For k = Len(ls1) To 1 Step -1
  li_mul = li_num * CInt(Mid(ls1, k, 1)) + li_elde ' ex : 7 times 7 = 49
  If k = 1 Then
   ls_mul = CStr(li_mul) + ls_mul
  Else
   ls_mul = CStr(li_mul Mod 10) + ls_mul ' get 9 from 49
   li_elde = (li_mul - (li_mul Mod 10)) / 10 ' remains 4 from 49
  End If
 Next
 ' add extra zeros to the Right
 For k = 1 To Len(ls2) - j + 1 - 1
  ls_mul = ls_mul + "0"
 Next
 ' store result as a one line string
 ReDim Preserve ls_line(1 To Len(ls2) - j + 1)
 ls_line(Len(ls2) - j + 1) = ls_mul
 If Len(ls_mul) > li_maxlen Then li_maxlen = Len(ls_mul)
Next
li_linecount = UBound(ls_line)
' add extra zeros to the Left
For k = 1 To li_linecount
 li_up = li_maxlen - Len(ls_line(k))
 For j = 1 To li_up
  ls_line(k) = "0" + ls_line(k)
 Next
Next
' start summation
li_elde = 0
ls_mul = ""
For k = li_maxlen To 1 Step -1
 li_sum = 0
 For j = 1 To li_linecount
  li_sum = li_sum + CInt(Mid(ls_line(j), k, 1))
 Next
 li_sum = li_sum + li_elde
 If k = 1 Then
  ls_mul = CStr(li_sum) + ls_mul
 Else
  ls_mul = CStr(li_sum Mod 10) + ls_mul
  li_elde = (li_sum - (li_sum Mod 10)) / 10
 End If
Next
Multiply = ls_mul
End Function
```

