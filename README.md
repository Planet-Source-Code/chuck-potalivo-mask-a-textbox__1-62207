<div align="center">

## Mask a Textbox


</div>

### Description

This function will turn any textbox into a MaskEdit box! Just call this function from the textbox Change event, or KeyUp events.

It may not seem like much, but it works, and I worked hard to figure this out just right. This is PURE VB code! Email me for a VB.NET version of this routine, although it is pretty easy to convert.

Please rate my submissions! I enjoy hearing feedback on my code. Thanks all.
 
### More Info
 
A textbox object, and a mask string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chuck Potalivo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chuck-potalivo.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chuck-potalivo-mask-a-textbox__1-62207/archive/master.zip)





### Source Code

```
Private Function MaskText(txtTarget As VB.Textbox, strMask As String)
 Static bolRunning   As Boolean
 If bolRunning Then
 Exit Function
 End If
 bolRunning = True
 Dim strTarget_Text   As String
 strTarget_Text = txtTarget.Text
 Dim lngCursor_Pos   As Long
 lngCursor_Pos = txtTarget.SelStart
 If Len(strMask) > Len(strTarget_Text) Then
 strTarget_Text = strTarget_Text & Space(Len(strMask) - Len(strTarget_Text))
 ElseIf Len(strMask) < Len(strTarget_Text) Then
 strTarget_Text = Left(strTarget_Text, Len(strMask))
 ElseIf Len(strMask) = 0 Then
 Exit Function
 End If
 Dim strTarget_Char   As String * 1
 Dim strMask_Char   As String * 1
 Dim strTemp     As String
 Dim bolAlpha    As Boolean
 Dim aryLiterals    As Variant
 aryLiterals = Array("(", ")", "-", ".", ",", ":", ";", "/", "\", " ")
 Dim lngLiteral_Index  As Long
 Dim bolLiteral    As Boolean
 Dim lngChar_Index   As Long
 For lngChar_Index = 1 To Len(strMask)
 strTarget_Char = Mid(strTarget_Text, lngChar_Index, 1)
 strMask_Char = Mid(strMask, lngChar_Index, 1)
 For lngLiteral_Index = LBound(aryLiterals) To UBound(aryLiterals)
  bolLiteral = (strMask_Char = aryLiterals(lngLiteral_Index))
  If bolLiteral Then
  Exit For
  End If
 Next lngLiteral_Index
 Select Case strMask_Char
  Case "#":
  If (Not IsNumeric(strTarget_Char)) And (strTarget_Char <> " ") Then
   strTemp = Right(strTarget_Text, Len(strTarget_Text) - lngChar_Index)
   If lngChar_Index > 1 Then
   strTarget_Text = Left(strTarget_Text, lngChar_Index - 1)
   Else
   strTarget_Text = ""
   End If
   strTarget_Text = strTarget_Text & " " & strTemp
  End If
  Case "@":
  bolAlpha = ((Asc(strTarget_Char) >= 65) And (Asc(strTarget_Char) <= 90)) Or ((Asc(strTarget_Char) >= 97) And (Asc(strTarget_Char) <= 122))
  If (Not bolAlpha) And (strTarget_Char <> " ") Then
   strTemp = Right(strTarget_Text, Len(strTarget_Text) - lngChar_Index)
   If lngChar_Index > 1 Then
   strTarget_Text = Left(strTarget_Text, lngChar_Index - 1)
   Else
   strTarget_Text = ""
   End If
   strTarget_Text = strTarget_Text & " " & strTemp
  End If
  Case Else:
  If (strTarget_Char <> strMask_Char) And bolLiteral Then
   strTemp = Right(strTarget_Text, Len(strTarget_Text) - (lngChar_Index - 1))
   strTarget_Text = Left(strTarget_Text, lngChar_Index - 1)
   strTarget_Text = strTarget_Text & strMask_Char & strTemp
   If lngChar_Index = lngCursor_Pos Then
   lngCursor_Pos = lngCursor_Pos + 1
   End If
  End If
 End Select
 Next lngChar_Index
 txtTarget.Text = Left(strTarget_Text, Len(strMask))
 txtTarget.SelStart = lngCursor_Pos
 bolRunning = False
End Function
```

