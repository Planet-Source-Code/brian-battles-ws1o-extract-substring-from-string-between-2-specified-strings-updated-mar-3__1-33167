<div align="center">

## Extract substring from string between 2 specified strings, UPDATED Mar 30, 2002


</div>

### Description

Extracts the contents of a "child" substring from a "parent" string between 2 additional specified substrings.

For example, if you called it like this:

GetStringBetween("Fourscore and seven years ago, our fathers", " and ", " our ")

It will return

"seven years ago,"
 
### More Info
 
"parent" string, string to begin after, optional string to end before, optional case sensitive flag

Pass this function a string (strCompleteString), and it will return a substring consisting of everything between 2 other specified strings (ie, everything between strFirst and strLast)

You can also optionally specify if it should be case sensitive (default is False)

Example: GetStringBetween("Fourscore and seven years ago, our fathers", "and", "our") would return "seven years ago,"

Bonus: if you leave out the last string, you'll just get the word following the first word

Example: GetStringBetween("Fourscore and seven years ago, our fathers", "and") would return "seven"

String

Not really a "side effect," more of a "bonus," if you pass in a space " " as the end string, you just get back the one full word after the first string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Battles WS1O](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-battles-ws1o.md)
**Level**          |Beginner
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-battles-ws1o-extract-substring-from-string-between-2-specified-strings-updated-mar-3__1-33167/archive/master.zip)





### Source Code

```
Public Function GetStringBetween(strCompleteString As String, strFirst As String, Optional strLast As String, Optional bCaseSensitive As Boolean = False) As String
  ' Purpose  : Pass this function a string (strCompleteString),
  '        and it will return a substring consisting of
  '        everything between 2 other specified strings
  '        (ie, everything between strFirst and strLast)
  '       You can also optionally specify if it should be case sensitive (default is False)
  '       Bonus: if you leave out the last string, you'll
  '           just get the word following the first word
  '          Or if you leave off the first string, it will start from the
  '           first character in the main string
  ' Example  : GetStringBetween("Fourscore and seven years ago, our fathers", "and", "our")
  '        would return "seven years ago,"
  ' Parameters: strCompleteString, strFirst, strLast, bCaseSensitive
  ' Returns  : String
  ' Modified : 3/30/2002 By BB
  Dim iPos   As Integer
  Dim iLen   As Integer
  Dim strTemp1 As String
  Dim strTemp2 As String
  On Error GoTo Err_GetStringBetween
  ' make sure we have valid values to work with
  If Len(strCompleteString) = 0 Then
    ' no string to parse
    MsgBox "Missing Main String, Nothing to Parse", vbInformation, "Advisory"
    strTemp2 = ""
    GoTo Exit_GetStringBetween
  ElseIf Len(strFirst) = 0 Then
    ' no beginning string, so begin at first character
    iPos = 1
  ElseIf Len(strLast) = 0 Then
    ' no ending string, so we'll make it a space
    strLast = " "
  End If
  ' if no beginning was specified, we can skip this
  If iPos < 1 Then
    ' get the location in the string where our first string occurs
    If bCaseSensitive Then
      ' case sensitive
      iPos = InStr(1, strCompleteString, strFirst, vbBinaryCompare)
    Else
      ' case insensitive
      iPos = InStr(1, strCompleteString, strFirst, vbTextCompare) ' default
    End If
  End If
  ' assuming we did find the first string...
  If iPos > 0 Then
    ' extract everything to the right of the first string;
    ' we use the expression
    '   Len(strCompleteString) - (iPos + Len(Trim$(strFirst)
    ' to determine where the first string actually ends,
    ' the Trim$ call makes sure we don't include any spaces the user may have passed in
    ' (you have to pass in the spaces around a word to distinguish a complete word
    ' from a string that may appear within a word, eg, the "and" in "thousand" would
    ' mess us up if we had called it like this:
    '  GetStringBetween("Four thousand and seven years ago", "and", "ago")
    ' so the right way to call it would be this:
    '  GetStringBetween("Four thousand and seven years ago", " and ", "ago")
    '
    ' I hope that makes it clear!
    If iPos = 1 Then
      strTemp1 = Trim$(Right$(strCompleteString, Len(strCompleteString)))
    Else
      strTemp1 = Trim$(Right$(strCompleteString, Len(strCompleteString) - (iPos + Len(Trim$(strFirst)))))
    End If
  End If
  If (LCase$(strFirst) = " inner join ") And (LCase$(strLast) = " on ") Then
    iLen = Len(strTemp1)
    If bCaseSensitive Then
      ' case sensitive
      iPos = InStrRev(strTemp1, strLast, iLen, vbBinaryCompare)
    Else
      ' case insensitive
      iPos = InStrRev(strTemp1, strLast, iLen, vbTextCompare) ' default
    End If
    If iPos > 0 Then
      strTemp2 = " INNER JOIN " & Trim$(Left$(strTemp1, iPos - 1)) & " ON "
    Else
      strTemp2 = strTemp1
    End If
  Else
    If bCaseSensitive Then
      ' case sensitive
      iPos = InStr(1, strTemp1, strLast, vbBinaryCompare)
    Else
      ' case insensitive
      iPos = InStr(1, strTemp1, strLast, vbTextCompare) ' default
    End If
    If iPos > 0 Then
      strTemp2 = Trim$(Left$(strTemp1, iPos - 1))
    Else
      strTemp2 = strTemp1
    End If
  End If
Exit_GetStringBetween:
  On Error Resume Next
  GetStringBetween = strTemp2
  On Error GoTo 0
  Exit Function
Err_GetStringBetween:
  Select Case Err
    Case 0
      Resume Next
    Case Else
      MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during GetStringBetween" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
      strTemp2 = ""
      Resume Exit_GetStringBetween
  End Select
End Function
```

