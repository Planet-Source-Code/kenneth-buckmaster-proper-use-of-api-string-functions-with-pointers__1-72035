<div align="center">

## Proper use of API string functions \(with pointers\)


</div>

### Description

People usually declare Api string functions using string variables, think they can only print a string from the first character and use functions like mid to trim the string first. These functions are actually meant to be used with pointers to any starting character.

*Added use of integer arrays with wide functions
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kenneth Buckmaster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kenneth-buckmaster.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kenneth-buckmaster-proper-use-of-api-string-functions-with-pointers__1-72035/archive/master.zip)





### Source Code

```
'People usually declare Api string
'functions using string variables,
'think they can only print a
'string from the first character
'and use functions like mid
'to trim the string first.
'These functions are actually
'meant to be used with pointers
'And can be called from any
'starting character
'place in a form
'instead of declaring lpstring
'as a string,
'we declare it as Any
' (or we could use a long
'if we only ever wanted
'to use pointers)
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal _
X As Long, ByVal Y As Long, ByVal lpString As Any, ByVal nCount As Long) As Long
Private Declare Function TextOutW Lib "gdi32.dll" (ByVal hdc As Long, ByVal _
X As Long, ByVal Y As Long, ByRef lpString As Any, ByVal nCount As Long) As Long
'(API functions typically have
'Ansi (one byte character)
'and Wide (two byte charecter)
'versions)
'for helper functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Sub Form_Load()
Me.AutoRedraw = True
Me.ScaleMode = 3
Dim st As String, substr As String
'working with TextoutW
'(two byte characters)
st = "the fish"
'strptr function gives the location
'in memory of the zeroth character
'add 8 to get the 4th character
'as each is two bytes long
'... prints fish
TextOutW Me.hdc, 0, 0, ByVal StrPtr(st) + 8, 4
'***********ANSI*********
'Ansi functions like Textout use
'single byte characters
'VB6 uses two byte characters
'so first we need to do what vb does
'when it calls textout with
'lpstring as a string declare
'- VB converts it's 2byte
'Character strings to a
'single byte one
st = StrConv("the fish", vbFromUnicode)
TextOut Me.hdc, 0, 30, ByVal StrPtr(st) + 4, 4
'without using pointers
'we'd need to first create a
'new "fish" string
substr = Mid("the fish", 5, 4)
TextOut Me.hdc, 0, 60, substr, 4
'vb did the same thing there
'that we did -
'converted the string to
'single byte characters
'then sent a pointer (to the first
'character in the new string)
'clearly its innefficient to
'create a new string just
'to use the TextOut function
'So we'd convert our string
'when we create it
'and then work in single bytes
'all the time
'eg no need to convert for
'other calls
TextOut Me.hdc, 0, 90, ByVal StrPtr(st), 3
'("the")
'**********byte arrays
'alternatively we can use byte arrays
Dim b() As Byte
b() = "the fish"
'varptr function gives the location
'of the byte in the array index
'remember each character
'is two bytes here
TextOutW Me.hdc, 0, 120, ByVal VarPtr(b(8)), 4
'For one byte character byte arrays
b = StrConv("the fish", vbFromUnicode)
TextOut Me.hdc, 0, 150, ByVal VarPtr(b(4)), 4
'byte arrays are better for
'manipulating
b(4) = 68 'vaue for "D"
TextOut Me.hdc, 0, 180, ByVal VarPtr(b(4)), 4
'**********integer arrays
'With wide functions
'we may also choose
'to use integer arrays
'instead of byte arrays
Dim i() As Integer
st = "The Bats"
'with integer arrays we
'can't convert from
'strings using =
'so we need the helper functions
i = stringToIntegerArray(st)
TextOutW Me.hdc, 0, 210, i(4), 3
i(4) = 72
TextOutW Me.hdc, 0, 240, i(4), 3
'Notice, I declared lpString
'in TextOutW byRef.
'This allowed me to put i(4)
'instead of Varptr
'(VB does that litte bit
'of work for us)
'we could do the same thing
'with TextOutA and a byte array
'for single byte Ansi
'characters if we wanted.
End Sub
'Helper functions
Function stringToIntegerArray(st As String) As Integer()
Dim i() As Integer, lenst As Long
If Len(st) > 0 Then
lenst = Len(st)
ReDim i(lenst - 1)
CopyMemory i(0), ByVal StrPtr(st), lenst * 2
stringToIntegerArray = i
End If
End Function
Function IntegerArrayToString(i() As Integer) As String
On Error GoTo NotDimensioned
Dim lenst As Long
lenst = UBound(i) + 1
IntegerArrayToString = Space(lenst)
CopyMemory ByVal StrPtr(IntegerArrayToString), i(0), lenst * 2
NotDimensioned:
End Function
'other string functions like
'drawtext, gettextextentpoint32
'etc can all be used in
'the same way
'NOTE VB fails with
'TextOutW when
'we don't use pointers
'TextOutW Me.hdc, 0, 0, "fish", 4
'fails because VB doesn't
'recognise it as a wide
'function and converts
'to single byte characters
'as it would with TextoutA
```

