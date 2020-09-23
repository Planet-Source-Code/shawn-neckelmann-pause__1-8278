<div align="center">

## Pause


</div>

### Description

Tired of having to pause in increments of 1 second? This coding will pauses based on MILLIseconds using the GetTickCount function.
 
### More Info
 
The number of seconds to pause for. You can put this value down to as little as a millisecond.

Use the CDbl type converter to avoid getting an Invalid Parameter Type error.

Ex:

dim lngPause as Long

lngPause = 2

Call Pause(CDbl(lngPause))

0


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shawn Neckelmann](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shawn-neckelmann.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shawn-neckelmann-pause__1-8278/archive/master.zip)

### API Declarations

```
Declare Function GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long 'this is for 32-bit versions of VB
'Declare Function GetTickCount& Lib "user" () 'this one is for 16-bit versions
```


### Source Code

```
Sub Pause (ByVal hInterval As Double)
Dim hCurrent As Long
hInterval = hInterval * 1000
hCurrent = GetTickCount()
Do While GetTickCount() - hCurrent < hInterval
  DoEvents
Loop
End Sub
```

