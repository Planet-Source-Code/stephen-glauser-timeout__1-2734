<div align="center">

## TimeOut \!


</div>

### Description

This is a neat little Sub that will Cause a Pause or "Time Out" in your program.
 
### More Info
 
Duration (1 equals 1 Second, .5 equals a Half Second)

This will cause the program to wait the specified duration before continueing.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephen Glauser](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephen-glauser.md)
**Level**          |Unknown
**User Rating**    |3.0 (18 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephen-glauser-timeout__1-2734/archive/master.zip)





### Source Code

```
Sub TimeOut (duration)
StartTime = Timer
Do While Timer - StartTime < duration
  X = DoEvents()
Loop
End Sub
```

