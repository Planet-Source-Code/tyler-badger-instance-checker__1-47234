<div align="center">

## Instance Checker


</div>

### Description

This simple few lines of code will give the user and error message if they are already running your program. Doesn't apply if you are in the ide!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tyler Badger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tyler-badger.md)
**Level**          |Beginner
**User Rating**    |3.6 (18 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tyler-badger-instance-checker__1-47234/archive/master.zip)





### Source Code

```
Private Sub Form_Load() ' the form1 loading sub
If App.PrevInstance Then ' check if running
  MsgBox "Error:" & vbCrLf & "Please switch to your already running app." ' error message
 Dim frm As Form ' set frm variable as a form
 For Each frm In Forms ' get all forms
  Unload frm ' unload the form
  Set frm = Nothing ' set frm variable as nothing
 Next frm ' go to the for each frm again
 End If ' end the if statement
End Sub ' end the form load
```

