<div align="center">

## Create A Windows Shortcut With 9 Lines Of Code


</div>

### Description

A Function To Create Windows Shorcut

.URL for hyperlinks

.LNK for anything else

There are many other variables that can be implemented besides just Name where it points hotkey and description
 
### More Info
 
here is an examle of the code

this will work with Windows XP

CreateShortcut("C:\CMD.lnk", "C:\Windows\System32\CMD.exe", "CTRL+D", "Shortcut to CMD.exe")

Would create a shortcut to the C:\Windows\System32\CMD.exe and place it at C:\CMD.lnk with a hotkey of CTRL+D and a description of Shortcut to CMD.exe


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TfAv1228](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tfav1228.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tfav1228-create-a-windows-shortcut-with-9-lines-of-code__1-62354/archive/master.zip)





### Source Code

```
Public Function CreateShortcut(ShortName As String, PointsTo As String, Optional HotKey As String = "", Optional Description As String = "")
  Dim WS As Object, SC As Object
  Set WS = CreateObject("Wscript.Shell")
  Set SC = WS.CreateShortcut(ShortName)
  SC.TargetPath = PointsTo
  SC.HotKey = HotKey
  SC.Description = Description
  SC.Save
End Function
```

