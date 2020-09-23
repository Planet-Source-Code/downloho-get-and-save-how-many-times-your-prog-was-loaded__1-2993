<div align="center">

## Get and Save how many times your prog was loaded


</div>

### Description

This code will allow you to Save & Get how many times your program has been loaded.

Add a few lines of code to your program's main forms' load procedure. Very easy to understand.
 
### More Info
 
Returns as Long


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[DoWnLoHo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/downloho.md)
**Level**          |Unknown
**User Rating**    |3.4 (47 globes from 14 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/downloho-get-and-save-how-many-times-your-prog-was-loaded__1-2993/archive/master.zip)





### Source Code

```
Public Sub SetLoaded()
'put this in your main forms' Load procedure
'this will set the count
Dim lTemp As Long, sPath As String
lTemp& = GetLoaded&
If Right$(App.Path, 1) <> "\" Then sPath$ = App.Path & "\" & App.EXEName & ".tmp" Else sPath$ = App.Path & App.EXEName & ".tmp"
Open sPath$ For Output As #1
Print #1, lTemp& + 1
Close #1
End Sub
Public Function GetLoaded() As Long
'call this to get how many times program has been loaded
On Error Resume Next
Dim sPath As String, sTemp As String
If Right$(App.Path, 1) <> "\" Then sPath$ = App.Path & "\" & App.EXEName & ".tmp" Else sPath$ = App.Path & App.EXEName & ".tmp"
Open sPath$ For Input As #1
sTemp$ = Input(LOF(1), #1)
Close #1
If sTemp$ = "" Then GetLoaded& = 0 Else GetLoaded& = CLng(sTemp$)
End Function
'works well
'DoWnLoHo
```

