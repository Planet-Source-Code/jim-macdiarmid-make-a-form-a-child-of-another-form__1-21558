<div align="center">

## Make a form a child of another form


</div>

### Description

Shows a good way to make a form a child of another form without the jumpyness of using just SetParent alone.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-03-06 19:12:52
**By**             |[Jim MacDiarmid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jim-macdiarmid.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD16119362001\.zip](https://github.com/Planet-Source-Code/jim-macdiarmid-make-a-form-a-child-of-another-form__1-21558/archive/master.zip)

### API Declarations

```
Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
```





