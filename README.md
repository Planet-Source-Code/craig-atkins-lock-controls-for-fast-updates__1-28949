<div align="center">

## Lock controls for fast updates


</div>

### Description

When you, for example, add listitems to a listbox or combo box Windows repaints the box after each addidtion. This takes up processing time, makes your app look bad (due to flickering) and if you add a lot of items it can take a long time(Especially if the Combo or Listbox Sorted properety is set to true) Windows provides an API called LockWindowUpdate that you can use to keep Windows from updating a control with a hWnd property. This code is a simple way to call that function.
 
### More Info
 
objX As Object (The object to lock, i.e combo1)

cLock As Boolean (True locks the object, False unlocks it)

To use the sub use something like:

LockControl Combo1, True

..which will prevent the control called Combo1 from repainting hence loading about 30% faster in a non sorted Combo box when adding 10.000 items.

To make the control update again use this code:

LockControl Combo1, False

...which makes Windows update the control Combo1 once again.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Craig Atkins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/craig-atkins.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/craig-atkins-lock-controls-for-fast-updates__1-28949/archive/master.zip)

### API Declarations

```
Option Explicit
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) _
  As Long
'-- End --'
```


### Source Code

```
Public Sub LockControl(objX As Object, cLock As Boolean)
 Dim i As Long
 If cLock Then
  ' This will lock the control
  LockWindowUpdate objX.hWnd
 Else
  ' This will unlock controls
  LockWindowUpdate 0
  objX.Refresh
 End If
End sub
'-- End --'
```

