<div align="center">

## StayOnTop


</div>

### Description

Keep a form always on top (topmost floating form) in windows 95.

Albetski, Allan" <AlbetsAl@amsworld.com>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-stayontop__1-692/archive/master.zip)

### API Declarations

```
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
```


### Source Code

```
Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
 Const HWND_TOPMOST = -1
 Const HWND_NOTOPMOST = -2
 Dim lState As Long
 Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
 With frmForm
  iLeft = .Left / Screen.TwipsPerPixelX
  iTop = .Top / Screen.TwipsPerPixelY
  iWidth = .Width / Screen.TwipsPerPixelX
  iHeight = .Height / Screen.TwipsPerPixelY
 End With
 If fOnTop Then
  lState = HWND_TOPMOST
 Else
  lState = HWND_NOTOPMOST
 End If
 Call SetWindowPos(frmForm.hWnd, lState, iLeft, iTop, iWidth, iHeight,0)
End Sub
```

