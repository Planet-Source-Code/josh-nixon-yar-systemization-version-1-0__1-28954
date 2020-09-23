<div align="center">

## Yar Systemization Version 1\.0 \(\)\)\)

<img src="YarSystem.jpg">
</div>

### Description

*!*!*!*!*!*!*!Attention:Screen shots

www.geocities.com/nit3shift PSC would not upload them. :(

----

----

)

I releaed this earlier as a beta. Which wasnt really good cause i wasnt done and the user could not take on the full power of the program. The thought came to me one day when I was chillen on my bed I go hey, it would be a real good idea I could make a program that goes through my whole HHD and sort out my files. So i set off to make this program took about a month, nothing real real specail, however it has a graphic interface to it with folder that may be moved around. desktop backgrounds that may be tiled. When you first run the program it notices that this the first time that you have ran the program. It then runs a wizard for you. The program has favorite start up wavs and favorite program that can be executed by the click on the button. All data can be saved and loaded back again for your comfort. This is all for now i ahve a crapy webpage for a screenshots which i am still working on. Please comments or votes. Tell about any errors i can fix or make the program better. Thanks peace Out and enjoy.
 
### More Info
 
Joy

None that i know off please tell me if there is. I did make it so it fit all resolutions.


<span>             |<span>
---                |---
**Submitted On**   |2001-11-16 05:58:04
**By**             |[Josh Nixon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/josh-nixon.md)
**Level**          |Advanced
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Yar\_System3566211162001\.zip](https://github.com/Planet-Source-Code/josh-nixon-yar-systemization-version-1-0__1-28954/archive/master.zip)

### API Declarations

```
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public counter
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public i As Long
Public ani As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal index As Long) As Long
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Swapmouse, auto_arrange As Boolean
Const AC_SRC_OVER = &H0
Public Declare Function Pie Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public binmar1 As Boolean
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Type POINTAPI
  X As Long
  Y As Long
End Type
Global Const MINUTES = 15
  Public TimeOver
  Public MinUp
Public Const LB_SETHORIZONTALEXTENT = &H194
Public tempstring, tempstring2, filena, filenas, no As String
Dim start, lstindex
Public images As Boolean
Public other As Boolean
Public app As Boolean
Public Text As Boolean
Public media As Boolean
```





