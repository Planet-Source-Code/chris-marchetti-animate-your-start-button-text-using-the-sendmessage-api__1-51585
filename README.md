<div align="center">

## Animate your start button text using the sendmessage api\.


</div>

### Description

This code will allow you to animate your start button text to whatever you want using the sendmessage api.
 
### More Info
 
The start bar button text will only stay this way untill your next restart of your computer.

The program was only texted on windows xp, i cant say if it will work on any other os's.


<span>             |<span>
---                |---
**Submitted On**   |2004-02-08 02:14:28
**By**             |[Chris Marchetti](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-marchetti.md)
**Level**          |Beginner
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Animate\_yo170567282004\.zip](https://github.com/Planet-Source-Code/chris-marchetti-animate-your-start-button-text-using-the-sendmessage-api__1-51585/archive/master.zip)

### API Declarations

```
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
```





