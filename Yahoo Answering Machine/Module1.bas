Attribute VB_Name = "Module1"
' sorry for including unused functions and consts
' i did it so that i won't have to later on

Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const WM_COMMAND = &H111

Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Sub Pause(interval)
'pauses for a number of seconds before doing anything else
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Sub StayOnTop(frm As Form)
'makes the form stay on top all the time
Call SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub YChatSend(what2say As String)
'sends text to chat
Dim imclass As Long
Dim richedit As Long
Dim Button As Long
imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call SendMessageByString(richedit, WM_SETTEXT, 0&, what2say)
imclass = FindWindow("imclass", vbNullString)
Button = FindWindowEx(imclass, 0&, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub SaveList()
'saves selected list1 text to desktop
Dim msgs As String
msgs = Form1.List1.Text
Open App.Path & "\yahoo messages.txt" For Append As #1
Write #1, msgs
Close #1

End Sub

Sub ClosePm()
'Closes the yahoo messenger window
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)
End Sub

Sub SendPM(who As String)
'opens up yahoo message box
If ShellExecute(&O0, "Open", "ymsgr:sendIM?" & who$, vbNullString, vbNullString, SW_NORMAL) < 33 Then
End If
End Sub

Public Sub settextoff()
'sets the yahoo window text to "answering status:off"
Dim yahoofriendlist As Long
Dim SetCaption As Long
yahoofriendlist = FindWindow("YahooBuddyMain", vbNullString)
SetCaption = SendMessageByString(yahoofriendlist, WM_SETTEXT, 0, "Answering Status: Off")
End Sub

Public Sub settexton()
'sets the yahoo window text to "anwswering status:on"
Dim yahoofriendlist As Long
Dim SetCaption As Long
yahoofriendlist = FindWindow("YahooBuddyMain", vbNullString)
SetCaption = SendMessageByString(yahoofriendlist, WM_SETTEXT, 0, "Answering Status: On")
End Sub

Public Sub settextyahoo()
'changes the yahoo window text back
Dim yahoofriendlist As Long
Dim SetCaption As Long
yahoofriendlist = FindWindow("YahooBuddyMain", vbNullString)
SetCaption = SendMessageByString(yahoofriendlist, WM_SETTEXT, 0, "Yahoo! Messenger")
End Sub

Sub Openlist()
'opens notepad and loads yahoo messages.txt
On Error GoTo NT ' If notepad wasn't found in C:\Windows, the user
                 ' could have NT
Dim Openlist As String
Openlist = Shell("C:\Windows\Notepad.exe " & App.Path & "\yahoo messages.txt", vbMaximizedFocus)
Exit Sub
NT:
Openlist = Shell("C:\WinNT\Notepad.exe " & App.Path & "\yahoo messages.txt", vbMaximizedFocus)
End Sub
