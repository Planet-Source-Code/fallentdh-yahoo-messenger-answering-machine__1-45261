Attribute VB_Name = "Spiyre"
Option Explicit
'CODED BY JOHN CASEY; SPIYRE@MSN.COM
'DONT FORGET REFERENCE TO "MICROSOFT HTML OBJECT LIBRARY"

Public Declare Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Long, riid As UUID, ByVal wParam As Long, ppvObject As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type
   
Public Const SMTO_ABORTIFHUNG = &H2


Function getietext(ByVal hWnd As Long) As String
On Error Resume Next
Dim doc As IHTMLDocument2

If hWnd <> 0 Then
Set doc = IEDOMFromhWnd(hWnd)
Else
getietext = "-[TEXT CANNOT BE FOUND]-"
Exit Function
End If
'---CHECKS-FOR-HWND------


If doc.body.innerText = vbNullString Then
getietext = "ERROR! [WINDOW DOESN'T CONTAIN HTML]"
Exit Function
End If
'---CHECKS-FOR-HTML-EMBEDDED

getietext = doc.body.innerText

End Function

Function IEDOMFromhWnd(ByVal hWnd As Long) As IHTMLDocument
Dim IID_IHTMLDocument As UUID
Dim doc As IHTMLDocument2
Dim lRes As Long 'if = 0 isn't inet window.
Dim lMsg As Long
Dim hr As Long
'---END-DECLARES---------

lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT") 'Register Wnd Message
Call SendMessageTimeout(hWnd, lMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, lRes) 'Get's Object


'---CHECKS-FOR-WINDOW----
hr = ObjectFromLresult(lRes, IID_IHTMLDocument, 0, IEDOMFromhWnd)

End Function
