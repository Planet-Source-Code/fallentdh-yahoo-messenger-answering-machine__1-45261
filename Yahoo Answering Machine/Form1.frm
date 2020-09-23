VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "Yahoo Answering Machine"
   ClientHeight    =   4230
   ClientLeft      =   2265
   ClientTop       =   2295
   ClientWidth     =   6915
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command3 
      Caption         =   "&Open List"
      Height          =   375
      Left            =   3450
      TabIndex        =   11
      Top             =   3785
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save Selected Text"
      Height          =   385
      Left            =   3450
      TabIndex        =   10
      Top             =   3295
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808080&
      Caption         =   "OFF"
      Height          =   375
      Left            =   4900
      TabIndex        =   7
      Top             =   3750
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000FF&
      Caption         =   "ON"
      Height          =   375
      Left            =   4900
      TabIndex        =   6
      Top             =   3295
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1725
      TabIndex        =   4
      Text            =   " BUSY"
      Top             =   3785
      Width           =   1600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2760
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   7
      TabIndex        =   0
      Top             =   840
      Width           =   6908
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Caption         =   "   About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   15
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Yahoo! Messenger Answering Machine 5.5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   248
      Left            =   1080
      TabIndex        =   8
      Top             =   15
      Width           =   4455
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   315
      Width           =   6975
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "Your Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   75
      TabIndex        =   3
      Top             =   3785
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Number Of PM'S:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   2
      Top             =   3295
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2325
      TabIndex        =   1
      Top             =   3295
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
list1.Clear
Label5.Caption = ""
Label1.Caption = " 0"
Call settextoff
End Sub

Private Sub Command2_Click()
SaveList
End Sub

Private Sub Command3_Click()
Openlist
End Sub

Private Sub Form_Load()
On Error Resume Next
Timer1.Enabled = False
Option2.Value = True
Form1.Width = 7035
Form1.Height = 4635
Call settextoff
End Sub

Private Sub Form_Resize()
Form1.Width = 7035
Form1.Height = 4635
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload frmAbout
Call settextyahoo
End Sub

Private Sub Label6_Click()
frmAbout.Visible = True
End Sub

Private Sub List1_Click()
Label5.Caption = list1.Text
End Sub

Private Sub List1_DblClick()
Option1.Value = False
Option2.Value = True
SendPM (vbNullString)
End Sub

Private Sub Option1_Click()
If Text1.Text = "" Then
MsgBox "Please Enter A Status", vbOKOnly
Option1.Value = False
Option2.Value = True
Exit Sub
Else
Timer1.Enabled = True
Call settexton
End If
Timer1.Enabled = True
Option2.Value = False
End Sub

Private Sub Option2_Click()
Timer1.Enabled = False
Option1.Value = False
Call settextoff
End Sub

Private Sub Timer1_Timer()
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
If imclass >= 1 Then
Pause (0.4)

Dim window(1 To 3) As Long
window(1) = FindWindow("IMClass", vbNullString)
window(2) = FindWindowEx(window(1), 0, "ATL:0054E0E0", vbNullString)
window(3) = FindWindowEx(window(2), 0, "Internet Explorer_Server", vbNullString)
list1.AddItem " " & getietext(window(3)) & " @ " & Now

Pause (0.2)
YChatSend "I am currently not available: [ " & Text1.Text & " ]"
Label1.Caption = Label1.Caption + 1
ClosePm
End If
End Sub
