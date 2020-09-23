VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4890
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   130
      TabIndex        =   8
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Elite Talker"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Black Name"
      Height          =   255
      Left            =   1800
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   130
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":4F49A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   3360
      ScaleHeight     =   2385
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   960
      Width           =   1215
      Begin VB.Image Image32 
         Height          =   180
         Left            =   840
         Picture         =   "Form1.frx":4F54A
         Top             =   1800
         Width           =   180
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New Line"
         ForeColor       =   &H80000008&
         Height          =   235
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   900
      End
      Begin VB.Image Image31 
         Height          =   180
         Left            =   840
         Picture         =   "Form1.frx":4F73C
         Top             =   1560
         Width           =   180
      End
      Begin VB.Image Image30 
         Height          =   180
         Left            =   120
         Picture         =   "Form1.frx":4F92E
         Top             =   1320
         Width           =   180
      End
      Begin VB.Image Image29 
         Height          =   180
         Left            =   600
         Picture         =   "Form1.frx":4FB20
         Top             =   1320
         Width           =   180
      End
      Begin VB.Image Image28 
         Height          =   180
         Left            =   360
         Picture         =   "Form1.frx":4FD12
         Top             =   840
         Width           =   180
      End
      Begin VB.Image Image27 
         Height          =   180
         Left            =   120
         Picture         =   "Form1.frx":4FF04
         Top             =   840
         Width           =   195
      End
      Begin VB.Image Image26 
         Height          =   180
         Left            =   600
         Picture         =   "Form1.frx":50126
         Top             =   1800
         Width           =   180
      End
      Begin VB.Image Image25 
         Height          =   180
         Left            =   360
         Picture         =   "Form1.frx":50318
         Top             =   600
         Width           =   180
      End
      Begin VB.Image Image24 
         Height          =   180
         Left            =   120
         Picture         =   "Form1.frx":5050A
         Top             =   600
         Width           =   180
      End
      Begin VB.Image Image23 
         Height          =   180
         Left            =   600
         Picture         =   "Form1.frx":506FC
         Top             =   1560
         Width           =   180
      End
      Begin VB.Image Image22 
         Height          =   180
         Left            =   360
         Picture         =   "Form1.frx":508EE
         Top             =   1800
         Width           =   180
      End
      Begin VB.Image Image21 
         Height          =   180
         Left            =   600
         Picture         =   "Form1.frx":50AE0
         Top             =   840
         Width           =   180
      End
      Begin VB.Image Image20 
         Height          =   180
         Left            =   840
         Picture         =   "Form1.frx":50CD2
         Top             =   120
         Width           =   180
      End
      Begin VB.Image Image19 
         Height          =   180
         Left            =   840
         Picture         =   "Form1.frx":50EC4
         Top             =   1080
         Width           =   180
      End
      Begin VB.Image Image18 
         Height          =   180
         Left            =   600
         Picture         =   "Form1.frx":510B6
         Top             =   1080
         Width           =   180
      End
      Begin VB.Image Image17 
         Height          =   180
         Left            =   840
         Picture         =   "Form1.frx":512A8
         Top             =   840
         Width           =   180
      End
      Begin VB.Image Image16 
         Height          =   180
         Left            =   120
         Picture         =   "Form1.frx":5149A
         Top             =   1080
         Width           =   180
      End
      Begin VB.Image Image15 
         Height          =   180
         Left            =   360
         Picture         =   "Form1.frx":5168C
         Top             =   1320
         Width           =   180
      End
      Begin VB.Image Image14 
         Height          =   180
         Left            =   120
         Picture         =   "Form1.frx":5187E
         Top             =   1800
         Width           =   180
      End
      Begin VB.Image Image13 
         Height          =   180
         Left            =   840
         Picture         =   "Form1.frx":51A70
         Top             =   1320
         Width           =   180
      End
      Begin VB.Image Image12 
         Height          =   180
         Left            =   120
         Picture         =   "Form1.frx":51C62
         Top             =   1560
         Width           =   180
      End
      Begin VB.Image Image11 
         Height          =   180
         Left            =   360
         Picture         =   "Form1.frx":51E54
         Top             =   1560
         Width           =   180
      End
      Begin VB.Image Image10 
         Height          =   180
         Left            =   840
         Picture         =   "Form1.frx":52046
         Top             =   600
         Width           =   180
      End
      Begin VB.Image Image9 
         Height          =   180
         Left            =   840
         Picture         =   "Form1.frx":52238
         Top             =   360
         Width           =   180
      End
      Begin VB.Image Image8 
         Height          =   180
         Left            =   600
         Picture         =   "Form1.frx":5242A
         Top             =   360
         Width           =   180
      End
      Begin VB.Image Image7 
         Height          =   180
         Left            =   120
         Picture         =   "Form1.frx":5261C
         Top             =   120
         Width           =   180
      End
      Begin VB.Image Image6 
         Height          =   180
         Left            =   120
         Picture         =   "Form1.frx":5280E
         Top             =   360
         Width           =   180
      End
      Begin VB.Image Image5 
         Height          =   180
         Left            =   360
         Picture         =   "Form1.frx":52A00
         Top             =   1080
         Width           =   180
      End
      Begin VB.Image Image4 
         Height          =   180
         Left            =   600
         Picture         =   "Form1.frx":52BF2
         Top             =   600
         Width           =   180
      End
      Begin VB.Image Image3 
         Height          =   180
         Left            =   600
         Picture         =   "Form1.frx":52DE4
         Top             =   120
         Width           =   180
      End
      Begin VB.Image Image2 
         Height          =   180
         Left            =   360
         Picture         =   "Form1.frx":52FD6
         Top             =   360
         Width           =   180
      End
      Begin VB.Image Image1 
         Height          =   180
         Left            =   360
         Picture         =   "Form1.frx":531C8
         Top             =   120
         Width           =   180
      End
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  X  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MSN Name Maker V1.0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   300
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear Name"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4120
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Name"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents MSN As MsgrObject
Attribute MSN.VB_VarHelpID = -1
Dim X As Integer
Dim ChatName2 As String
'Dim clip As String
'Dim clip2 As IPictureDisp

Private Sub Form_Load()
On Error Resume Next
DoEvents
Set MSN = New MsgrObject
rt.SelColor = &H80000001
End Sub

Private Sub Check1_Click()
Dim str
If Check1.Value = 0 Then
rt.SelColor = &H80000001
Else
rt.SelColor = &H80000007
End If
End Sub
Function SetClipboard()
'Clipboard.SetData = clip2
'Clipboard.SetText = clip
End Function
Function GetClipboard()
Clipboard.Clear
'clip = Clipboard.GetText
'clip2 = Clipboard.GetData
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim l As ColorConstants
l = vbWhite
Label2.BackColor = l
Label5.BackColor = l
End Sub

Private Sub Image32_Click()
ChatName "(T)"
GetClipboard
Clipboard.SetData Image32.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub Image7_Click()
ChatName ":d"
GetClipboard
Clipboard.SetData Image7.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image1_Click()
ChatName ":)"
GetClipboard
Clipboard.SetData Image1.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image2_Click()
ChatName ":("
GetClipboard
Clipboard.SetData Image2.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image3_Click()
ChatName ":|"
GetClipboard
Clipboard.SetData Image3.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image4_Click()
ChatName "(B)"
GetClipboard
Clipboard.SetData Image4.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image5_Click()
ChatName "(U)"
GetClipboard
Clipboard.SetData Image5.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image6_Click()
ChatName ":s"
GetClipboard
Clipboard.SetData Image6.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub Image8_Click()
ChatName ":p"
GetClipboard
Clipboard.SetData Image8.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image9_Click()
ChatName ":o"
GetClipboard
Clipboard.SetData Image9.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image10_Click()
ChatName "(D)"
GetClipboard
Clipboard.SetData Image10.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image11_Click()
ChatName "(G)"
GetClipboard
Clipboard.SetData Image11.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image12_Click()
ChatName "(E)"
GetClipboard
Clipboard.SetData Image12.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image13_Click()
ChatName "(F)"
GetClipboard
Clipboard.SetData Image13.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image14_Click()
ChatName "(K)"
GetClipboard
Clipboard.SetData Image14.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image15_Click()
ChatName "(M)"
GetClipboard
Clipboard.SetData Image15.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image16_Click()
ChatName "(L)"
GetClipboard
Clipboard.SetData Image16.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image17_Click()
ChatName ":["
GetClipboard
Clipboard.SetData Image17.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image18_Click()
ChatName "(N)"
GetClipboard
Clipboard.SetData Image18.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image19_Click()
ChatName "(Y)"
GetClipboard
Clipboard.SetData Image19.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image20_Click()
ChatName ";)"
GetClipboard
Clipboard.SetData Image20.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image21_Click()
ChatName "(P)"
GetClipboard
Clipboard.SetData Image21.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image22_Click()
ChatName "(X)"
GetClipboard
Clipboard.SetData Image22.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image23_Click()
ChatName "(Z)"
GetClipboard
Clipboard.SetData Image23.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image24_Click()
ChatName "(I)"
GetClipboard
Clipboard.SetData Image24.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image25_Click()
ChatName "(H)"
GetClipboard
Clipboard.SetData Image25.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image26_Click()
ChatName "(S)"
GetClipboard
Clipboard.SetData Image26.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image27_Click()
ChatName "(*)"
GetClipboard
Clipboard.SetData Image27.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image28_Click()
ChatName "(%)"
GetClipboard
Clipboard.SetData Image28.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image29_Click()
ChatName "(8)"
GetClipboard
Clipboard.SetData Image29.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image30_Click()
ChatName "(H)"
GetClipboard
Clipboard.SetData Image30.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Image31_Click()
ChatName "(@)"
GetClipboard
Clipboard.SetData Image31.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub Label2_Click()
If MSN.LocalState = MSTATE_OFFLINE Then
MsgBox "Please Log-on to MSN", vbExclamation + vbApplicationModal, "Log on MSN"
Exit Sub
End If
If Check1.Value = 1 Then
MSN.Services.PrimaryService.FriendlyName = Chr(10) & Chr(13) & Chr(147) & Chr(10) & "- " & Text2.Text & "-" & Chr(147)
Else
MSN.Services.PrimaryService.FriendlyName = Text2.Text
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim l As ColorConstants
l = vbGreen
Label2.BackColor = l
End Sub

Private Sub Label3_Click()
ChatName vbCrLf
rt.SelText = vbCrLf
End Sub

Private Sub Label5_Click()
Dim res As VbMsgBoxResult
res = MsgBox("Are You Sure?", vbInformation + vbApplicationModal + vbYesNo, "Clear Name?")
If res = vbYes Then
rt.Text = ""
ChatName2 = ""
Text2.Text = ""
Text1.Text = ""
Else
End If
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim l As ColorConstants
l = vbGreen

Label5.BackColor = l
End Sub

Private Sub Label7_Click()
Dim res As VbMsgBoxResult
res = MsgBox("Are You Sure You Wanna Quit?", vbInformation + vbApplicationModal + vbYesNo, "Quit?")
If res = vbYes Then
End
Else
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If Len(Text1.Text) + Len(Text2.Text) >= 130 Then
Text1.MaxLength = Len(Text1.Text)
MsgBox "Name Can Only Be 130 Characters long", vbApplicationModal + vbInformation, "MSN Name Maker"
Else
Text1.MaxLength = 130
End If
Dim source1$, source2$, source3$, X%
'makes source1= something
source1 = "01²³456789"
'makes source2=something
source2 = "åb¢dèƒghîjklmñºÞq®$tµvw×ýz"
'makes source3=something
source3 = "ÁßÇÐÊFGH‡JK£MÑØ¶QR§TÚVWX¥Z"
'if checkbox is checked....
If Check2.Value = 1 Then
Select Case KeyAscii
Case Asc("0") To Asc("9")
X = KeyAscii - 47
KeyAscii = Asc(Mid(source1, X, 1))
Case Asc("A") To Asc("Z")
X = KeyAscii - 64
KeyAscii = Asc(Mid(source3, X, 1))
Case Asc("a") To Asc("z")
X = KeyAscii - 96
KeyAscii = Asc(Mid(source2, X, 1))
End Select
End If

If Check1.Value = 0 Then rt.SelColor = &H80000001 Else rt.SelColor = &H80000007

If KeyAscii = 13 Then
Clipboard.Clear
Clipboard.SetText Text1.Text
SendMessage rt.hwnd, WM_PASTE, 0, 0
ChatName Text1.Text
Text1.Text = ""
End If
End Sub

Private Sub Text2_click()
Text2.Text = rt.Text
End Sub

Function ChatName(Optional str As String) As String
If Len(ChatName2) >= 130 Then MsgBox "Name Can Only Be 130 Characters long", vbApplicationModal + vbInformation, "MSN Name Maker"
ChatName2 = ChatName2 & str
Text2.Text = ChatName2
End Function


