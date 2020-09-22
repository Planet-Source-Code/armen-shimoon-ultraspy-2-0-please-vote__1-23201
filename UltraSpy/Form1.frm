VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Form1.frx":0E42
   ScaleHeight     =   6105
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4770
      Top             =   5130
   End
   Begin VB.Frame Frame2 
      Caption         =   "Altering Information"
      Height          =   2670
      Left            =   90
      TabIndex        =   2
      Top             =   2250
      Width           =   5055
      Begin VB.TextBox txtStatic 
         Height          =   285
         Left            =   2430
         TabIndex        =   22
         Text            =   "<null>"
         Top             =   2070
         Width           =   2445
      End
      Begin VB.TextBox txtDis 
         Height          =   285
         Left            =   2430
         TabIndex        =   20
         Text            =   "<null>"
         Top             =   630
         Width           =   2445
      End
      Begin VB.TextBox txtKill 
         Height          =   285
         Left            =   2430
         TabIndex        =   18
         Text            =   "<null>"
         Top             =   1710
         Width           =   2445
      End
      Begin VB.TextBox txtClassText 
         Height          =   285
         Left            =   2430
         TabIndex        =   16
         Text            =   "<null>"
         Top             =   1350
         Width           =   2445
      End
      Begin VB.TextBox txtEnable 
         Height          =   285
         Left            =   2430
         TabIndex        =   15
         Text            =   "<null>"
         Top             =   270
         Width           =   2445
      End
      Begin VB.TextBox txtWinTxt 
         Height          =   285
         Left            =   2430
         TabIndex        =   14
         Text            =   "<null>"
         Top             =   990
         Width           =   2445
      End
      Begin VB.Label Label10 
         Caption         =   "Unmasked text to change:"
         Height          =   285
         Left            =   180
         TabIndex        =   21
         Top             =   2160
         Width           =   2265
      End
      Begin VB.Label Label9 
         Caption         =   "Disable window with name:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label8 
         Caption         =   "Window name to kill:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1800
         Width           =   1545
      End
      Begin VB.Label Label7 
         Caption         =   "Change window text of class:"
         Height          =   330
         Left            =   180
         TabIndex        =   13
         Top             =   1440
         Width           =   2325
      End
      Begin VB.Label Label6 
         Caption         =   "Change text of window text:"
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   1080
         Width           =   2265
      End
      Begin VB.Label Label3 
         Caption         =   "Enable window with name:"
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "UltraSpy Information"
      Height          =   1995
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5055
      Begin VB.TextBox txtParent 
         Height          =   285
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   2265
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   1710
         TabIndex        =   7
         Top             =   720
         Width           =   2265
      End
      Begin VB.TextBox txtClass 
         Height          =   285
         Left            =   1710
         TabIndex        =   6
         Top             =   360
         Width           =   2265
      End
      Begin VB.TextBox txtUM 
         Height          =   285
         Left            =   1710
         TabIndex        =   5
         Top             =   1440
         Width           =   2265
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4230
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "Parent:"
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   1170
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "Window Text:"
         Height          =   180
         Left            =   630
         TabIndex        =   11
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Class:"
         Height          =   270
         Left            =   1170
         TabIndex        =   10
         Top             =   450
         Width           =   450
      End
      Begin VB.Label Label5 
         Caption         =   "Unmasked Text:"
         Height          =   195
         Left            =   450
         TabIndex        =   9
         Top             =   1530
         Width           =   1275
      End
   End
   Begin VB.Image banner 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   450
      Top             =   5040
      Width           =   4245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim verylong As String * 100
Dim gParent As String * 100
Dim SndMsg As String * 100
Dim windowname As String * 100
Dim sztext As String * 100
Dim mousemove As Boolean
Dim Pic01 As Boolean




Private Sub Form_Load()
Picture1.Picture = LoadResPicture(102, vbResIcon)
Form1.Icon = LoadResPicture(101, vbResIcon)
Form1.Caption = LoadResString(101) & " " & LoadResString(102) & " v2.0"
mousemove = False
TextRO txtClass
TextRO txtText
TextRO txtParent
TextRO txtUM
banner.Picture = LoadResPicture(105, vbResBitmap)
Pic01 = True
banner.Left = (Form1.Width - banner.Width) \ 2
Form1.Height = (banner.Top + banner.Height) + 500
KeepOnTop Form1

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = Nothing
Form1.MouseIcon = LoadResPicture(102, vbResIcon)
Form1.MousePointer = 99
mousemove = True

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim cursorpos1 As POINTAPI
   Dim wintext As String
   Dim garmon As String
   Dim gIcon As Image
   Dim OldX As Integer
   Dim OldY As Integer
   Dim ttxt As String
   Dim abc As String
 If mousemove = True Then
 
    r = GetCursorPos(cursorpos1)
    hwnd1 = WindowFromPoint(cursorpos1.X, cursorpos1.Y)
    r = GetClassName(hwnd1, sztext, 100)
    hwnd2 = WindowFromPoint(cursorpos1.X, cursorpos1.Y)
    p = GetWindowText(hwnd2, windowname, 100)
    hwnd3 = WindowFromPoint(cursorpos1.X, cursorpos1.Y)
    q = GetParent(hwnd3)
    
            
              ttxt = Space(100)
              errval = GetCursorPos(cursorpos1)
              thwnd = WindowFromPoint(cursorpos1.X, cursorpos1.Y)
              errval = SendMessage(thwnd, WM_GETTEXT, ByVal TXT_LEN, ByVal ttxt)
              ttxt = RTrim(ttxt)
              txtUM.Text = ttxt
     
    txtParent.Text = q
    txtText.Text = windowname
    txtClass.Text = sztext


If txtText.Text = txtEnable.Text Then
    a = EnableWindow(hwnd1, 1)
    txtEnable.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtText.Text = txtWinTxt.Text Then
    a = InputBox("New string for " & txtText & ":", "New string")
    b = SetWindowText(hwnd2, a)
    txtWinTxt.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtClass.Text = txtClassText.Text Then
    a = InputBox("New string for " & txtClass & ":", "New string")
    b = SetWindowText(hwnd1, a)
    txtClassText.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtKill.Text = txtText.Text Then
    a = CloseWindow(hwnd2)
    txtKill.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtDis.Text = txtText.Text Then
    a = EnableWindow(hwnd1, 0)
    txtDis.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtUM.Text = txtStatic.Text Then
    abc = InputBox("New string for static " & txtUM.Text, "New string")
    Call SendMessage(hwnd2, WM_SETTEXT, 0&, ByVal abc)
    txtStatic.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
End If
 
 
    
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = LoadResPicture(102, vbResIcon)
Form1.MousePointer = 0
mousemove = False
End Sub


Private Function TextRO(textbx As TextBox)
a = SendMessage(textbx.hwnd, EM_SETREADONLY, 1, 0)
End Function

Private Sub Timer1_Timer()
If Pic01 = True Then
    banner.Picture = LoadResPicture(106, vbResBitmap)
    Pic01 = False
Else
    banner.Picture = LoadResPicture(105, vbResBitmap)
    Pic01 = True
End If
End Sub

Public Function AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)


    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Function


