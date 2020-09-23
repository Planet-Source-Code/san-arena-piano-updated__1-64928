VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Armin Piano"
   ClientHeight    =   2295
   ClientLeft      =   3195
   ClientTop       =   4335
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   9300
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1620
      Top             =   15
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Record"
      Height          =   270
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   375
      Width           =   795
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Play"
      Height          =   270
      Left            =   4020
      TabIndex        =   74
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton Command66 
      BackColor       =   &H80000018&
      Caption         =   "Save"
      Height          =   270
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   375
      Width           =   795
   End
   Begin VB.CommandButton Command67 
      BackColor       =   &H80000018&
      Caption         =   "Load"
      Height          =   270
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   120
      Width           =   795
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -30
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Armin Piano"
      Filter          =   "*.apo|*.apo"
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   210
      Top             =   195
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   195
      Top             =   180
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFDE8&
      Caption         =   "Use Keyboard"
      Height          =   525
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin MSComctlLib.Slider vlm 
      Height          =   300
      Left            =   4920
      TabIndex        =   66
      Top             =   345
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      Max             =   127
      SelStart        =   127
      TickStyle       =   3
      TickFrequency   =   10
      Value           =   127
      TextPosition    =   1
   End
   Begin VB.CommandButton Command70 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   8625
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command69 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command65 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command64 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command63 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command62 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   7905
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command61 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command60 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   6465
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command59 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   6225
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command58 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   5985
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command57 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   5505
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command56 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   5265
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command55 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   4785
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command54 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command53 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   4305
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command52 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command51 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   3585
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command50 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   2625
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command47 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   2145
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command46 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   1905
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command43 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   765
      Width           =   195
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   7770
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   7290
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6810
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   570
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   750
      Width           =   255
   End
   Begin MSComctlLib.Slider chn 
      Height          =   300
      Left            =   5880
      TabIndex        =   68
      Top             =   345
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      Min             =   1
      Max             =   16
      SelStart        =   16
      TickStyle       =   3
      Value           =   16
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider ptc 
      Height          =   300
      Left            =   6870
      TabIndex        =   70
      Top             =   345
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      Min             =   1
      Max             =   50
      SelStart        =   20
      TickStyle       =   3
      Value           =   20
      TextPosition    =   1
   End
   Begin VB.TextBox rec 
      Height          =   405
      Left            =   4155
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   540
      Left            =   45
      TabIndex        =   79
      Top             =   135
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Piano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   1680
      TabIndex        =   76
      Top             =   135
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Armin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   585
      Left            =   120
      TabIndex        =   75
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pitch:"
      Height          =   225
      Left            =   7080
      TabIndex        =   71
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Channel:"
      Height          =   225
      Left            =   6000
      TabIndex        =   69
      Top             =   105
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume:"
      Height          =   225
      Left            =   5040
      TabIndex        =   67
      Top             =   105
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hmidi As Long
Dim baseNote
Dim channel
Dim volume
Dim lnote
Dim note
Dim Playin
Dim playinc
Dim timers
Dim recording
Dim pdemo

Private Sub Check1_GotFocus()
If Not Check1.Value = 1 Then Check1.Value = 1
End Sub

Private Sub Check1_Keydown(KeyCode As Integer, Shift As Integer)
domusickey KeyCode
End Sub

Private Sub Check1_KeyUp(KeyCode As Integer, Shift As Integer)
domusickeystop KeyCode
End Sub

Private Sub Check1_LostFocus()
If Check1.Value = 1 Then Check1.Value = 0
End Sub

Private Sub chn_Change()
channel = chn.Value - 1
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 1
End Sub

Private Sub Command10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 17
End Sub

Private Sub Command11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 18
End Sub

Private Sub Command12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 20
End Sub

Private Sub Command13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 22
End Sub

Private Sub Command14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 24
End Sub

Private Sub Command15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 25
End Sub

Private Sub Command16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 27
End Sub

Private Sub Command17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 29
End Sub

Private Sub Command18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 30
End Sub

Private Sub Command19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 32
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 3
End Sub

Private Sub Command20_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 34
End Sub

Private Sub Command21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 36
End Sub

Private Sub Command22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 37
End Sub

Private Sub Command23_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 39
End Sub

Private Sub Command24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 41
End Sub

Private Sub Command25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 42
End Sub

Private Sub Command26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 44
End Sub

Private Sub Command27_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 46
End Sub

Private Sub Command28_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 48
End Sub

Private Sub Command29_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 49
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 5
End Sub

Private Sub Command30_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 51
End Sub

Private Sub Command31_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 53
End Sub

Private Sub Command32_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 54
End Sub

Private Sub Command33_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 56
End Sub

Private Sub Command34_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 58
End Sub

Private Sub Command35_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 60
End Sub

Private Sub Command36_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 61
End Sub

Private Sub Command37_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 63
End Sub

Private Sub Command38_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 65
End Sub

Private Sub Command39_Click()
rec.Text = ""
recording = 1
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 6
End Sub

Private Sub Command40_Click()
recording = 0
Playin = Split(rec.Text, " ")
playinc = 0
Timer1.Interval = 1000
Timer1.Enabled = True
End Sub

Private Sub Command41_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 2
End Sub

Private Sub Command42_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 4
End Sub

Private Sub Command43_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 7
End Sub

Private Sub Command44_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 9
End Sub

Private Sub Command45_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 11
End Sub

Private Sub Command46_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 14
End Sub

Private Sub Command47_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 16
End Sub

Private Sub Command48_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 19
End Sub

Private Sub Command49_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 21
End Sub

Private Sub Command5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 8
End Sub

Private Sub Command50_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 23
End Sub

Private Sub Command51_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 26
End Sub

Private Sub Command52_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 28
End Sub

Private Sub Command53_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 31
End Sub

Private Sub Command54_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 33
End Sub

Private Sub Command55_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 35
End Sub

Private Sub Command56_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 38
End Sub

Private Sub Command57_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 40
End Sub

Private Sub Command58_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 43
End Sub

Private Sub Command59_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 45
End Sub

Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 10
End Sub

Private Sub Command60_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 47
End Sub

Private Sub Command61_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 59
End Sub

Private Sub Command62_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 57
End Sub

Private Sub Command63_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 55
End Sub

Private Sub Command64_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 52
End Sub

Private Sub Command65_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 50
End Sub

Private Sub Command66_Click()
CommonDialog1.ShowSave
If Not CommonDialog1.FileName = "" Then
f1 = FreeFile
Open CommonDialog1.FileName For Binary Access Write As #f1
Put #f1, , rec.Text
Close #f1
End If
End Sub

Private Sub Command67_Click()
domusic 20
CommonDialog1.ShowOpen
F = FreeFile
rec = ""
If Not CommonDialog1.FileName = "" Then
Open CommonDialog1.FileName For Input As F
rec = Input(LOF(F), F)
Close F
End If
domusic 60
End Sub

Private Sub Command69_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 64
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 12
End Sub

Private Sub Command70_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 62
End Sub

Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 13
End Sub

Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusicstop 15
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 1
End Sub

Private Sub Command10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 17
End Sub

Private Sub Command11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 18
End Sub

Private Sub Command12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 20
End Sub

Private Sub Command13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 22
End Sub

Private Sub Command14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 24
End Sub

Private Sub Command15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 25
End Sub

Private Sub Command16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 27
End Sub

Private Sub Command17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 29
End Sub

Private Sub Command18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 30
End Sub

Private Sub Command19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 32
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 3
End Sub

Private Sub Command20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 34
End Sub

Private Sub Command21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 36
End Sub

Private Sub Command22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 37
End Sub

Private Sub Command23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 39
End Sub

Private Sub Command24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 41
End Sub

Private Sub Command25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 42
End Sub

Private Sub Command26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 44
End Sub

Private Sub Command27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 46
End Sub

Private Sub Command28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 48
End Sub

Private Sub Command29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 49
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 5
End Sub

Private Sub Command30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 51
End Sub

Private Sub Command31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 53
End Sub

Private Sub Command32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 54
End Sub

Private Sub Command33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 56
End Sub

Private Sub Command34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 58
End Sub

Private Sub Command35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 60
End Sub

Private Sub Command36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 61
End Sub

Private Sub Command37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 63
End Sub

Private Sub Command38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 65
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 6
End Sub

Private Sub Command41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 2
End Sub

Private Sub Command42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 4
End Sub

Private Sub Command43_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 7
End Sub

Private Sub Command44_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 9
End Sub

Private Sub Command45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 11
End Sub

Private Sub Command46_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 14
End Sub

Private Sub Command47_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 16
End Sub

Private Sub Command48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 19
End Sub

Private Sub Command49_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 21
End Sub

Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 8
End Sub

Private Sub Command50_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 23
End Sub

Private Sub Command51_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 26
End Sub

Private Sub Command52_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 28
End Sub

Private Sub Command53_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 31
End Sub

Private Sub Command54_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 33
End Sub

Private Sub Command55_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 35
End Sub

Private Sub Command56_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 38
End Sub

Private Sub Command57_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 40
End Sub

Private Sub Command58_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 43
End Sub

Private Sub Command59_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 45
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 10
End Sub

Private Sub Command60_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 47
End Sub

Private Sub Command61_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 59
End Sub

Private Sub Command62_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 57
End Sub

Private Sub Command63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 55
End Sub

Private Sub Command64_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 52
End Sub

Private Sub Command65_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 50
End Sub

Private Sub Command69_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 64
End Sub

Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 12
End Sub

Private Sub Command70_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 62
End Sub

Private Sub Command8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 13
End Sub

Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
domusic 15
End Sub

Private Sub Form_Load()
rc = midiOutClose(hmidi)
rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
If (rc <> 0) Then
MsgBox "Couldn't open midi out, rc = " & rc
End If
pdemo = 0
baseNote = 20
channel = 15
volume = 127
timers = 0
recording = 0
End Sub

Function domusic(note)
playit note
End Function

Function domusicstop(note)
midimsg = &H80 + ((baseNote + note) * &H100) + channel
midiOutShortMsg hmidi, midimsg
lnote = 0
End Function

Function domusickey(note)
If note = vbKeyQ Then note = 2
If note = vbKeyW Then note = 4
If note = vbKeyE Then note = 6
If note = vbKeyR Then note = 8
If note = vbKeyT Then note = 10
If note = vbKeyY Then note = 12
If note = vbKeyU Then note = 14
If note = vbKeyI Then note = 16
If note = vbKeyO Then note = 18
If note = vbKeyP Then note = 20
If note = vbKeyA Then note = 22
If note = vbKeyS Then note = 24
If note = vbKeyD Then note = 26
If note = vbKeyF Then note = 28
If note = vbKeyG Then note = 30
If note = vbKeyH Then note = 32
If note = vbKeyJ Then note = 34
If note = vbKeyK Then note = 36
If note = vbKeyL Then note = 38
If note = vbKeyZ Then note = 40
If note = vbKeyX Then note = 42
If note = vbKeyC Then note = 44
If note = vbKeyV Then note = 46
If note = vbKeyB Then note = 48
If note = vbKeyN Then note = 50
If note = vbKeyM Then note = 52
If note = vbKey1 Then note = 54
If note = vbKey2 Then note = 56
If note = vbKey3 Then note = 58
If note = vbKey4 Then note = 60
If note = vbKey5 Then note = 62
If note = vbKey6 Then note = 64
If note = vbKey7 Then note = 66
If note = vbKey8 Then note = 68
If note = vbKey9 Then note = 70
If note = vbKey0 Then note = 72
If Not note = lnote Then
playit note
End If
lnote = note
End Function
Function domusickeystop(note)
If note = vbKeyQ Then note = 2
If note = vbKeyW Then note = 4
If note = vbKeyE Then note = 6
If note = vbKeyR Then note = 8
If note = vbKeyT Then note = 10
If note = vbKeyY Then note = 12
If note = vbKeyU Then note = 14
If note = vbKeyI Then note = 16
If note = vbKeyO Then note = 18
If note = vbKeyP Then note = 20
If note = vbKeyA Then note = 22
If note = vbKeyS Then note = 24
If note = vbKeyD Then note = 26
If note = vbKeyF Then note = 28
If note = vbKeyG Then note = 30
If note = vbKeyH Then note = 32
If note = vbKeyJ Then note = 34
If note = vbKeyK Then note = 36
If note = vbKeyL Then note = 38
If note = vbKeyZ Then note = 40
If note = vbKeyX Then note = 42
If note = vbKeyC Then note = 44
If note = vbKeyV Then note = 46
If note = vbKeyB Then note = 48
If note = vbKeyN Then note = 50
If note = vbKeyM Then note = 52
If note = vbKey1 Then note = 54
If note = vbKey2 Then note = 56
If note = vbKey3 Then note = 58
If note = vbKey4 Then note = 60
If note = vbKey5 Then note = 62
If note = vbKey6 Then note = 64
If note = vbKey7 Then note = 66
If note = vbKey8 Then note = 68
If note = vbKey9 Then note = 70
If note = vbKey0 Then note = 72
midimsg = &H80 + ((baseNote + note) * &H100) + channel
midiOutShortMsg hmidi, midimsg
If note = lnote Then lnote = 0
End Function

Private Sub Form_Terminate()
rc = midiOutClose(hmidi)
End Sub

Private Sub Form_Unload(Cancel As Integer)
rc = midiOutClose(hmidi)
End Sub
Function playit(note)
midimsg = &H90 + ((baseNote + note) * &H100) + (volume * &H10000) + channel
midiOutShortMsg hmidi, midimsg
If recording = 1 Then rec.Text = rec.Text & note & "x" & timers & " "
timers = 0
End Function

Private Sub Label6_Click()
Timer3.Enabled = True
End Sub

Private Sub ptc_Change()
baseNote = ptc.Value
End Sub

Private Sub Timer1_Timer()
On Error GoTo Errs
getnote = Split(Playin(playinc), "x")
nots = getnote(0)
midimsg = &H90 + ((baseNote + nots) * &H100) + (volume * &H10000) + channel
midiOutShortMsg hmidi, midimsg
playinc = playinc + 1
On Error Resume Next
getnotes = Split(Playin(playinc), "x")
Times = getnotes(1) * 10
Timer1.Interval = Times
Exit Sub
Errs:
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
timers = timers + 1
End Sub

Private Sub Timer3_Timer()
pdemo = pdemo + 20
domusic pdemo
If pdemo > 180 Then
pdemo = 0
Timer3.Enabled = False
End If
End Sub

Private Sub vlm_Change()
volume = vlm.Value
End Sub

