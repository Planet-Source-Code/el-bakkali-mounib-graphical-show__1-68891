VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8250
      Left            =   200
      TabIndex        =   3
      Top             =   540
      Width           =   1395
      Begin VB.OptionButton Option12 
         BackColor       =   &H80000012&
         Caption         =   "Mode 12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   5310
         Width           =   1155
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H80000012&
         Caption         =   "Mode 11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   135
         TabIndex        =   18
         Top             =   4980
         Width           =   1155
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00000000&
         Caption         =   "Mode 10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   4635
         Width           =   1125
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00000000&
         Caption         =   "Mode 9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   135
         TabIndex        =   14
         Top             =   4245
         Width           =   1125
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00000000&
         Caption         =   "Mode 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   3870
         Width           =   1125
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00000000&
         Caption         =   "Mode 7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   135
         TabIndex        =   12
         Top             =   3450
         Width           =   1140
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000012&
         Caption         =   "Mode 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   3045
         Width           =   1140
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00000000&
         Caption         =   "Mode 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   135
         TabIndex        =   10
         Top             =   2595
         Width           =   1125
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000012&
         Caption         =   "Mode 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   2175
         Width           =   1110
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Mode 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   135
         TabIndex        =   8
         Top             =   1770
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Mode 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   1260
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000012&
         Caption         =   "Mode 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Exit"
         Height          =   300
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   105
         TabIndex        =   17
         Top             =   5610
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6060
      Left            =   -6650
      ScaleHeight     =   6030
      ScaleWidth      =   6405
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   6435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   0
      Top             =   15
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   285
      Top             =   -75
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   3
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   8595
      Left            =   45
      ScaleHeight     =   8565
      ScaleWidth      =   13590
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   13620
      Begin VB.Timer Timer3 
         Interval        =   2
         Left            =   2115
         Top             =   120
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   285
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t(8) As Integer
Dim X As Integer
Dim Y As Integer
Public df As Integer

Sub main()

 fomr2.Show
 
End Sub


Private Sub Check1_Click()
 Picture1.Cls
End Sub

Private Sub Command1_Click()

 On Error Resume Next
 galaxie.BackColor = 0
galaxie.ForeColor = 10000


 For n = 1 To 12000 Step 12
For X = 1 To 20
Me.ForeColor = Me.ForeColor + 10
        
 
Line (Cos(n * X), 200 * X * n)-(3000 * X, Sin(X))
Circle (n * 3, X * 2), n * 2 + X * 20


 DoEvents
Next X
Next n


End Sub

Private Sub Command2_Click()
 Timer1.Enabled = True
End Sub




Private Sub Command5_Click()
   Dim s As Integer
   Timer1.Enabled = False
   Timer2.Enabled = False
   
 For s = 20 To -1300 Step -44
     Frame1.Move s, Frame1.Top, Frame1.Width, Frame1.Height
     DoEvents
     Next s
  
   End
End Sub

Private Sub Form_Activate()
 On Error Resume Next

Me.BackColor = 0
Me.ForeColor = 10000
Picture1.Visible = True
Timer2.Enabled = True
End Sub

Private Sub Form_Click()
'   Unload Me
End Sub

Private Sub Form_Load()


        Picture2.Width = Picture1.Width
        Picture2.Height = Picture1.Height
        Picture2.Top = Picture1.Top
        Picture2.Left = Picture1.Left
        Frame1.Height = Picture1.Height
        Picture1.Top = 0
        Picture1.Height = Me.Height + 200
        Picture1.Width = Me.Width
        X = 5200
        Y = 2600
        k = 0
        
  For i = 0 To 7
        t(i) = k
        k = k + 80
  Next
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
 Timer1.Enabled = False
 Timer2.Enabled = False

End Sub

       
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim s As Integer
    
    For s = -1300 To 20 Step 44
       Frame1.Move s, Frame1.Top, Frame1.Width, Frame1.Height
       DoEvents
    Next s
    
 
End Sub

Private Sub Option1_Click()
df = 0
   Picture1.Cls
End Sub

Private Sub Option10_Click()
df = 0
 Picture1.Cls
End Sub

Private Sub Option11_Click()
 df = 0
 Picture1.Cls
End Sub

Private Sub Option12_Click()
  df = 0
 Picture1.Cls
End Sub

Private Sub Option2_Click()
df = 0
   Picture1.Cls
   
End Sub

Private Sub Option3_Click()
df = 0
   Picture1.Cls
End Sub

Private Sub Option4_Click()
df = 0
 Picture1.Cls
End Sub

Private Sub Option5_Click()
df = 0
 Picture1.Cls
End Sub

Private Sub Option6_Click()
df = 0
 Picture1.Cls
End Sub

Private Sub Option7_Click()
df = 0
 Picture1.Cls
End Sub

Private Sub Option8_Click()
df = 0
 Picture1.Cls
End Sub

Private Sub Option9_Click()
df = 0
 Picture1.Cls
End Sub

Private Sub Picture1_Click()
     Dim s As Integer
    
    For s = 20 To -1300 Step -44
       Frame1.Move s, Frame1.Top, Frame1.Width, Frame1.Height
       DoEvents
    Next s
End Sub

Private Sub timer1_timer()
Form1.Cls
' Picture1.Cls
End Sub

Sub fun()
 On Error Resume Next
 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
 Picture1.Width = Screen.Width
X = 6100
Y = 4450
k = 0


For i = 0 To 7 Step 10
df = df + 1

k = i + 10

If k > 7 Then k = 0
Picture1.FillColor = vbBlack

Picture1.Circle (X - 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 600, vbRed, , , 1
Picture1.Circle (X + 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 600, vbRed, , , 1
Picture1.Circle (X - 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 600, vbRed, , 0.2
Picture1.Circle (X + 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 600, vbRed, , 0.2


Picture1.Circle (X - 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y - 3500 * Cos(t(i) * 63.1415 / 4580)), 600, vbRed, , , 1
Picture1.Circle (X + 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y - 3500 * Cos(t(i) * 63.1415 / 4580)), 600, vbRed, , , 1
Picture1.Circle (X - 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y - 3500 * Cos(t(i) * 63.1415 / 4580)), 600, vbRed, , 0.2
Picture1.Circle (X + 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 600, vbRed, , 0.2

Picture1.FillColor = &H80FFFF
Picture1.Circle (X - 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , , 1
Picture1.Circle (X + 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , , 1
Picture1.Circle (X - 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , 0.2
Picture1.Circle (X + 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , 0.2


Picture1.Circle (X - 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y - 3500 * Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , , 1
Picture1.Circle (X + 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y - 3500 * Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , , 1
Picture1.Circle (X - 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y - 3500 * Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , 0.2
Picture1.Circle (X + 4500 * Sin(Tan(t(i) * 3.1415 / 3480)), Y + 3500 * Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , 0.2

Picture1.FillColor = vbBlack




Picture1.Circle (X - 4500 * Sin(t(i) * 3.1415 / 1480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 100, &H8000&, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 1480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 100, &H8000&, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 3.1415 / 1480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 100, &H8000&, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 1480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 100, &H8000&, , 0.2


Picture1.Circle (X - 4500 * Sin(t(i) * 3.1415 / 1480), Y - 3500 * Cos(t(i) * 3.1415 / 580)), 100, &H8000&, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 1480), Y - 3500 * Cos(t(i) * 3.1415 / 580)), 100, &H8000&, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 3.1415 / 1480), Y - 3500 * Cos(t(i) * 3.1415 / 580)), 100, &H8000&, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 1480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 100, &H8000&, , 0.2

Picture1.Circle (X - 4500 * Sin(t(i) * 3.1415 / 480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 200, vbYellow, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 200, vbYellow, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 3.1415 / 480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 200, vbYellow, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 200, vbYellow, , 0.2


Picture1.Circle (X - 4500 * Sin(t(i) * 3.1415 / 480), Y - 3500 * Cos(t(i) * 3.1415 / 580)), 200, vbYellow, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 480), Y - 3500 * Cos(t(i) * 3.1415 / 580)), 200, vbYellow, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 3.1415 / 480), Y - 3500 * Cos(t(i) * 3.1415 / 580)), 200, vbYellow, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 480), Y + 3500 * Cos(t(i) * 3.1415 / 580)), 200, vbYellow, , 0.2
 Label1.Caption = df
 If df >= 11000 Then df = 0: Option2.Value = True: Exit For
Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 0
For i = 0 To 7
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next


End Sub
Private Sub Timer2_Timer()

        If Option1.Value = True Then
        fun
        End If
        
        If Option2.Value = True Then
        
        
        fun1
        End If
        
        
        
        
        If Option3.Value = True Then
        fun3
        End If
        
        If Option4.Value = True Then
        fun4
        End If
        
        If Option5.Value = True Then
        fun5
        End If
        If Option6.Value = True Then
        fun6
        End If
        If Option7.Value = True Then
        fun7
        End If
        If Option8.Value = True Then
        fun8
        End If
        If Option9.Value = True Then
        fun9
        End If
        If Option10.Value = True Then
        fun10
        End If
        
        If Option11.Value = True Then
        fun11
        End If
 
If Option12.Value = True Then
        fun13
        End If
          
End Sub
Sub fun1()

 On Error GoTo Mounib
 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
 Picture1.Width = Screen.Width
 
 
 
X = 6100
Y = 4450
k = 0


For i = 0 To 7 Step 10
 df = df + 1
k = i + 1
If k > 7 Then k = 0


Picture1.FillColor = vbBlack
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 1000, &H8000&, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 1000, &H8000&, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 1000, &H8000&, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 1000, &H8000&, , 0.2


Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y - 100 / Cos(t(i) * 63.1415 / 4580)), 1000, &H8000&, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y - 100 / Cos(t(i) * 63.1415 / 4580)), 1000, &H8000&, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y - 100 / Cos(t(i) * 63.1415 / 4580)), 1000, &H8000&, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 1000, &H8000&, , 0.2

Picture1.FillColor = vbYellow
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , 0.2


Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y - 100 / Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y - 100 / Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y - 100 / Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y + 100 / Cos(t(i) * 63.1415 / 4580)), 300, vbYellow, , 0.2


Picture1.FillColor = vbBlack
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Tan(t(i) * 73.1415 / 580)), 100, vbYellow, , , 1
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Tan(t(i) * 73.1415 / 580)), 100, vbYellow, , , 1
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Tan(t(i) * 73.1415 / 580)), 100, vbYellow, , 0.2
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Tan(t(i) * 73.1415 / 580)), 100, vbYellow, , 0.2

Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y - 1500 * Tan(t(i) * 73.1415 / 580)), 100, vbYellow, , , 1
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y - 1500 * Tan(t(i) * 73.1415 / 580)), 100, vbYellow, , , 1
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y - 1500 * Tan(t(i) * 73.1415 / 580)), 100, vbYellow, , 0.2
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Tan(t(i) * 73.1415 / 580)), 100, vbYellow, , 0.2


Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 480), Y + 1500 * Sin(t(i) * 3.1415 / 580)), 800, vbYellow, , , 1
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 480), Y + 1500 * Sin(t(i) * 3.1415 / 580)), 800, vbYellow, , , 1
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 480), Y + 1500 * Sin(t(i) * 3.1415 / 580)), 800, vbYellow, , 0.2
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 480), Y + 1500 * Sin(t(i) * 3.1415 / 580)), 800, vbYellow, , 0.2

Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 480), Y - 1500 * Sin(t(i) * 3.1415 / 580)), 800, vbYellow, , , 1
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 480), Y - 1500 * Sin(t(i) * 3.1415 / 580)), 800, vbYellow, , , 1
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 480), Y - 1500 * Sin(t(i) * 3.1415 / 580)), 800, vbYellow, , 0.2
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 480), Y + 1500 * Sin(t(i) * 3.1415 / 580)), 800, vbYellow, , 0.2




 If df >= 11000 Then df = 0: Option3.Value = True: Exit For

Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 0
For i = 0 To 7
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next

Mounib: bab

  

End Sub
Sub bab()
   Me.Refresh

  Timer2.Enabled = True
 End Sub
  


Sub fun3()
  On Error Resume Next
 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
' If Check1.Value = vbChecked Then
' Picture1.Cls
'
'   End If
'   Check1.Value = vbUnchecked
X = 6400
Y = 4450
k = 0


For i = 0 To 7 Step 10

k = i + 10
If k > 7 Then k = 0


Next

For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 10
For i = 5 To 7 Step 10
df = df + 1
Picture1.Circle (X - 2500 * Sin(Tan(Cos(t(i) / Tan(k * 73.1415) / 750))), Y - 2500 * Sin(Tan(Tan(Tan(t(i) * 44.1415 / 6340))))), 400, vbWhite, , , 5
Picture1.Circle (X - 2500 * Sin(Tan(Cos(t(i) / Tan(k * 73.1415) / 750))), Y + 2500 * Sin(Tan(Tan(Tan(t(i) * 44.1415 / 6340))))), 400, &H80FFFF, , , 5
Picture1.Circle (X - 2500 * Sin(Tan(Cos(t(i) / Tan(k * 73.1415) / 750))), Y - 2500 * Sin(Tan(Tan(Tan(t(i) * 44.1415 / 6340))))), 400, &H8000&, , , 0.425
Picture1.Circle (X - 2500 * Sin(Tan(Cos(t(i) / Tan(k * 73.1415) / 750))), Y + 2500 * Sin(Tan(Tan(Tan(t(i) * 44.1415 / 6340))))), 400, &H8000&, , , 0.425


Picture1.Circle (X + 2500 * Sin(Tan(Cos(t(i) / Tan(k * 73.1415) / 750))), Y - 2500 * Sin(Tan(Tan(Tan(t(i) * 44.1415 / 6340))))), 400, vbWhite, , , 5
Picture1.Circle (X + 2500 * Sin(Tan(Cos(t(i) / Tan(k * 73.1415) / 750))), Y + 2500 * Sin(Tan(Tan(Tan(t(i) * 44.1415 / 6340))))), 400, &H80FFFF, , , 5
Picture1.Circle (X + 2500 * Sin(Tan(Cos(t(i) / Tan(k * 73.1415) / 750))), Y - 2500 * Sin(Tan(Tan(Tan(t(i) * 44.1415 / 6340))))), 400, &H8000&, , , 0.425
Picture1.Circle (X + 2500 * Sin(Tan(Cos(t(i) / Tan(k * 73.1415) / 750))), Y + 2500 * Sin(Tan(Tan(Tan(t(i) * 44.1415 / 6340))))), 400, &H8000&, , , 0.425


Picture1.Circle (X + 2500 * Tan(Sin(t(i) / k * 73.1415 / 1080)), Y - 2500 / Sin(Tan(t(i) * 45.1415 / 10340))), 200, vbYellow, , , 5
Picture1.Circle (X + 2500 * Tan(Sin(t(i) / k * 73.1415 / 1080)), Y + 2500 / Sin(Tan(t(i) * 45.1415 / 10340))), 200, vbYellow, , , 5
Picture1.Circle (X + 2500 * Tan(Sin(t(i) / k * 73.1415 / 1080)), Y - 2500 / Sin(Tan(t(i) * 45.1415 / 10340))), 200, &H8000&, , , 0.425
Picture1.Circle (X + 2500 * Tan(Sin(t(i) / k * 73.1415 / 1080)), Y + 2500 / Sin(Tan(t(i) * 45.1415 / 10340))), 200, &H8000&, , , 0.425

Picture1.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 1080)), Y - 2500 / Sin(Tan(t(i) * 45.1415 / 10340))), 200, vbYellow, , , 5
Picture1.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 1080)), Y + 2500 / Sin(Tan(t(i) * 45.1415 / 10340))), 200, vbYellow, , , 5
Picture1.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 1080)), Y - 2500 / Sin(Tan(t(i) * 45.1415 / 10340))), 200, &H8000&, , , 0.425
Picture1.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 1080)), Y + 2500 / Sin(Tan(t(i) * 45.1415 / 10340))), 200, &H8000&, , , 0.425



If df >= 11000 Then df = 0: Option4.Value = True: Exit For


k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next



End Sub

Sub fun4()
 On Error Resume Next
 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
 Picture1.Width = Screen.Width
 
X = 6000
Y = 4450
k = 0


For i = 0 To 7 Step 10
df = df + 1
k = i + 10
If k > 7 Then k = 0


Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 10
For i = 5 To 7 Step 10
'

df = df + 1
Picture1.Circle (X + 2500 * Sin(Sin(Sin(t(i) / k * 73.1415 / 580))), Y - 2500 / Tan(t(i) * 45.1415 / 7340)), 200, vbYellow, , , 5
Picture1.Circle (X + 2500 * Sin(Sin(Sin(t(i) / k * 73.1415 / 580))), Y + 2500 / Tan(t(i) * 45.1415 / 7340)), 200, vbYellow, , , 5
Picture1.Circle (X + 2500 * Sin(Sin(Sin(t(i) / k * 73.1415 / 580))), Y - 2500 / Tan(t(i) * 45.1415 / 7340)), 200, &H8000&, , , 0.425
Picture1.Circle (X + 2500 * Sin(Sin(Sin(t(i) / k * 73.1415 / 580))), Y + 2500 / Tan(t(i) * 45.1415 / 7340)), 200, &H8000&, , , 0.425

Picture1.Circle (X - 2500 * Sin(Sin(Sin(t(i) / k * 73.1415 / 580))), Y - 2500 / Tan(t(i) * 45.1415 / 7340)), 200, vbYellow, , , 5
Picture1.Circle (X - 2500 * Sin(Sin(Sin(t(i) / k * 73.1415 / 580))), Y + 2500 / Tan(t(i) * 45.1415 / 7340)), 200, vbYellow, , , 5
Picture1.Circle (X - 2500 * Sin(Sin(Sin(t(i) / k * 73.1415 / 580))), Y - 2500 / Tan(t(i) * 45.1415 / 7340)), 200, &H8000&, , , 0.425
Picture1.Circle (X - 2500 * Sin(Sin(Sin(t(i) / k * 73.1415 / 580))), Y + 2500 / Tan(t(i) * 45.1415 / 7340)), 200, &H8000&, , , 0.425



Picture1.Circle (X + 2500 * Tan(Cos(Sin(Sin(t(i) / k * 73.1415 / 580)))), Y - 2500 / Tan(t(i) * 45.1415 / 7340)), 200, vbYellow, , , 5
Picture1.Circle (X + 2500 * Tan(Cos(Sin(Sin(t(i) / k * 73.1415 / 580)))), Y + 2500 / Tan(t(i) * 45.1415 / 7340)), 200, vbYellow, , , 5
Picture1.Circle (X + 2500 * Tan(Cos(Sin(Sin(t(i) / k * 73.1415 / 580)))), Y - 2500 / Tan(t(i) * 45.1415 / 7340)), 200, &H8000&, , , 0.425
Picture1.Circle (X + 2500 * Tan(Cos(Sin(Sin(t(i) / k * 73.1415 / 580)))), Y + 2500 / Tan(t(i) * 45.1415 / 7340)), 200, &H8000&, , , 0.425


Picture1.Circle (X - 2500 * Tan(Cos(Sin(Sin(t(i) / k * 73.1415 / 580)))), Y - 2500 / Tan(t(i) * 45.1415 / 7340)), 200, vbYellow, , , 5
Picture1.Circle (X - 2500 * Tan(Cos(Sin(Sin(t(i) / k * 73.1415 / 580)))), Y + 2500 / Tan(t(i) * 45.1415 / 7340)), 200, vbYellow, , , 5
Picture1.Circle (X - 2500 * Tan(Cos(Sin(Sin(t(i) / k * 73.1415 / 580)))), Y - 2500 / Tan(t(i) * 45.1415 / 7340)), 200, &H8000&, , , 0.425
Picture1.Circle (X - 2500 * Tan(Cos(Sin(Sin(t(i) / k * 73.1415 / 580)))), Y + 2500 / Tan(t(i) * 45.1415 / 7340)), 200, &H8000&, , , 0.425





If df >= 11000 Then df = 0: Option5.Value = True: Exit For

k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next



End Sub
Sub fun5()
 On Error Resume Next
 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
 Picture1.Width = Screen.Width
X = 6400
Y = 4450
k = 0


For i = 0 To 7 Step 10

k = i + 10
If k > 7 Then k = 0
df = df + 1


'
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 100, &H8000&, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 100, &H8000&, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 100, &H8000&, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 100, &H8000&, , 0.2


Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y - 500 / Cos(t(i) * 63.1415 / 4580)), 100, &H8000&, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y - 500 / Cos(t(i) * 63.1415 / 4580)), 100, &H8000&, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y - 500 / Cos(t(i) * 63.1415 / 4580)), 100, &H8000&, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 100, &H8000&, , 0.2


Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , , 1
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , , 1
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , 0.2
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , 0.2

Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y - 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , , 1
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y - 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , , 1
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y - 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , 0.2
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , 0.2

If df >= 11000 Then df = 0: Option6.Value = True: Exit For
Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 0
For i = 0 To 7
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next


End Sub
Sub fun6()
 On Error Resume Next


 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
  Picture1.Width = Screen.Width
X = 6000
Y = 4450
k = 0


For i = 0 To 7 Step 10

k = i + 10
If k > 7 Then k = 0
df = df + 1

Picture1.Circle (X - 3500 * Cos(t(i / 20) * 23.1415 / 1480), Y - 3500 * Tan(t(i) * 3.1415 / 880)), 500, &HFFFF&
Picture1.Circle (X + 3500 * Cos(t(i / 20) * 23.1415 / 1480), Y + 3500 * Tan(t(i) * 3.1415 / 880)), 500, vbRed


Picture1.Circle (X - 2500 * Tan(t(i) * 3.1415 / 480), Y - 2500 * Cos(t(i) * 65.1415 / 3340)), 500, &HC0FFFF, , , 0.425
Picture1.Circle (X + 2500 * Tan(t(i) * 3.1415 / 480), Y - 2500 * Cos(t(i) * 65.1415 / 3340)), 500, vbYellow, , , 0.425
Picture1.Circle (X - 2500 * Tan(t(i) * 3.1415 / 480), Y + 2500 * Cos(t(i) * 65.1415 / 3340)), 500, &HFF&, , , 0.425
Picture1.Circle (X + 2500 * Tan(t(i) * 3.1415 / 480), Y + 2500 * Cos(t(i) * 65.1415 / 3340)), 500, &H8000&, , , 0.425

If df >= 11000 Then df = 0: Option7.Value = True: Exit For


Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 0
For i = 0 To 7
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next


End Sub
Sub fun7()
 On Error Resume Next
 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
  Picture1.Width = Me.Width + 600
X = 6000
Y = 4450
k = 0


For i = 0 To 7 Step 10

k = i + 10
If k > 7 Then k = 0

df = df + 1

Picture1.Circle (X + 2500 * Cos(t(i) * 3.1415 / 208), Y - 2500 * Tan(t(i) * 55.1415 / 3340)), 500, &HC0FFFF, , , 5
Picture1.Circle (X + 2500 * Cos(t(i) * 3.1415 / 208), Y + 2500 * Tan(t(i) * 55.1415 / 3340)), 500, vbYellow, , , 5
Picture1.Circle (X + 2500 * Cos(t(i) * 3.1415 / 208), Y - 2500 * Tan(t(i) * 55.1415 / 3340)), 500, &HFF&, , , 0.425
Picture1.Circle (X + 2500 * Cos(t(i) * 3.1415 / 208), Y + 2500 * Tan(t(i) * 55.1415 / 3340)), 500, &H8000&, , , 0.425


Picture1.Circle (X - 2500 * Cos(t(i) * 3.1415 / 208), Y - 2500 * Tan(t(i) * 55.1415 / 3340)), 500, &HC0FFFF, , , 5
Picture1.Circle (X - 2500 * Cos(t(i) * 3.1415 / 208), Y + 2500 * Tan(t(i) * 55.1415 / 3340)), 500, vbYellow, , , 5
Picture1.Circle (X - 2500 * Cos(t(i) * 3.1415 / 208), Y - 2500 * Tan(t(i) * 55.1415 / 3340)), 500, &HFF&, , , 0.425
Picture1.Circle (X - 2500 * Cos(t(i) * 3.1415 / 208), Y + 2500 * Tan(t(i) * 55.1415 / 3340)), 500, &H8000&, , , 0.425



Picture1.Circle (X - 2500 * Tan(t(i) * 3.1415 / 2080), Y - 2500 * Cos(t(i) * 5.1415 / 340)), 600, &HC0FFFF, , , 0.425
Picture1.Circle (X + 2500 * Tan(t(i) * 3.1415 / 2080), Y - 2500 * Cos(t(i) * 5.1415 / 340)), 600, vbYellow, , , 0.425
Picture1.Circle (X - 2500 * Tan(t(i) * 3.1415 / 2080), Y + 2500 * Cos(t(i) * 5.1415 / 340)), 600, &HFF&, , , 0.425
Picture1.Circle (X + 2500 * Tan(t(i) * 3.1415 / 2080), Y + 2500 * Cos(t(i) * 5.1415 / 340)), 600, &H8000&, , , 0.425


If df >= 11000 Then df = 0: Option8.Value = True: Exit For

t(i) = t(i) + 2

Next

k = 0
For i = 0 To 7
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next



End Sub
Sub fun8()
 On Error Resume Next
 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
 
X = 6400
Y = 4450
k = 0


For i = 0 To 7 Step 10

k = i + 10
If k > 7 Then k = 0





Next
'For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
''Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
'T(i) = T(i)
'Next
k = 10
For i = 5 To 7 Step 10
 df = df + 1
Picture1.Circle (X - 2500 * Cos(t(i) / k * 73.1415 / 2080), Y - 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, vbYellow, , , 5
Picture1.Circle (X - 2500 * Cos(t(i) / k * 73.1415 / 2080), Y + 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, vbYellow, , , 5
Picture1.Circle (X - 2500 * Cos(t(i) / k * 73.1415 / 2080), Y - 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, &H8000&, , , 0.425
Picture1.Circle (X - 2500 * Cos(t(i) / k * 73.1415 / 2080), Y + 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, &H8000&, , , 0.425

Picture1.Circle (X + 2500 * Cos(t(i) / k * 73.1415 / 2080), Y - 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, vbYellow, , , 5
Picture1.Circle (X + 2500 * Cos(t(i) / k * 73.1415 / 2080), Y + 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, vbYellow, , , 5
Picture1.Circle (X + 2500 * Cos(t(i) / k * 73.1415 / 2080), Y - 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, &H8000&, , , 0.425
Picture1.Circle (X + 2500 * Cos(t(i) / k * 73.1415 / 2080), Y + 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, &H8000&, , , 0.425

Picture1.Circle (X - 2500 * Sin(t(i) / k * 73.1415 / 2080), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture1.Circle (X - 2500 * Sin(t(i) / k * 73.1415 / 2080), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture1.Circle (X - 2500 * Sin(t(i) / k * 73.1415 / 2080), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425
Picture1.Circle (X - 2500 * Sin(t(i) / k * 73.1415 / 2080), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425

Picture1.Circle (X + 2500 * Sin(t(i) / k * 73.1415 / 2080), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture1.Circle (X + 2500 * Sin(t(i) / k * 73.1415 / 2080), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture1.Circle (X + 2500 * Sin(t(i) / k * 73.1415 / 2080), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425
Picture1.Circle (X + 2500 * Sin(t(i) / k * 73.1415 / 2080), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425

Picture1.Circle (X - 2500 * Sin(t(i) / k * 73.1415 / 2080), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture1.Circle (X - 2500 * Sin(t(i) / k * 73.1415 / 2080), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture1.Circle (X - 2500 * Sin(t(i) / k * 73.1415 / 2080), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425
Picture1.Circle (X - 2500 * Sin(t(i) / k * 73.1415 / 2080), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425
t(i) = t(i) + 2

If df >= 11000 Then df = 0: Option9.Value = True: Exit For
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
'For i = 0 To 7
'Picture1.DrawWidth = 1
''Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
''Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
'
'Next


End Sub
Sub fun9()
 On Error Resume Next
 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
 Picture1.Width = Screen.Width
X = 6400
Y = 4450
k = 0


For i = 0 To 7 Step 10

k = i + 10
If k > 7 Then k = 0

 df = df + 1


Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 300, &H8000&, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 300, &H8000&, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 300, &H8000&, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 3.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 300, &H8000&, , 0.2


Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y - 500 / Cos(t(i) * 63.1415 / 4580)), 300, &H8000&, , , 1
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y - 500 / Cos(t(i) * 63.1415 / 4580)), 300, &H8000&, , , 1
Picture1.Circle (X - 4500 * Sin(t(i) * 73.1415 / 3480), Y - 500 / Cos(t(i) * 63.1415 / 4580)), 300, &H8000&, , 0.2
Picture1.Circle (X + 4500 * Sin(t(i) * 73.1415 / 3480), Y + 500 / Cos(t(i) * 63.1415 / 4580)), 300, &H8000&, , 0.2


Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y + 3500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , , 1
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 3500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , , 1
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y + 3500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , 0.2
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 1500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , 0.2

Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y - 3500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , , 1
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y - 3500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , , 1
Picture1.Circle (X - 3500 * Sin(t(i) * 3.1415 / 2480), Y - 3500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , 0.2
Picture1.Circle (X + 3500 * Sin(t(i) * 3.1415 / 2480), Y + 3500 * Cos(t(i) * 33.1415 / 580)), 500, vbYellow, , 0.2

If df >= 11000 Then df = 0: Option10.Value = True: Exit For
Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 0
For i = 0 To 7
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next


End Sub
Sub fun10()
 On Error Resume Next


 Picture1.Top = 200
 Picture1.Height = Me.Height + 200
 
X = 6400
Y = 4450
k = 0


For i = 0 To 7 Step 10

k = i + 10
If k > 7 Then k = 0


Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 10
For i = 5 To 7 Step 10
 df = df + 1
Picture1.Circle (X - 2500 * Sin(Tan(Cos(t(i) / k * 73.1415 / 2080))), Y - 2500 * Tan(Cos(Tan(t(i) * 5.1415 / 3340)))), 500, &H80FF&, , , 0.425
Picture1.Circle (X - 2500 * Sin(Tan(Cos(t(i) / k * 73.1415 / 2080))), Y + 2500 * Tan(Cos(Tan(t(i) * 5.1415 / 3340)))), 200, vbYellow, , , 5
Picture1.Circle (X - 2500 * Sin(Tan(Cos(t(i) / k * 73.1415 / 2080))), Y - 2500 * Tan(Cos(Tan(t(i) * 5.1415 / 3340)))), 200, vbRed, , , 0.425
Picture1.Circle (X - 2500 * Sin(Tan(Cos(t(i) / k * 73.1415 / 2080))), Y + 2500 * Tan(Cos(Tan(t(i) * 5.1415 / 3340)))), 500, vbGreen, , , 0.425

Picture1.Circle (X + 2500 * Sin(Tan(Cos(t(i) / k * 73.1415 / 2080))), Y - 2500 * Tan(Cos(Tan(t(i) * 5.1415 / 3340)))), 500, vbYellow, , , 0.425
Picture1.Circle (X + 2500 * Sin(Tan(Cos(t(i) / k * 73.1415 / 2080))), Y + 2500 * Tan(Cos(Tan(t(i) * 5.1415 / 3340)))), 200, vbYellow, , , 5
Picture1.Circle (X + 2500 * Sin(Tan(Cos(t(i) / k * 73.1415 / 2080))), Y - 2500 * Tan(Cos(Tan(t(i) * 5.1415 / 3340)))), 200, vbRed, , , 0.425
Picture1.Circle (X + 2500 * Sin(Tan(Cos(t(i) / k * 73.1415 / 2080))), Y + 2500 * Tan(Cos(Tan(t(i) * 5.1415 / 3340)))), 500, vbBlue, , , 0.425

If df >= 11000 Then df = 0: Option11.Value = True: Exit For
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next


End Sub
Sub fun11()
On Error Resume Next


 Picture1.Top = 0
 Picture1.Height = Me.Height + 200
 
X = 6400
Y = 4450
k = 0


For i = 0 To 7 Step 7
df = df + 1
k = i + 1
If k > 7 Then k = 6
Picture1.DrawMode = 13
Picture1.Circle (X - 4300 * Sin(Tan(t(k / 1.2) / 180)), Y + 4300 * Cos(t(k / 1.2) / 110)), 500, vbYellow
Picture1.Circle (X + 4300 * Sin(Tan(t(k / 1.2) / 180)), Y + 4300 * Cos(t(k / 1.2) / 110)), 500, vbYellow

Picture1.Circle (X - 4300 * Sin(Tan(t(k / 1.2) / 180)), Y - 4300 * Cos(t(k / 1.2) / 110)), 500, vbYellow
Picture1.Circle (X + 4300 * Sin(Tan(t(k / 1.2) / 180)), Y - 4300 * Cos(t(k / 1.2) / 110)), 500, vbYellow


If df >= 11000 Then df = 0: Option12.Value = True: Exit For
Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 0
For i = 0 To 7
Picture1.DrawMode = 13
Circle (X + 4300 * Sin(t(k / 1.2) / 180), Y + 4300 * Cos(t(k / 1.2) / 180)), 300, vbYellow
Circle (X - 4300 * Sin(t(k / 1.2) / 180), Y + 4300 * Cos(t(k / 1.2) / 180)), 300, vbYellow
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
'Circle (x - 4300 * Sin(t(k / 1.2) / 180), y + 4300 * Cos(t(k / 1.2) / 180)), 300, vbGreen

Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next
End Sub
Private Sub Timer3_Timer()
' fun12
Label1.Caption = df
End Sub
Sub fun13()
  On Error Resume Next

 Picture1.Top = 0
 Picture1.Height = Me.Height + 200
 
X = 6400
Y = 4450
k = 0
k = 0
For i = 0 To 7 Step 7
df = df + 1
k = i + 1
If k > 7 Then k = 2
Picture1.DrawMode = 13

Picture1.Circle (X - 4300 * Sin(Cos(Tan(Cos(t(k / 1.2) / 360)))), Y + 4300 * Cos(t(k / 1.2) / 410)), 500, &HFFFF&
Picture1.Circle (X + 4300 * Sin(Cos(Tan(Cos(t(k / 1.2) / 360)))), Y + 4300 * Cos(t(k / 1.2) / 410)), 500, &HFFFF&

Picture1.Circle (X - 4300 * Sin(Cos(Tan(Cos(t(k / 1.2) / 360)))), Y - 4300 * Cos(t(k / 1.2) / 410)), 500, &HFFFF&
Picture1.Circle (X + 4300 * Sin(Cos(Tan(Cos(t(k / 1.2) / 360)))), Y - 4300 * Cos(t(k / 1.2) / 410)), 500, &HFFFF&



'DrawMode = 2
Picture1.Circle (X + 2300 * Sin(t(k / 1.2) / 180), Y + 2300 * Cos(t(k / 1.2) / 180)), 200, vbYellow
Picture1.Circle (X - 2300 * Sin(t(k / 1.2) / 180), Y + 2300 * Cos(t(k / 1.2) / 180)), 200, vbYellow

Picture1.Circle (X + 2300 * Sin(t(k / 1.2) / 180), Y - 2300 * Cos(t(k / 1.2) / 180)), 200, vbYellow
Picture1.Circle (X - 2300 * Sin(t(k / 1.2) / 180), Y - 2300 * Cos(t(k / 1.2) / 180)), 200, vbYellow




If df >= 11000 Then df = 0: Option1.Value = True: Exit For

Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 0
For i = 0 To 7
Picture1.DrawMode = 13
'Circle (x + 3300 * Sin(t(k / 1.2) / 1180), y + 3300 * Cos(t(k / 1.2) / 180)), 300, vbRed
'Circle (x - 3300 * Sin(t(k / 1.2) / 1180), y + 3300 * Cos(t(k / 1.2) / 180)), 300, vbRed
k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
'Circle (x - 4300 * Sin(t(k / 1.2) / 180), y + 4300 * Cos(t(k / 1.2) / 180)), 300, vbGreen

Picture1.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next
End Sub
Sub fun12()
  On Error Resume Next

 Picture2.Top = 200
 Picture2.Height = Me.Height + 200
 
X = 6400
Y = 4450
k = 0


For i = 0 To 7 Step 10

k = i + 10
If k > 7 Then k = 0





Next
For i = 0 To 7 'Line (x, y + 2000)-(x + 500 * Sin(t(i) * 3.1415), y), vbYellow
'Line (x, y - 1500)-(x + 150 * Sin(t(i) * 3.1415 / 180), y + 10 * Cos(t(i) * 3.1415 / 10)), vbWhite, BF
t(i) = t(i)
Next
k = 10
For i = 5 To 7 Step 10

Picture2.Circle (X - 2500 * Cos(t(i) / k * 73.1415 / 2080), Y - 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, vbYellow, , , 5
Picture2.Circle (X - 2500 * Cos(t(i) / k * 73.1415 / 2080), Y + 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, vbYellow, , , 5
Picture2.Circle (X - 2500 * Cos(t(i) / k * 73.1415 / 2080), Y - 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, &H8000&, , , 0.425
Picture2.Circle (X - 2500 * Cos(t(i) / k * 73.1415 / 2080), Y + 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, &H8000&, , , 0.425

Picture2.Circle (X + 2500 * Cos(t(i) / k * 73.1415 / 2080), Y - 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, vbYellow, , , 5
Picture2.Circle (X + 2500 * Cos(t(i) / k * 73.1415 / 2080), Y + 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, vbYellow, , , 5
Picture2.Circle (X + 2500 * Cos(t(i) / k * 73.1415 / 2080), Y - 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, &H8000&, , , 0.425
Picture2.Circle (X + 2500 * Cos(t(i) / k * 73.1415 / 2080), Y + 2500 * Cos(Tan(t(i) * 5.1415 / 3340))), 500, &H8000&, , , 0.425

Picture2.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture2.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture2.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425
Picture2.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425

Picture2.Circle (X + 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture2.Circle (X + 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture2.Circle (X + 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425
Picture2.Circle (X + 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425

Picture2.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture2.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, vbYellow, , , 5
Picture2.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y - 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425
Picture2.Circle (X - 2500 * Tan(Sin(t(i) / k * 73.1415 / 2080)), Y + 2500 * Tan(t(i) * 5.1415 / 3340)), 500, &H8000&, , , 0.425


k = i + 1
If k > 7 Then k = 0
'Line (x + 1 + Sin(t(i) + 3.1415 / 10), y + 1500 * Cos(t(i) * 3.1415 / 1))-(x + 1500 * Sin(t(k) * 3.1415 / 386), y + 100 * Tan(t(k / 2))), &H8000&, BF
Next
For i = 0 To 7
Picture2.DrawWidth = 1
'Line (x, y)-(x + 0 * Sin(t(i) * 3.1415 / 180), 0 * Cos(t(i) * 0)), vbYellow
'Line (x, y)-(x + 1500 * Cos(t(i / 3) * 3.1415 / 180), 500 * Tan(t(i) * 0)), &H8000&
t(i) = t(i) + 2
Next


End Sub
