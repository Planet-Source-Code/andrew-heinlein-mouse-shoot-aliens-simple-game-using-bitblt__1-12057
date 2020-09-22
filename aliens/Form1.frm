VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "src pics"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture7 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   600
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   480
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   120
         Picture         =   "Form1.frx":0842
         ScaleHeight     =   450
         ScaleWidth      =   360
         TabIndex        =   14
         Top             =   1440
         Width           =   360
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":0F54
         ScaleHeight     =   480
         ScaleWidth      =   360
         TabIndex        =   13
         Top             =   840
         Width           =   360
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":1696
         ScaleHeight     =   480
         ScaleWidth      =   675
         TabIndex        =   4
         Top             =   240
         Width           =   675
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   75
         Left            =   960
         Picture         =   "Form1.frx":20D8
         ScaleHeight     =   75
         ScaleWidth      =   150
         TabIndex        =   3
         Top             =   480
         Width           =   150
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   600
         Picture         =   "Form1.frx":21BA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1440
      Top             =   5040
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4575
      Left            =   120
      MouseIcon       =   "Form1.frx":29FC
      MousePointer    =   99  'Custom
      ScaleHeight     =   4515
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   8880
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   8880
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   8880
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
'ships X/Y:
Private CurrY As Long, CurrX As Long
'Max of 21 bullets at one time:
Private BullX(20) As Long, BullY(20) As Long
'max of 21 stars at one time:
Private StarX(20) As Long, StarY(20) As Long, StarC(20) As Long
'max of 6 aliens at one time:
Private AlienX(5) As Long, AlienY(5) As Long, AlienT(5) As Integer
'score keepers:
Private AliensHit As Integer, BullsFired As Integer, IDied As Integer, BullsMissed As Integer, AliensMissed As Integer

Private Sub Command1_Click()
    AliensHit = 0
    BullsFired = 0
    IDied = 0
    BullsMissed = 0
    AliensMissed = 0
End Sub

Sub SetScore()
    Label1.Caption = "I Died: " & IDied
    Label2.Caption = "Aliens Hit: " & AliensHit
    Label3.Caption = "Aliens Missed: " & AliensMissed
    Label4.Caption = "Bullets Fired: " & BullsFired
    Label5.Caption = "Waisted Bullets: " & BullsMissed
    If BullsFired <> 0 Then
        Label6.Caption = "Firing Accuracy: " & Int((AliensHit / BullsFired) * 100) & "%"
    Else
        Label6.Caption = "Firing Accuracy: 100%"
    End If
    Label7.Caption = "Total: " & (AliensHit * 5) - (IDied * 6) - (AliensMissed * 2) & " points"
End Sub

Private Sub Form_Load()
    'set the initial values
    Randomize
    For i = 0 To 20
        StarX(i) = Rnd * (Picture1.Width / 15)
        StarY(i) = Rnd * (Picture1.Height / 15)
        StarC(i) = Rnd * &HFFFFFF
        If i <= 5 Then
            AlienX(i) = Rnd * (Picture1.Width / 15)
            AlienY(i) = Rnd * (Picture1.Height / 15)
            AlienT(i) = Int(Rnd * 4)
        End If
    Next i
End Sub

Private Sub Picture1_Click()
    'shoot a bullet if it isnt active
    For i = 0 To 20
        If BullX(i) = 0 And BullY(i) = 0 Then
            BullsFired = BullsFired + 1
            BullX(i) = CurrX + (Picture2.Width / 15)
            BullY(i) = CurrY + (Picture2.Height / 15) / 2
            Exit Sub
        End If
    Next i
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'when the mouse moves over, set the X/Y of your ship
    'when you divide twips by 15, that will convert it to pixels
    CurrX = x / 15
    CurrY = y / 15
End Sub

Private Sub Timer1_Timer()
    
    For i = 0 To 20
        If BullX(i) * 15 > Picture1.Width Then
            'a bullet hit the end of the screen, reset it
            BullsMissed = BullsMissed + 1
            BullX(i) = 0
            BullY(i) = 0
        End If
        If BullX(i) <> 0 And BullY(i) <> 0 Then
            'move an active bullet forward
            BullX(i) = BullX(i) + 10
        End If
        If StarX(i) <= 0 Then
            'if a star is all the way to the left, move it to the front
            Randomize
            StarC(i) = Rnd * &HFFFFFF
            StarX(i) = Picture1.Width / 15
            StarY(i) = Rnd * (Picture1.Height / 15)
        End If
        'move a star back
        StarX(i) = StarX(i) - 2
        If i <= 5 Then
            Randomize
            If AlienX(i) < 0 Then
                'an alien made it all the way to the left, move it to the front
                AliensMissed = AliensMissed + 1
                AlienT(i) = Int(Rnd * 4)
                AlienX(i) = Picture1.Width / 15
                AlienY(i) = Rnd * (Picture1.Height / 15)
            End If
            If AlienX(i) <> 0 And AlienY(i) <> 0 Then
                'move the alien back
                AlienX(i) = AlienX(i) - 5
            Else
                'reset the alien to the front if it was destroyed
                AlienT(i) = Int(Rnd * 4)
                AlienX(i) = Picture1.Width / 15
                AlienY(i) = Rnd * (Picture1.Height / 15)
            End If
        End If
    Next i
    
    'SetAll' will paint everything
    SetAll
    
    'check if bullets hit an alien
    For i = 0 To 20
        If BullX(i) <> 0 And BullY(i) <> 0 Then
            For n = 0 To 5
                If BullX(i) >= AlienX(n) And BullX(i) <= AlienX(n) + (Picture4.Width / 15) Then
                    If BullY(i) >= AlienY(n) And BullY(i) <= AlienY(n) + (Picture4.Height / 15) Then
                        'i hit an alien with a bullet
                        AliensHit = AliensHit + 1
                        BullX(i) = 0
                        BullY(i) = 0
                        AlienX(n) = 0
                        AlienY(n) = 0
                    End If
                End If
            Next n
        End If
    Next i
    
    'check if an alien hit me
    For i = 0 To 5
        If AlienX(i) >= CurrX And AlienX(i) <= CurrX + (Picture2.Width / 15) Then
            If AlienY(i) >= CurrY And AlienY(i) <= CurrY + (Picture2.Height / 15) Then
                'i was hit by an alien
                IDied = IDied + 1
                AlienX(i) = 0
                AlienY(i) = 0
            End If
        End If
    Next i
    
End Sub

Public Sub SetAll()
    Dim RND_NUM As Integer
    Picture1.Cls
    For i = 0 To 20
        'paint the active bullets
        If BullX(i) <> 0 And BullY(i) <> 0 Then
            BitBlt Picture1.hdc, BullX(i), BullY(i), Picture3.Width / 15, Picture3.Height / 15, Picture3.hdc, 0, 0, SRCINVERT
        End If
        'paint the stars,(setting a pixel) the random crap makes them "twinkle"
        Randomize
        RND_NUM = Rnd * 10
        If RND_NUM >= 5 Then
            SetPixel Picture1.hdc, StarX(i), StarY(i), StarC(i)
            SetPixel Picture1.hdc, StarX(i) + 1, StarY(i) + 1, StarC(i)
        Else
            SetPixel Picture1.hdc, StarX(i), StarY(i), StarC(i)
        End If
        'paint the aliens
        If i <= 5 Then
            If AlienX(i) <> 0 And AlienY(i) <> 0 Then
                If AlienT(i) = 0 Then
                    BitBlt Picture1.hdc, AlienX(i), AlienY(i), Picture4.Width / 15, Picture4.Height / 15, Picture4.hdc, 0, 0, SRCINVERT
                ElseIf AlienT(i) = 1 Then
                    BitBlt Picture1.hdc, AlienX(i), AlienY(i), Picture5.Width / 15, Picture5.Height / 15, Picture5.hdc, 0, 0, SRCINVERT
                ElseIf AlienT(i) = 2 Then
                    BitBlt Picture1.hdc, AlienX(i), AlienY(i), Picture6.Width / 15, Picture6.Height / 15, Picture6.hdc, 0, 0, SRCINVERT
                Else
                    BitBlt Picture1.hdc, AlienX(i), AlienY(i), Picture7.Width / 15, Picture7.Height / 15, Picture7.hdc, 0, 0, SRCINVERT
                End If
            End If
        End If
    Next i
    'paint the ship
    BitBlt Picture1.hdc, CurrX, CurrY, Picture2.Width / 15, Picture2.Height / 15, Picture2.hdc, 0, 0, SRCINVERT
    'set the score
    SetScore
End Sub
