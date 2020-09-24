VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Noisy Class Test"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fr 
      Caption         =   " Speaker Let "
      Height          =   2160
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   225
      Width           =   1710
      Begin VB.CommandButton btBeep 
         Caption         =   "Beep"
         Height          =   375
         Left            =   780
         TabIndex        =   3
         Top             =   555
         Width           =   780
      End
      Begin VB.CheckBox ckMuteLet 
         Caption         =   "Mute"
         Height          =   300
         Left            =   645
         TabIndex        =   2
         Top             =   1620
         Width           =   825
      End
      Begin VB.VScrollBar scrVolLet 
         Height          =   1530
         LargeChange     =   20
         Left            =   270
         Max             =   0
         Min             =   100
         SmallChange     =   5
         TabIndex        =   1
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame fr 
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   2160
      Index           =   1
      Left            =   2235
      TabIndex        =   4
      Top             =   225
      Width           =   1710
      Begin VB.VScrollBar scrVolGet 
         Height          =   1530
         LargeChange     =   20
         Left            =   270
         Max             =   0
         Min             =   100
         SmallChange     =   100
         TabIndex        =   5
         Top             =   360
         Width           =   240
      End
      Begin VB.CheckBox ckMuteGet 
         Caption         =   "Mute"
         Height          =   300
         Left            =   660
         TabIndex        =   6
         Top             =   1620
         Width           =   825
      End
      Begin VB.Label lb 
         Caption         =   " Speaker Get "
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   0
         Width           =   990
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rudimentary Mixer Wrapper Class Test Driver
''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private Mixer           As New cMixer
Private BeepOn          As Boolean
Private WasScrolling    As Boolean

Private Sub btBeep_Click()

    Beep 'for test
    scrVolLet.SetFocus

End Sub

Private Sub ckMuteLet_Click()

    LetIt 'set / reset Speakers Mute
    GetIt 'verify
    scrVolLet.SetFocus

End Sub

Private Sub Form_Load()

    With Mixer
        'get current settings into let side (we have to start somewhere..)
        If .Choose(SpeakersOut, Volume) Then
            scrVolLet = .Value
        End If
        If .Choose(SpeakersOut, Mute) Then
            ckMuteLet = IIf(.Value = 0, vbUnchecked, vbChecked)
        End If
    End With 'MIXER
    '...and into get side
    GetIt

End Sub

Private Sub GetIt()

    With Mixer 'get current settings
        If .Choose(SpeakersOut, Volume) Then
            scrVolGet = .Value
        End If
        If .Choose(SpeakersOut, Mute) Then
            ckMuteGet = IIf(.Value = 0, vbUnchecked, vbChecked) '100 = true; 0 = false
        End If
    End With 'MIXER

End Sub

Private Sub LetIt()

    With Mixer 'set current values
        If .Choose(SpeakersOut, Volume) Then
            .Value = scrVolLet
        End If
        If .Choose(SpeakersOut, Mute) Then
            .Value = IIf(ckMuteLet = vbChecked, 100, 0) '100 = true; 0 = false
        End If
    End With 'MIXER
    If BeepOn Then
        Beep
    End If
    BeepOn = True

End Sub

Private Sub scrVolLet_Change()

    BeepOn = WasScrolling
    WasScrolling = False
    LetIt 'set Speakers Volume
    GetIt 'verify
    BeepOn = True

End Sub

Private Sub scrVolLet_Scroll()

    BeepOn = False
    WasScrolling = True
    LetIt 'set Speakers Volume
    GetIt 'verify

End Sub

':) Ulli's VB Code Formatter V2.17.4 (2004-Aug-02 14:36) 6 + 83 = 89 Lines
