VERSION 5.00
Begin VB.Form frmRank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Army Ranks"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   6360
      Top             =   2760
   End
   Begin VB.PictureBox e 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   27
      Left            =   6240
      Picture         =   "frmRank.frx":0000
      ScaleHeight     =   1305
      ScaleWidth      =   2265
      TabIndex        =   56
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   26
      Left            =   6240
      Picture         =   "frmRank.frx":0657
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   55
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   25
      Left            =   6480
      Picture         =   "frmRank.frx":0CF5
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   54
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   24
      Left            =   6240
      Picture         =   "frmRank.frx":13BC
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   53
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   23
      Left            =   6240
      Picture         =   "frmRank.frx":1A97
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   52
      Top             =   720
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   22
      Left            =   6480
      Picture         =   "frmRank.frx":2164
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   51
      Top             =   720
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   21
      Left            =   6360
      Picture         =   "frmRank.frx":2EC8
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   50
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   20
      Left            =   6360
      Picture         =   "frmRank.frx":3A49
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   49
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   19
      Left            =   6360
      Picture         =   "frmRank.frx":44A5
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   48
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   18
      Left            =   6360
      Picture         =   "frmRank.frx":4D1B
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   47
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   17
      Left            =   6360
      Picture         =   "frmRank.frx":53C7
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   46
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   16
      Left            =   6360
      Picture         =   "frmRank.frx":60AF
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   45
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   15
      Left            =   6360
      Picture         =   "frmRank.frx":6C0F
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   44
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   14
      Left            =   6360
      Picture         =   "frmRank.frx":7A69
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   43
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   13
      Left            =   6360
      Picture         =   "frmRank.frx":8430
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   42
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   12
      Left            =   6360
      Picture         =   "frmRank.frx":8B56
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   41
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Reload"
      Height          =   495
      Left            =   9120
      TabIndex        =   40
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   11
      Left            =   6360
      Picture         =   "frmRank.frx":92EB
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   39
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   10
      Left            =   6360
      Picture         =   "frmRank.frx":9C2E
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   38
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   9
      Left            =   6360
      Picture         =   "frmRank.frx":A62C
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   37
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   8
      Left            =   6360
      Picture         =   "frmRank.frx":B82B
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   36
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   7
      Left            =   6360
      Picture         =   "frmRank.frx":CA45
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   35
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   6
      Left            =   6360
      Picture         =   "frmRank.frx":DB18
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   34
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   5
      Left            =   6360
      Picture         =   "frmRank.frx":EB36
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   33
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   4
      Left            =   6360
      Picture         =   "frmRank.frx":FA1D
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   32
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   3
      Left            =   6360
      Picture         =   "frmRank.frx":10AF0
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   31
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   2
      Left            =   6360
      Picture         =   "frmRank.frx":11A44
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   30
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   1
      Left            =   6360
      Picture         =   "frmRank.frx":126FF
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   29
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox e 
      Height          =   1335
      Index           =   0
      Left            =   6360
      Picture         =   "frmRank.frx":13572
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   28
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton c 
      Caption         =   "Cheif Warrant Officer 5"
      Height          =   495
      Index           =   27
      Left            =   5640
      TabIndex        =   27
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton c 
      Caption         =   "Chief Warrant Officer 4"
      Height          =   495
      Index           =   26
      Left            =   5640
      TabIndex        =   26
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton c 
      Caption         =   "Chief Warrant Officer 3"
      Height          =   495
      Index           =   25
      Left            =   5640
      TabIndex        =   25
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton c 
      Caption         =   "Chief Warrant Officer 2"
      Height          =   495
      Index           =   24
      Left            =   5640
      TabIndex        =   24
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton c 
      Caption         =   "Warrant Officer 1"
      Height          =   495
      Index           =   23
      Left            =   5640
      TabIndex        =   23
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton c 
      Caption         =   "General of the Army"
      Height          =   495
      Index           =   22
      Left            =   3120
      TabIndex        =   22
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "General"
      Height          =   495
      Index           =   21
      Left            =   3120
      TabIndex        =   21
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Lieutenant General"
      Height          =   495
      Index           =   20
      Left            =   3120
      TabIndex        =   20
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Major General"
      Height          =   495
      Index           =   19
      Left            =   3120
      TabIndex        =   19
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Brigadier General"
      Height          =   495
      Index           =   18
      Left            =   3120
      TabIndex        =   18
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Colonel"
      Height          =   495
      Index           =   17
      Left            =   3120
      TabIndex        =   17
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Lieutenant Colonel"
      Height          =   495
      Index           =   16
      Left            =   3120
      TabIndex        =   16
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Major"
      Height          =   495
      Index           =   15
      Left            =   3120
      TabIndex        =   15
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Captain"
      Height          =   495
      Index           =   14
      Left            =   3120
      TabIndex        =   14
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "First Lieutenant"
      Height          =   495
      Index           =   13
      Left            =   3120
      TabIndex        =   13
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Second Lieutenant"
      Height          =   495
      Index           =   12
      Left            =   3120
      TabIndex        =   12
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Sergeant Major of the Army"
      Height          =   495
      Index           =   11
      Left            =   480
      TabIndex        =   11
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Command Sergeant Major"
      Height          =   495
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Sergeant Major"
      Height          =   495
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "First Sergeant"
      Height          =   495
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Master Sergeant"
      Height          =   495
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Sergeant First Class"
      Height          =   495
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Staff Sergeant"
      Height          =   495
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Sergeant"
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Corporal"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Specialist"
      Height          =   495
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Private First Class"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton c 
      Caption         =   "Private"
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "frmRank.frx":14173
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblwrong 
      Caption         =   "Wrong !"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6360
      TabIndex        =   58
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblcorrect 
      Caption         =   "Correct !"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   6360
      TabIndex        =   57
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########################################
'# This software may only be used for
'# educational purposes. If you want to use
'# it commercially please contact me for
'# permission to do so.
'# Otherwise you are free to modify it
'# to meet your needs.
'############################################
'# Please note: I am not responsibly for the
'# outcome of using this software and it doesn't
'# include any written or implied warranties of
'# anykind... in english ... use at your own risk.
'#
'# you may contact me at jason@vzio.com
'#
'###########################################



Private Sub c_Click(Index As Integer)
'# onclick run CheckAnswer Function
'# heck with the picture displayed to see if the index numbers match for a correct answer.

CheckAnswer (Index)
End Sub

Private Sub Command29_Click()
'# Reload / next button
'# displays new insignia rank
doRandomRank
End Sub

Private Sub e_Click(Index As Integer)
'# onclick of picture show answer in messagebox
'# basicly just grab the caption from the button which matches index numbers for this picture...
MsgBox c(Index).Caption, vbInformation, "Answer"
End Sub

Private Sub Form_Load()
'# Reload / next button
'# displays random insignia rank
doRandomRank
End Sub


Function doRandomRank()
'# this function loads a random Rank Insignia based on 27 index numbers (PictureBox)

Dim i As Integer
Dim intRankId As Integer

'# hide all images
For i = 0 To 27
       e(i).Visible = False
Next

'# grab random index number to display Insignia
Randomize
intRankId = Int(Rnd * 27)

'# display image
e(intRankId).Visible = True


End Function


Function GrabCurrentImage()
'# grab current displayed image so we can check answer later on click
    For i2 = 0 To 27
        If e(i2).Visible = True Then
           GrabCurrentImage = i2
        'Else
        End If
    Next
End Function



Function CheckAnswer(intIndex As Integer)

Timer1.Enabled = False  '# the timer hides the wrong/correct labels after 3-4 seconds

If CInt(intIndex) = CInt(GrabCurrentImage) Then  '# button clicked matches image
    doRandomRank                                '# load new Insignia
    lblcorrect.Visible = True                   '# show correct!
    lblwrong.Visible = False                    '# hide wrong label
    'c(intIndex).Enabled = False
   Timer1.Enabled = True                        'hide correct after 3 secs
Else

    lblcorrect.Visible = False                  '#answer incorrect show false...
    lblwrong.Visible = True
    Timer1.Enabled = True

End If


End Function



Private Sub Timer1_Timer()
If lblcorrect.Visible = True Then
    lblcorrect.Visible = False
End If
If lblwrong.Visible = True Then
lblwrong.Visible = False
End If
'Me.Enabled = False  freezes me
End Sub
