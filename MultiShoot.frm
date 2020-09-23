VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCS"
   ClientHeight    =   7860
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox grass 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6720
      Picture         =   "MultiShoot.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   8160
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   7110
      Left            =   8400
      Picture         =   "MultiShoot.frx":236B
      ScaleHeight     =   7050
      ScaleWidth      =   5745
      TabIndex        =   14
      Top             =   0
      Width           =   5805
   End
   Begin VB.PictureBox minepic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   345
      Left            =   5520
      Picture         =   "MultiShoot.frx":1BA5C
      ScaleHeight     =   285
      ScaleWidth      =   360
      TabIndex        =   12
      Top             =   8160
      Width           =   420
   End
   Begin VB.PictureBox bulletpic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   6240
      Picture         =   "MultiShoot.frx":1BFF6
      ScaleHeight     =   105
      ScaleWidth      =   90
      TabIndex        =   11
      Top             =   8040
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox hero 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2010
      Left            =   7680
      Picture         =   "MultiShoot.frx":1C0C6
      ScaleHeight     =   1950
      ScaleWidth      =   1080
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2010
      Left            =   8880
      Picture         =   "MultiShoot.frx":22EB8
      ScaleHeight     =   1950
      ScaleWidth      =   1080
      TabIndex        =   9
      Top             =   7920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox stone 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   10080
      Picture         =   "MultiShoot.frx":29CAC
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   8
      Top             =   7320
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   0
      Picture         =   "MultiShoot.frx":471AE
      ScaleHeight     =   7035
      ScaleWidth      =   5715
      TabIndex        =   7
      Top             =   0
      Width           =   5775
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   1455
         Index           =   8
         Left            =   4080
         Top             =   4200
         Width           =   495
      End
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   975
         Index           =   7
         Left            =   0
         Top             =   4800
         Width           =   2775
      End
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   1695
         Index           =   6
         Left            =   4200
         Top             =   1800
         Width           =   495
      End
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   135
         Index           =   5
         Left            =   3720
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   135
         Index           =   0
         Left            =   480
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   1455
         Index           =   1
         Left            =   3000
         Top             =   360
         Width           =   135
      End
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   135
         Index           =   2
         Left            =   3240
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   1455
         Index           =   3
         Left            =   1200
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Wall 
         BackColor       =   &H00000080&
         Height          =   1095
         Index           =   4
         Left            =   2640
         Top             =   2280
         Width           =   135
      End
   End
   Begin MSComctlLib.ProgressBar pLife 
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   7560
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ProgressBar pAmmo 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7560
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   30
   End
   Begin VB.Timer ReloadTimer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   7200
   End
   Begin MSWinsockLib.Winsock udpClient 
      Left            =   3000
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtChat 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7335
      Left            =   5880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Timer Movee 
      Interval        =   50
      Left            =   2400
      Top             =   7320
   End
   Begin VB.Label Label5 
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   7920
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Press Ctrl+R when you are READY"
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
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LIFE"
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
      Left            =   3360
      TabIndex        =   4
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AMMO"
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
      Left            =   0
      TabIndex        =   3
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Not connected"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7100
      Width           =   5775
   End
   Begin VB.Line rightb 
      X1              =   5760
      X2              =   5760
      Y1              =   7080
      Y2              =   0
   End
   Begin VB.Line downb 
      X1              =   0
      X2              =   5760
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Menu mNetwork 
      Caption         =   "Network"
      Begin VB.Menu mHost 
         Caption         =   "Host"
      End
      Begin VB.Menu mConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mDisconnect 
         Caption         =   "Disconnect"
      End
   End
   Begin VB.Menu mInteraction 
      Caption         =   "Interaction"
      Begin VB.Menu mMessage 
         Caption         =   "Message"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mReady 
         Caption         =   "READY!"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mOptions 
      Caption         =   "Options"
      Begin VB.Menu mClear 
         Caption         =   "Clear chat"
      End
      Begin VB.Menu mNewGame 
         Caption         =   "New Game"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mControls 
         Caption         =   "Controls"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple UDP Winsock Multiplayer Game by Mephisto
'Keys:
'Mouse1: Fire, ArrowKeys: Move, Shift: Mines, Mouse2: Position/Reload

'CREDITS FOR LINE COLLISION TO: ICEPLUG
'who wrote the VertCollision and Between Functions, which i then adjusted a little bit
'Thanks to Iceplug for all his help in my dark times :)
'My code for collision was faulty in vertical lines and his code fixed that

'CREATING YOUR OWN MAP:
'the code i made is really flexible, and you can have aas many walls as you want, until your
'computer can take it :) Simple add in more wall() controls or rearange them. Make sure the guy
'you play with has the same walls !
'Note: I wanted to make it so that you have an editor and map creator, but right now i decided that
'i will make a second version of this game, which will be a deathmatch for a lot more people, not
'just two players, but more people, how many wanna join that many will be able to play. And then
'a map creator is worth it
'Note2: No square is allowed to be around left up corner... Coordinate 0,0 and 10,10 has to
'be free, thats pretty low space but just making sure you know that
'Note3: The code at times is a little chaotic mainly because i never imagined i would take the game
'this far, i wanted to release it once when you were only some dots and were shooting another
'dot, but i took it extra step with bitblt, mines, added music, and more small things

'Enjoy

Dim CX As Double, CY As Double 'currentX and currentY
Dim NX As Double, NY As Double 'nextX , nextY

Dim TarX As Double, TarY As Double 'Target X and Target Y (the opponent)
Dim TarDir As Integer, TarStep As Integer
Dim LastTarX As Double, LastTarY As Double

Dim bUp As Boolean, bDown As Boolean, bLeft As Boolean, bRight As Boolean, bShift As Boolean 'for movement
Dim V As Integer, BV As Integer 'velocity and bullet velocity
Dim Multip As Double 'a temporary value used in Mouse_down

Private Type Lin 'Each line from rectangles
x1 As Double
y1 As Double
x2 As Double
y2 As Double
End Type

Private Type bullet 'Bullet Type
X As Double
Y As Double
Xvec As Single
Yvec As Single
enabled As Boolean
End Type

Private Type Mine 'mine Type
X As Double
Y As Double
enabled As Boolean
used As Boolean
End Type

Dim Lin() As Lin
Dim bullet(1 To 30) As bullet
Dim Mine(1 To 5) As Mine

Dim num As Integer 'a temp value

Dim i As Integer 'for loops
Dim collision As Boolean, see As Boolean 'boolean values for exactly what it says :)

Dim ammo As Integer, life As Single 'fundamental values of ammo and life
Dim canShoot As Boolean 'if user is reloading this is false

Dim BulletDam As Integer, MineDam As Integer 'the damage bullet and mine inflict

Dim state As String 'the state of the winsock is stored here
Dim div1 As String, div2 As String 'divisors used in transfer of data
Dim nick As String, nick2 As String 'nick of this player and the other player
Dim Points As Integer, Points2 As Integer 'points of each

Dim GameOn As Boolean, Ready As Boolean, Ready2 As Boolean 'basic booleans. If Ready and Ready2 then gameon = true

'FUNCTIONS FOR SOUND
Private Const SND_ASYNC = &H1         ' play asynchronously
Private Const SND_FILENAME = &H20000     ' name is a file name
Private Const SND_LOOP = &H8         ' loop the sound until next sndPlaySound
Private Const SND_PURGE = &H40 'stop sound

'some more API's
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

'has to do with bitblt stuff
Dim stepp As Integer
Dim posx(2) As Integer
Dim SX As Integer, SY As Integer

'once again bitblt makes it harder :)
Dim Xpos As Integer, Ypos As Integer
Dim closest As Integer, clo As Single, clopos As Integer
Dim tppx As Integer, tppy As Integer

'for RECT collision detection with walls
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long

'for debugging that i used
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim LastTime As Long

'for the other player bullets array and stuff
Private Type TarBullet
    X As Integer
    Y As Integer
End Type
Dim TarBullet() As TarBullet, TarB As Integer

Option Explicit

Private Sub Label6_Click()
MsgBox CX & ":" & CY & vbCrLf & NX & ":" & NY
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
bUp = True
Case vbKeyDown
bDown = True
Case vbKeyLeft
bLeft = True
Case vbKeyRight
bRight = True
End Select

End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Movee.enabled = False
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
bUp = False
Case vbKeyDown
bDown = False
Case vbKeyLeft
bLeft = False
Case vbKeyRight
bRight = False
Case vbKeyShift

If GameOn = True Then
    'new mine
    Dim mn As Integer 'mine number
    
    For i = 1 To 5
    If Mine(i).used = False Then
    mn = i
    End If
    Next i
    
    If mn <> 0 Then 'if we CAN place a mine
    Mine(mn).enabled = True
    Mine(mn).used = True
    Mine(mn).X = CX
    Mine(mn).Y = CY
    Else
    Add "Cannot lay more mines"
    End If
    
End If

End Select
End Sub

Private Sub Form_Load()
Randomize

CreateBackground 'initial bitblt stuff

tppx = Screen.TwipsPerPixelX
tppy = Screen.TwipsPerPixelY

'the dimensions of the guy on the sprite... each squre of his animation
'is 24 wide and 32 high
SX = 24
SY = 32

'posx positions of the guy on the sprite... dont try to understand this
'in here, because you willg et confused, look it up later on
posx(0) = 0
posx(1) = SX
posx(2) = SX * 2

bUp = False
bDown = False
bLeft = False
bRight = False

GameOn = False
canShoot = True

BulletDam = 10
MineDam = 50

ammo = 30
life = 100

'progress bars show the values
pAmmo.Value = ammo
pLife.Value = life

state = "not connected"

'div1 and div2 are just made up they dont have to be this
div1 = "==::=="
div2 = "]==["

CX = 0
CY = 0

'initial other user position. Pretty much useless because he isnt shown until gameon = true anyway
TarX = 2800
TarY = 3900

'Velocity of people and bullet
V = 100
BV = 300

num = 0

'store all sides of rectangles into the lin() array for line collision detection that
'is very important because it determines if one player can see the other
For i = 0 To Wall.UBound

'TOP SIDE
ReDim Preserve Lin(num)
Lin(num).x1 = Wall(i).Left
Lin(num).x2 = Wall(i).Left + Wall(i).Width
Lin(num).y1 = Wall(i).Top
Lin(num).y2 = Wall(i).Top
num = num + 1

'RIGHT SIDE
ReDim Preserve Lin(num)
Lin(num).x1 = Wall(i).Left + Wall(i).Width
Lin(num).x2 = Wall(i).Left + Wall(i).Width
Lin(num).y1 = Wall(i).Top
Lin(num).y2 = Wall(i).Top + Wall(i).Height

num = num + 1

'BOTTOM SIDE
ReDim Preserve Lin(num)
Lin(num).x1 = Wall(i).Left
Lin(num).x2 = Wall(i).Left + Wall(i).Width
Lin(num).y1 = Wall(i).Top + Wall(i).Height
Lin(num).y2 = Wall(i).Top + Wall(i).Height
num = num + 1

'LEFT SIDE
ReDim Preserve Lin(num)
Lin(num).x1 = Wall(i).Left
Lin(num).x2 = Wall(i).Left
Lin(num).y1 = Wall(i).Top + Wall(i).Height
Lin(num).y2 = Wall(i).Top
num = num + 1

Next i

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GameOn = True Then
'shoot
If Button = 1 Then
    If ammo > 0 Then
        If canShoot = True Then 'not reloading
        ammo = ammo - 1
        PlaySound App.Path & "\audio\fire.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
        
        pAmmo.Value = ammo
        
        'CODE FOR SHOOTING
        
        'we search an available bullet that wasnt shot yet
        For i = 1 To 30
        If bullet(i).enabled = False Then GoTo found
        Next i
        
found:
        bullet(i).enabled = True
        bullet(i).X = CX
        bullet(i).Y = CY
        
        'took me a long time to figure out. Baseically we have to consider where the
        'mouse is, and where the player is and then give the bullet vectorX it should
        'move in and VectorY it should move in.
        Multip = Sqr(((X - CX) ^ 2 + (Y - CY) ^ 2)) / BV
        bullet(i).Xvec = (X - CX) / Multip
        bullet(i).Yvec = (Y - CY) / Multip
        
        End If
    End If
    
ElseIf Button = 2 Then
'reload
PlaySound App.Path & "\audio\reload.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
ReloadTimer.enabled = True '2 seconds
canShoot = False
End If

Else
    'if the game is NOT on, then the user just wants to reposition himself. Do that
    
    NX = X
    NY = Y
    CX = X
    CY = Y
    
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'stop the sound
PlaySound vbNullString, 0&, SND_PURGE
End Sub

Private Sub Form_Unload(Cancel As Integer)
'tell the other player bye bye
If state = "connected" Then
udpClient.SendData "quit"
End If

DoEvents
udpClient.Close
End Sub

Private Sub mAbout_Click()
MsgBox "Made by: Mephisto" & vbCrLf & "Purpose: To show others how to do games like this, through internet. Enjoy"
End Sub

Private Sub mClear_Click()
txtChat.Text = ""
End Sub

Private Sub mConnect_Click()
Dim hostip As String
state = "Attempting to join game..."

hostip = InputBox("Please enter the host IP")

With udpClient
.Bind 5431
.RemoteHost = hostip
.RemotePort = 5432
End With

nick = InputBox("Choose a nick please")
If nick = "" Then
nick = "user" & Int((Rnd * 10) + 1)
MsgBox "error: null input. Automatically generated nick: " & nick
End If

'we tell the other person Hello :) And send over our nick
udpClient.SendData "hello" & div1 & nick & div2
mHost.enabled = False
mConnect.enabled = False

End Sub

Private Sub mControls_Click()
MsgBox "Move: Arrow Keys" & vbCrLf & "Shooting: Mouse1" & vbCrLf & "Sprint: Shift :Warning, sprint very slowly decreases your life!" & _
vbCrLf & "Oh and tip: When you are dead you can use your mouse to set the position you desire, dont have to walk there."
End Sub

Private Sub mDisconnect_Click()
state = "disconnected"
udpClient.Close

mHost.enabled = True
mConnect.enabled = True
End Sub

Private Sub mHost_Click()
state = "Hosting... waiting for connection"

nick = InputBox("Choose a nick please")
If nick = "" Then
nick = "user" & Int((Rnd * 10) + 1)
MsgBox "error: null input. Automatically generated nick: " & nick
End If

With udpClient
'host always has port 5432 listening, and client is 5431
.Bind 5432
End With

mHost.enabled = False
mConnect.enabled = False
End Sub

Private Sub mMessage_Click()
'send over a chat message
If state = "connected" Then

    Dim Message As String
    Message = InputBox("Say to other player: ")
    If Message <> "" Then
    udpClient.SendData "chat" & div1 & nick & div1 & Message & div2
    Add nick & " : " & Message
    End If

Else
Add "Error, Connect first!"
End If
End Sub

Private Sub mNewGame_Click()
'the other user must agree to this so we have to ask him if he does
Dim a As Integer

a = MsgBox("Are you sure you want to ask the other player for a new game ?", vbYesNo)

If a = vbYes Then
udpClient.SendData "question" & div1 & "newgame" & div2
End If

End Sub

Private Sub Movee_Timer()

LastTime = GetTickCount
'move da user :D
'We dont update the CX and CY immidiatelly. We store the new wannabe CX and CY into NX and NY
'which stands for NextX and NextY. Then we check if these NX and NY arent in a wall
'or outside a map. only then do we update CX and CY

Picture1.Picture = Picture2.Picture
Picture1.Picture = Picture1.Image

NY = CY
NX = CX

If bUp Then
NY = CY - V
End If

If bDown Then
NY = CY + V
End If

If bLeft Then
NX = CX - V
End If

If bRight Then
NX = CX + V
End If

    'check for walls
    collision = False
    
    'loop or rectangles and see if we are in any
    Dim r1 As RECT
    Dim r2 As RECT
    Dim r3 As RECT
    
    With r1
        .Top = NY - SY / 2
        .Bottom = NY + SY * tppy / 2
        .Left = NX - SX / 2
        .Right = NX + SX * tppx / 2
    End With
    
    For i = 0 To Wall.UBound
    With r2
        .Top = Wall(i).Top
        .Bottom = Wall(i).Top + Wall(i).Height
        .Left = Wall(i).Left
        .Right = Wall(i).Left + Wall(i).Width
    End With
        
    IntersectRect r3, r1, r2
    
    If IsRectEmpty(r3) = False Then
        collision = True
        Exit For
    End If

    Next i
    
    'if we are outside the map which is determined by 0's on left and top
    'and one line on down and one line on right so the arenas is customizable
    If NY < 0 Or NX < 0 Or NY > downb.y1 Or NX > rightb.x1 Then
    collision = True
    End If

Label5.Caption = "Collision check: " & GetTickCount - LastTime
LastTime = GetTickCount

    'if nothing happened then update CX and CY
If collision = False Then
    CX = NX
    CY = NY
End If

'here we determine which way to make the character face. THis is done by finding out
'the position of mouse relevant to 4 points. One point is 50 units up from character, one
'is right, other down and last one left. We use pytaghorian theorem to determine which
'is closest to mouse and we make him face the cursor
    closest = 10000
    
    clo = Sqr((CX / tppx - 50 - Xpos) ^ 2 + (CY / tppy - Ypos) ^ 2)
    If clo < closest Then
    closest = clo 'left
    clopos = 4
    End If
    
    clo = Sqr((CX / tppx - Xpos) ^ 2 + (CY / tppy - 50 - Ypos) ^ 2)
    If clo < closest Then
    closest = clo 'up is closest
    clopos = 1
    End If
    
    clo = Sqr((CX / tppx + 50 - Xpos) ^ 2 + (CY / tppy - Ypos) ^ 2)
    If clo < closest Then
    closest = clo 'right
    clopos = 2
    End If
    
    clo = Sqr((CX / tppx - Xpos) ^ 2 + (CY / tppy + 50 - Ypos) ^ 2)
    If clo < closest Then
    closest = clo 'down
    clopos = 3
    End If
    
    'the bitblting of the hero
    BitBlt Picture1.hDC, CX / tppx - SX / 2, CY / tppy - SY / 2, SX, SY, mask.hDC, posx(stepp), SY * (clopos - 1), vbMergePaint
    BitBlt Picture1.hDC, CX / tppx - SX / 2, CY / tppy - SY / 2, SX, SY, hero.hDC, posx(stepp), SY * (clopos - 1), vbSrcAnd
   
   'step is the x position of the sprite, so it looks like the guy is animated, like he is running
    If bUp Or bRight Or bDown Or bLeft Then
    stepp = stepp + 1
    End If
    
    If stepp = 3 Then stepp = 0
    
    Picture1.Picture = Picture1.Image

    Label5.Caption = Label5.Caption & vbCrLf & "Animate: " & GetTickCount - LastTime
    LastTime = GetTickCount

    'if the user tries to teleport himself into the wall, or walks into a wall while the game
    'is not ON, just teleport him back to origin, made it 10 10... not 0 0
    If collision = True And GameOn = False Then
    CX = 10
    CY = 10
    NX = 10
    NY = 10
    End If
    
    'we let other guy know our position
    If state = "connected" And GameOn = True Then
    udpClient.SendData "move" & div1 & (CX - SX / 2) & div1 & (CY - SY / 2) & div1 & clopos & div2
    End If

If GameOn = True Then
'check if target is visible

'we assume that we can see him and check if we cant
see = True

For i = 0 To UBound(Lin)
'VertCollide sub determines the line detection collision and if line between CX and CY - TarX and TarY
'crosses ANY other lin() then its no go. HE cant see him
If VertCollide(CX, CY, TarX, TarY, Lin(i).x1, Lin(i).y1, Lin(i).x2, Lin(i).y2) = True Then
see = False
GoTo done 'one of walls prevent seeing its over
End If

Next i

done:

'draw other players position
If see = True Then

    BitBlt Picture1.hDC, TarX / tppx - SX / 2, TarY / tppy - SY / 2, SX, SY, mask.hDC, posx(TarStep), SY * (TarDir - 1), vbMergePaint
    BitBlt Picture1.hDC, TarX / tppx - SX / 2, TarY / tppy - SY / 2, SX, SY, hero.hDC, posx(TarStep), SY * (TarDir - 1), vbSrcAnd
    
    If TarX <> LastTarX Or TarY <> LastTarY Then
        TarStep = TarStep + 1
        If TarStep = 3 Then TarStep = 0
    End If
    
    LastTarX = TarX
    LastTarY = TarY
End If

Label5.Caption = Label5.Caption & vbCrLf & "See/No see: " & GetTickCount - LastTime
LastTime = GetTickCount

'check for mines collision and draw mines that still exist
For i = 1 To 5
If Mine(i).enabled = True Then
    
    If Sqr((Abs(Mine(i).X - TarX)) ^ 2 + (Abs(Mine(i).Y - TarY)) ^ 2) < 100 Then
    'collision! enemy has touched the mine. haha
    udpClient.SendData "hit" & div1 & MineDam & div2
    Mine(i).enabled = False
    Add "Opponent hit your mine!"
    Else
    'just draw it
    BitBlt Picture1.hDC, Mine(i).X / tppx, Mine(i).Y / tppy, minepic.Width, minepic.Height, minepic.hDC, 0, 0, vbSrcCopy
    End If

End If
Next i

Label5.Caption = Label5.Caption & vbCrLf & "Mines: " & GetTickCount - LastTime
LastTime = GetTickCount

'move bullets and collision detection

For i = 1 To 30
If bullet(i).enabled = True Then
    'move each enabled bullet
    bullet(i).X = bullet(i).X + bullet(i).Xvec
    bullet(i).Y = bullet(i).Y + bullet(i).Yvec
    'Picture1.PSet (bullet(i).X, bullet(i).Y), vbRed
    BitBlt Picture1.hDC, bullet(i).X / tppx, bullet(i).Y / tppy, bulletpic.Width, bulletpic.Height, bulletpic.hDC, 0, 0, vbSrcAnd
    
    'if hits an edge
    If bullet(i).X < 0 Or bullet(i).Y < 0 Or bullet(i).X > rightb.x1 Or bullet(i).Y > downb.y1 Then
        bullet(i).enabled = False
    Else
        'we have to report position of EACH bullet to the other player
        If state = "connected" Then
        udpClient.SendData "bullet" & div1 & Int(bullet(i).X) & div1 & Int(bullet(i).Y) & div2
        End If
        
        'If a bullet hits the other guy then this computer reports it to him and he adjusts
        'his life value
        If Sqr((Abs(bullet(i).X - TarX)) ^ 2 + (Abs(bullet(i).Y - TarY)) ^ 2) < 100 Then
        'collision
        udpClient.SendData "hit" & div1 & BulletDam & div2
        End If
    
    End If
End If
Next i

For i = 0 To TarB - 1
'draw other players bullets
BitBlt Picture1.hDC, TarBullet(i).X, TarBullet(i).Y, bulletpic.Width, bulletpic.Height, bulletpic.hDC, 0, 0, vbSrcAnd
Next i

End If

If Label1.Caption <> state Then Label1.Caption = state 'on change of state tell that to user

Label5.Caption = Label5.Caption & vbCrLf & "Bullets: " & GetTickCount - LastTime
LastTime = GetTickCount

'erase the array, so it starts from again the next time timer is run and we have refreshed the
'bullets positions
ReDim TarBullet(0)
TarB = 0
End Sub

'A handy function for determining if a value lies between two other values.
Public Function Between(ByVal CValue As Double, ByVal Bound1 As Double, ByVal Bound2 As Double) As Boolean
'This function is a part of VertCollide
CValue = CLng(CValue)
Bound1 = CLng(Bound1)
Bound2 = CLng(Bound2)

    If (CValue >= Bound1 And CValue <= Bound2) Or (CValue <= Bound1 And CValue >= Bound2) Then Between = True
End Function

Private Function VertCollide(x1 As Double, y1 As Double, x2 As Double, y2 As Double, xx1 As Double, yy1 As Double, xx2 As Double, yy2 As Double)
'Iceplug's VertCollide funtion
'This is math stuff of lines, really confusing :)
Dim M1 As Double, M2 As Double, B1 As Double, B2 As Double, Collided As Boolean
Dim X As Double, Y As Double

If x1 = x2 Then
    M1 = 1000000
    B1 = 0
Else
    M1 = (y2 - y1) / (x2 - x1)
    B1 = y1 - M1 * x1
End If

If xx1 = xx2 Then
    M2 = 1000000
    B2 = 0
Else
    M2 = (yy2 - yy1) / (xx2 - xx1)
    B2 = yy1 - M2 * xx1
End If

If M1 = M2 Then
  If M2 = 1000000 Then
    If xx1 = x1 Then  'Vertical line collision.
      Collided = True
    End If
  ElseIf B1 = B2 Then  'This determines if the lines lie on top of each other.
    Collided = True
  End If
Else
    If M1 = 1000000 Then
      X = x1  'Vertical line.
      Y = M2 * X + B2  'use the decent line.
    ElseIf M2 = 1000000 Then
      X = xx1
      Y = M1 * X + B1  'use the decent line.
    Else
        X = (B1 - B2) / (M2 - M1)
        Y = M1 * X + B1
    End If
    
    If Between(X, xx1, xx2) And Between(X, x1, x2) And _
    Between(Y, yy1, yy2) And Between(Y, y1, y2) Then
      Collided = True
    End If
End If

If Collided Then
    VertCollide = True
Else
    VertCollide = False
End If

Close #1
End Function

Private Sub mReady_Click()
'If user is ready, report so
If state = "connected" Then
udpClient.SendData "ready" & div1 & "ready" & div2
Ready = True

'If the other user is ready already, then just start the game!
If Ready2 = True Then
GameOn = True
Add "FIGHT!!!!"
End If

Label4.Visible = False 'hes ready so disable the label
mReady.enabled = False
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xpos = X
Ypos = Y
    
    Xpos = Xpos / tppx
    Ypos = Ypos / tppy
    
End Sub

Private Sub ReloadTimer_Timer()
'reloading
canShoot = True
ammo = 30
pAmmo.Value = ammo
ReloadTimer.enabled = False

For i = 1 To 30
bullet(i).enabled = False
Next i

End Sub

Private Sub udpClient_DataArrival(ByVal bytesTotal As Long)
'CORE OF THE GAME - The communication centre
Dim sData As String
Dim Data() As String
udpClient.GetData sData

'Each command was structured A & div1 & B & div1 & C & div2 for example
'The Div2 has to be there because sometimes the packets join and this is formed :
'A & div1 & B & div1 & C & div2A & div1 & B & div1 & C & div2

'So we first split the Data by Div2 and then work only with the first part
'Then we split the remaineder by Div1 and we look what the other player is sending
Data = Split(sData, div2)
Data = Split(Data(0), div1)

Select Case Data(0)
Case "hello"
    'other user introduces himself
    state = "connected"
    nick2 = Data(1)
    MsgBox nick2 & " has joined the game!"
    udpClient.SendData "hi" & div1 & nick & div2
    Call New_Game
Case "hi"
    'host responded
    state = "connected"
    nick2 = Data(1)
    Call New_Game
Case "chat"
    'chat
    nick2 = Data(1)
    Add nick2 & " : " & Data(2)
Case "move"
    'new position of the other player
    TarX = Data(1)
    TarY = Data(2)
    TarDir = Data(3)
Case "quit"
    'other player left
    state = "not connected"
    MsgBox "The other user has left the game!"
Case "bullet"
    'There is a bullet that needs to be drawn
    'PlaySound App.Path & "\audio\fire.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
    
    ReDim Preserve TarBullet(TarB)
    TarBullet(TarB).X = Data(1) / tppx
    TarBullet(TarB).Y = Data(2) / tppy
    
    TarB = TarB + 1
Case "hit"
    'ohter player hit us
    life = life - Data(1) 'data(1) is the damage done
    If life < 0 Then life = 0 'make sure we dont cause life to be negative
    pLife.Value = life
    
    If life < 1 Then
    'over, this guy dead.
    Points2 = Points2 + 1
    udpClient.SendData "dead" & div1 & "dead" & div2 'let other player know we are dead
    Add nick & " : " & Points & vbCrLf & nick2 & " : " & Points2
    Call New_Round
    End If
    
Case "ready"
    'The other player is ready! Are we ? IF we are lets go!
    Ready2 = True
    Add nick2 & " is ready!"
    If Ready = True Then
    GameOn = True
    Add "FIGHT!!!!"
    End If
    
Case "dead"
    'other player dead
    Points = Points + 1
    Add nick & " : " & Points & vbCrLf & nick2 & " : " & Points2
    Call New_Round
Case "question"
    'Other user wants new game. This is not used much but i just wanted to demonstrate
    'how this is to be handled
        Dim a As Integer
        a = MsgBox("The other player wants to play a new game (reset points) do you agree ?")
        
        If a = vbYes Then
        udpClient.SendData "yes" & div1 & "yes" & div2
        Call New_Game
        Call New_Round
        Else
        udpClient.SendData "no" & div1 & "no" & div2
        End If
        
Case "yes"
    MsgBox nick2 & " has accepted your proposal"
    Call New_Game
    Call New_Round
Case "no"
    MsgBox nick2 & " has declined your proposal"
End Select

End Sub

Private Sub udpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical
End Sub

Sub New_Game()
Points = 0
Points2 = 0
End Sub

Sub New_Round()
'reininitalize crucial parts
Cls

life = 100
ammo = 30

pAmmo.Value = ammo
pLife.Value = life

Ready = False
Label4.Visible = True
GameOn = False

're-enable mines
For i = 1 To 5
Mine(i).enabled = False
Mine(i).used = False
Next i

mReady.enabled = True
End Sub

Sub Add(Message As String)
'add to txt
txtChat.Text = txtChat.Text & Message & vbCrLf
txtChat.SelStart = Len(txtChat.Text)
End Sub

Sub CreateBackground()
'i hate bitblt :)

'Now, i am not that good in bitblt and all this is determined also by an expermintal process
'i had many problems with bitblt and this worked, so it may be done even faster, i dont know,
'but this finally worked so i just left it like that. :) Please dont call me noob and close this
'code, but concentrate on winsock and how to use it to communicate to others =p

    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim intRows As Integer
    Dim intColumns As Integer
    Dim r As Integer
    Dim c As Integer
    Dim xc As Integer
    Dim yc As Integer
    
    Picture1.AutoRedraw = True
    
    intWidth = grass.ScaleWidth
    intHeight = grass.ScaleHeight
    intColumns = Int(Picture1.ScaleWidth / intWidth + 1)
    intRows = Int(Picture1.ScaleHeight / intHeight + 1)
   'copy the background tile... this is where we spray the grass onto the picture2
    yc = 0
    For r = 1 To intRows
        xc = 0
        For c = 1 To intColumns
            Picture1.PaintPicture grass.Picture, xc, yc
            xc = xc + intWidth
        Next c
        yc = yc + intHeight
    Next r
    
    'do stones
    For i = 0 To Wall.UBound
    'this code puts stones where we have rectangles
    Picture1.PaintPicture stone.Picture, Wall(i).Left, Wall(i).Top, , , 0, 0, Wall(i).Width, Wall(i).Height
    Wall(i).Visible = False
    Next i
    
    'ensure graphic persistance
    Picture1.Picture = Picture1.Image
    
    'make sure they same sizes
    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
    
    'and transfer the whole package to picture1
    Picture2.PaintPicture Picture1.Picture, 0, 0
    Picture2.Picture = Picture1.Picture
    
End Sub
