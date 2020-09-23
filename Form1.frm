VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8265
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   4000
      Left            =   360
      Top             =   6000
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   960
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   360
      Top             =   5400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   6735
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         TabIndex        =   17
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Image Image12 
         Height          =   510
         Left            =   4680
         Picture         =   "Form1.frx":1A389
         Top             =   1200
         Width           =   525
      End
      Begin VB.Image Image11 
         Height          =   495
         Left            =   4320
         Picture         =   "Form1.frx":1A967
         Top             =   2520
         Width           =   420
         Visible         =   0   'False
      End
      Begin VB.Image Image10 
         Height          =   4095
         Left            =   0
         Picture         =   "Form1.frx":1AEB6
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   4095
         Left            =   6120
         Picture         =   "Form1.frx":1BDAF
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   525
         Left            =   0
         Picture         =   "Form1.frx":1CCA8
         Top             =   0
         Width           =   6660
      End
      Begin VB.Image Image6 
         Height          =   525
         Left            =   0
         Picture         =   "Form1.frx":1E31E
         Top             =   4200
         Width           =   6660
      End
      Begin VB.Image Image4 
         Height          =   390
         Left            =   600
         Picture         =   "Form1.frx":1F994
         Top             =   360
         Width           =   420
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   1440
         Picture         =   "Form1.frx":1FEDC
         Top             =   1560
         Width           =   450
      End
      Begin VB.Image Image7 
         Height          =   510
         Left            =   2760
         Picture         =   "Form1.frx":2030A
         Top             =   1200
         Width           =   525
      End
      Begin VB.Image Image1 
         Height          =   345
         Left            =   2280
         Picture         =   "Form1.frx":208E8
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8880
      TabIndex        =   16
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   8880
      TabIndex        =   15
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Det är du!"
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
      Left            =   5520
      TabIndex        =   13
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Du får normal hastighet men mister 4 hjärtan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Du får ett hjärta mer "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Antal liv kvar"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Du rör dig lånsammare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Mister ett liv"
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
      Left            =   3120
      TabIndex        =   8
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Image Image18 
      Height          =   345
      Left            =   4920
      Picture         =   "Form1.frx":20E37
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   4920
      Picture         =   "Form1.frx":21386
      Top             =   6840
      Width           =   420
   End
   Begin VB.Image Image16 
      Height          =   480
      Left            =   2520
      Picture         =   "Form1.frx":218D5
      Top             =   7440
      Width           =   510
   End
   Begin VB.Image Image15 
      Height          =   510
      Left            =   2520
      Picture         =   "Form1.frx":22001
      Top             =   6840
      Width           =   525
   End
   Begin VB.Image Image14 
      Height          =   390
      Left            =   2520
      Picture         =   "Form1.frx":225DF
      Top             =   6240
      Width           =   420
   End
   Begin VB.Image Image13 
      Height          =   375
      Left            =   4920
      Picture         =   "Form1.frx":22B27
      Top             =   6240
      Width           =   450
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lycka till!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Flytta Gubbo med piltangenterna. Försök att ta hjärtat utan att Storgubben får tag i dig!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 - stoppa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "F1 - starta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Left            =   8880
      TabIndex        =   2
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   8880
      TabIndex        =   1
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   8760
      Picture         =   "Form1.frx":22F55
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   8760
      Picture         =   "Form1.frx":233E6
      Top             =   1920
      Width           =   510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HS As Single
Dim hjärtan As Integer
Dim liv As Integer       'De variabler som ska
Dim fx As Integer    'användas på fler ställen
Dim fy As Integer
Dim gflytt As Integer
Dim slut As Single
Dim birj As Single
Dim stid As Integer
Dim mell As Single
Dim h5 As Boolean
Dim score As Single
Dim gActive As Boolean


'kollar om "uttrycket" är sant och returnerar true eller false
'värdena har de fått genom obk funktionen
Private Function r1(x1, x2, y1, y2, x3, x4, y3, y4) As Boolean
r1 = (x2 >= x3 And x4 >= x1) And (y2 >= y3 And y4 >= y1)
End Function



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim maxX As Single
Dim maxY As Single

Dim dx As Single
Dim dy As Single
dx = Image1.Left
dy = Image1.Top
Dim RamOffset As Single
RamOffset = 100

If KeyCode = vbKeyF1 Then
fy = 60
fx = 60
Image1.Visible = True
Image2.Visible = True
Image4.Visible = True
Timer1.Enabled = True
gActive = True
Label7.Caption = ""
Image6.Visible = False
Image8.Visible = False
Image9.Visible = False
Image10.Visible = False
Image7.Visible = True
Image12.Visible = True
Image11.Visible = True
gflytt = 200
birj = Timer
Label15.Visible = False
End If

If KeyCode = vbKeyF2 Then
Call slutt
End If

If gActive Then
Select Case KeyCode                 'Dy, Dx får olika
Case vbKeyDown: dy = dy + gflytt     'värden beroende på vilken
Case vbKeyUp: dy = dy - gflytt       'knapp man tryckt på
Case vbKeyLeft: dx = dx - gflytt
Case vbKeyRight: dx = dx + gflytt
End Select


maxX = Frame1.Width - Image1.Width - 100
maxY = Frame1.Height - Image1.Height - 100

If dx < 0 Then dx = 0           'Ej kunna gå utanför frame1
If dx > maxX Then dx = maxX     'med gubben
If dy < 0 Then dy = 0
If dy > maxY Then dy = maxY

Image1.Move dx, dy




If obk(Image1, Image7) Then
gflytt = 50
Image7.Move _
Rnd * (Frame1.Width - Image7.Width - RamOffset * 2), _
Rnd * (Frame1.Height - Image7.Height - RamOffset * 2)
End If


If obk(Image1, Image12) Then
gflytt = 50
Image12.Move _
Rnd * (Frame1.Width - Image12.Width - RamOffset * 2), _
Rnd * (Frame1.Height - Image12.Height - RamOffset * 2)
End If


If obk(Image1, Image11) Then
gflytt = 200
hjärtan = hjärtan - 4
Label1.Caption = hjärtan
Image11.Move _
Rnd * (Frame1.Width - Image11.Width - RamOffset * 2), _
Rnd * (Frame1.Height - Image11.Height - RamOffset * 2)
End If




If obk(Image1, Image2) Then 'sätter in  image1 och 2 i obk
Image2.Move _
Rnd * (Frame1.Width - Image2.Width - RamOffset * 2), _
Rnd * (Frame1.Height - Image2.Height - RamOffset * 2)
hjärtan = hjärtan + 1
Label1.Caption = hjärtan
End If

If obk(Image1, Image4) Then        'Kollar om image1 + 4 överlappar
liv = liv - 1                      'när man flyttar gubben, image1
Label2.Caption = liv
End If
End If

If liv < 1 Then
    Call dead
End If

If hjärtan = 2 Then
    slut = Timer
    Call slutt
    mell = slut - birj
    stid = mell * 10
    mell = stid / 10
    Label14.Visible = False
    Label15.Visible = True
    Label15.Caption = "Du klarade det på " & mell & " sekunder!"
    score = mell
        If score < HS Then
            MsgBox "Wow, the best time!!" & vbCrLf & "New High Score: " & score
            SaveHS Str(score)
        End If
  LoadHS
End If
End Sub

Private Sub Form_Load()

hjärtan = 0
liv = 3
gActive = False
Label2.Caption = liv
Label1.Caption = hjärtan
Image1.Visible = False
Image2.Visible = False
Image4.Visible = False
Timer1.Enabled = False
Image6.Visible = False
Image8.Visible = False
Image9.Visible = False
Image10.Visible = False
Image7.Visible = False
Image11.Visible = False
Image12.Visible = False

LoadHS
End Sub






Private Sub Timer1_Timer()      ' Flyttar image4

Dim höjd4 As Single
Dim bredd4 As Single

höjd4 = Image4.Top + fy
bredd4 = Image4.Left + fx


If höjd4 > Frame1.Height - Image4.Height - 100 Then
    fy = -fy                            'Image4 vänder uppåt
    höjd4 = Frame1.Height - Image4.Height - 100
End If

If höjd4 < 0 Then
    fy = -fy
    höjd4 = 0
End If

If bredd4 > Frame1.Width - Image4.Width Then
    fx = -fx
    bredd4 = Frame1.Width - Image4.Width
End If

If bredd4 < 0 Then
    fx = -fx
    bredd4 = 0
End If

Image4.Move bredd4, höjd4           'Kollar om image1 + 4 överlappar

If obk(Image1, Image4) Then         'när image4 flyttas
    liv = liv - 1
    Label2.Caption = liv
    bredd4 = Rnd * (Frame1.Width - Image4.Width - 100)
    Image4.Move bredd4, höjd4
End If

If liv < 1 Then
    Call dead
End If

slut = Timer
mell = slut - birj
stid = mell * 1
mell = stid / 1
Label14.Caption = mell

End Sub

'Man anropar denna funktion för att
'kolla om två objekt överlappar varandra

Private Function obk(c1 As Object, c2 As Object) As Boolean

Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single
Dim x3 As Single
Dim x4 As Single
Dim y3 As Single
Dim y4 As Single

x1 = c1.Left                  'Ger objekten värden och skickar
x2 = c1.Left + c1.Width       'värdena till funktionen r1
y1 = c1.Top
y2 = c1.Top + c1.Height

x3 = c2.Left
x4 = c2.Left + c2.Width
y3 = c2.Top
y4 = c2.Top + c2.Height
obk = r1(x1, x2, y1, y2, x3, x4, y3, y4)
End Function


Private Sub Timer2_Timer()
Image7.Move _
Rnd * (Frame1.Width - Image7.Width - RamOffset * 2), _
Rnd * (Frame1.Height - Image7.Height - RamOffset * 2)



End Sub

Private Sub Timer3_Timer()
Image12.Move _
Rnd * (Frame1.Width - Image12.Width - RamOffset * 2), _
Rnd * (Frame1.Height - Image12.Height - RamOffset * 2)
End Sub
Sub slutt()     'tryckt på F2
gActive = False

hjärtan = 0
liv = 3
gActive = False
Label2.Caption = liv
Label1.Caption = hjärtan
Image1.Visible = False
Image2.Visible = False
Image4.Visible = False
Timer1.Enabled = False
Image7.Visible = False
Image11.Visible = False
Image12.Visible = False
End Sub
Sub UpdateHS(sc As Single)
If sc < HS Then
   HS = sc
   SaveHS (HS)
End If
End Sub

Sub SaveHS(Scre As String)
    Open App.Path & "\hs.ini" For Output As #1
        Print #1, Str(Scre)
    Close #1
End Sub

Sub LoadHS()
    If Dir(App.Path & "\hs.ini") = "" Then SaveHS ("200")
    Open App.Path & "\hs.ini" For Input As #1
        Line Input #1, readln
    Close #1
    HS = readln
    Label17 = "High Score: " & readln & " Sek"
End Sub
Sub dead()


If liv < 1 Then
gActive = False
flytt = 20
hjärtan = 0
liv = 3
gActive = False
Image6.Visible = True
Label2.Caption = liv
Label1.Caption = hjärtan
Image1.Visible = False
Image2.Visible = False
Image4.Visible = False
Timer1.Enabled = False
Label7.Caption = "Hejsan dödis..."
Image8.Visible = True
Image9.Visible = True
Image10.Visible = True
Image11.Visible = False
Image12.Visible = False
Image7.Visible = False
Timer1.Enabled = False
Label14.Visible = False
End If
End Sub
