VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Yatzy 2002"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegler 
      BackColor       =   &H00FF8080&
      Caption         =   "Spelregler"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAvsluta 
      BackColor       =   &H00C0C000&
      Caption         =   "Avsluta"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdKasta 
      BackColor       =   &H000080FF&
      Caption         =   "Kasta tärningarna!"
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      Caption         =   "© 2002 Sandberg Productions Ltd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Caption         =   "YATZY!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   975
      Left            =   1800
      TabIndex        =   35
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Klicka sedan i formuläret till vänster för att bestämma vad du vill lägga dina poäng på"
      Height          =   495
      Left            =   1680
      TabIndex        =   33
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Klicka på tärningarna om du vill spara dem"
      Height          =   375
      Left            =   1680
      TabIndex        =   32
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Line Line8 
      X1              =   480
      X2              =   1440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblNamn 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   740
      TabIndex        =   30
      Top             =   960
      Width           =   975
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   1440
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   1440
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   1440
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   1440
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   1440
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblTotsumma 
      BackColor       =   &H0080C0FF&
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   29
      Top             =   5280
      Width           =   495
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   720
      Y1              =   1200
      Y2              =   5520
   End
   Begin VB.Label lblYatzy 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   28
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblChans 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   27
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lblKåk 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   26
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label lblSS 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label lblLS 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblFyrtal 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblTretal 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblTvåPar 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblEttPar 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblBonus 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblSiffersumma 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080C0FF&
      Caption         =   "Summa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblSexor 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblFemmor 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblFyror 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblTreor 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblTvåor 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblEttor 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Kåk Chans YATZY "
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
      Left            =   0
      TabIndex        =   10
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "LS SS"
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
      Left            =   420
      TabIndex        =   9
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "1 par 2 par 3-tal 4-tal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Summa Bonus "
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
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "1 2 3 4 5 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   585
      TabIndex        =   6
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label lblFyra 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblFem 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblTre 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblTvå 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblEtt 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ett As Integer, två As Integer, tre As Integer, fyra As Integer, fem As Integer, klicknummer As Integer

Private Sub cmdAvsluta_Click()
Dim svar As String
svar = MsgBox("Är det säkert att du vill avsluta?", vbQuestion + vbYesNo, "AVSLUTA")
  If svar = vbYes Then
    End
  ElseIf svar = vbNo Then
    Exit Sub
  End If
End Sub

Private Sub cmdKasta_Click()
klicknummer = klicknummer + 1
If klicknummer = 3 Then 'Efter tre kast får man inte kasta mer
  cmdKasta.Enabled = False
  klicknummer = 0
End If

lblEtt.Enabled = True 'Tärningarna blir klickbara så att deras värde kan sparas
lblTvå.Enabled = True
lblTre.Enabled = True
lblFyra.Enabled = True
lblFem.Enabled = True

  If lblEtt.BackColor = &H8080FF Then 'Om tärningen har denna bakgrundsfärg ska den slås om, om den inte har det så händer ingenting
    ett = Int(Rnd * 6) + 1 'Rnd * 6 ger ett tal från 0-5 så därför adderas 1 så det blir 1-6
    lblEtt = ett
  End If
  If lblTvå.BackColor = &H8080FF Then
    två = Int(Rnd * 6) + 1
    lblTvå = två
  End If
  If lblTre.BackColor = &H8080FF Then
    tre = Int(Rnd * 6) + 1
    lblTre = tre
  End If
  If lblFyra.BackColor = &H8080FF Then
    fyra = Int(Rnd * 6) + 1
    lblFyra = fyra
  End If
  If lblFem.BackColor = &H8080FF Then
    fem = Int(Rnd * 6) + 1
    lblFem = fem
  End If
  
  If lblEttor.BackColor = &H80C0FF Then    'om bakgrunden är orange så ska det gå att klicka på labeln
    lblEttor.Enabled = True
  End If
  If lblTvåor.BackColor = &H80C0FF Then
    lblTvåor.Enabled = True
  End If
  If lblTreor.BackColor = &H80C0FF Then
    lblTreor.Enabled = True
  End If
  If lblFyror.BackColor = &H80C0FF Then
    lblFyror.Enabled = True
  End If
  If lblFemmor.BackColor = &H80C0FF Then
    lblFemmor.Enabled = True
  End If
  If lblSexor.BackColor = &H80C0FF Then
    lblSexor.Enabled = True
  End If
  If lblEttPar.BackColor = &H80C0FF Then
    lblEttPar.Enabled = True
  End If
  If lblTvåPar.BackColor = &H80C0FF Then
    lblTvåPar.Enabled = True
  End If
  If lblTretal.BackColor = &H80C0FF Then
    lblTretal.Enabled = True
  End If
  If lblFyrtal.BackColor = &H80C0FF Then
    lblFyrtal.Enabled = True
  End If
  If lblLS.BackColor = &H80C0FF Then
    lblLS.Enabled = True
  End If
  If lblSS.BackColor = &H80C0FF Then
    lblSS.Enabled = True
  End If
  If lblKåk.BackColor = &H80C0FF Then
    lblKåk.Enabled = True
  End If
  If lblChans.BackColor = &H80C0FF Then
    lblChans.Enabled = True
  End If
  If lblYatzy.BackColor = &H80C0FF Then
    lblYatzy.Enabled = True
  End If
End Sub


Private Sub cmdRegler_Click()
  MsgBox "Yatzy spelas på följande sätt:" & Chr(13) & Chr(13) & "Högst upp i formuläret sparar man på ettor till sexor." & Chr(13) & "För att få en bonus på femtio poäng så måste du få tre av varje nummer, det blir 63 totalt." & Chr(13) & Chr(13) & "Ett par till fyrtal innebär samma sak som i poker." & Chr(13) & Chr(13) & "LS står för Liten Stege, och det betyder 1-5. Följaktligen är SS Stor Stege, 2-6." & Chr(13) & Chr(13) & "Kåk betyder tre av en sort och två av en annan." & Chr(13) & "På Chans sätts bara poängsumman av dina tärningar ihop." & Chr(13) & Chr(13) & "YATZY innebär att alla dina tärningar visar samma sak. Lyckas du med detta får du 50 poäng!" & Chr(13) & Chr(13) & "Lycka till!", vbInformation, "REGLER"
End Sub

Private Sub Form_Load()
Dim namn As String
Randomize
namn = InputBox("Välkommen till Yatzy 2002! Skriv i ditt namn:", "VÄLKOMMEN!")
lblNamn = namn 'Ditt namn fylls i här
End Sub

Private Sub lblChans_Click()
Dim svar As String
svar = MsgBox("Vill du sätta " & ett + två + tre + fyra + fem & " på Chans?", vbQuestion + vbYesNo, "CHANS?") 'För att man ska vara säker på vad man sätter på Chans kommer denna dialogruta upp
  If svar = vbYes Then
    lblChans = ett + två + tre + fyra + fem
    lblChans.BackColor = &HC0E0FF
    KnappLås
    Omställ
    RäknaTotSumma
  Else
  End If
End Sub

Private Sub lblEtt_Click()
  If lblEtt.BackColor = &H8080FF Then 'När du klickar på en tärning så ändras dess bakgrundsfärg, vilket gör att tärningen "låses" eller "låses upp"
    lblEtt.BackColor = &HFF8080
  ElseIf lblEtt.BackColor = &HFF8080 Then
    lblEtt.BackColor = &H8080FF
  End If 'Denna kod repeteras för alla tärningar
End Sub

Private Sub lblEttor_Click()
Dim svar As String
lblEttor.Caption = "0"
  If ett = 1 Then
    lblEttor = lblEttor + 1 'Varje tärning gås igenom, och för varje tärning som är en etta så ökar summan med ett
  End If
  If två = 1 Then
    lblEttor = lblEttor + 1
  End If
  If tre = 1 Then
    lblEttor = lblEttor + 1
  End If
  If fyra = 1 Then
    lblEttor = lblEttor + 1
  End If
  If fem = 1 Then
    lblEttor = lblEttor + 1
  End If
  
  If lblEttor = 0 Then 'om man vill stryka
    svar = MsgBox("Vill du stryka ettorna?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbNo Then
      lblEttor = ""
      Exit Sub
    End If
  End If
  
lblEttor.BackColor = &HC0E0FF 'knappen "låses" på detta sätt, så att den inte aktiveras igen
KnappLås
Omställ
RäknaSumma 'Denna kod ser likadan ut för ettor till sexor
End Sub

Private Sub lblEttPar_Click()
Dim temp1 As Integer, temp2 As Integer, svar As String
temp1 = 0
temp2 = 0

  If ett = två Then 'denna lilla kod ser om du har ett par och väljer ut det par som är större
    If temp1 = 0 Then
      temp1 = ett + två
    Else 'om variabeln temp1 är använd (dvs ett par har upptäckts tidigare) så sparas det andra paret i temp2
      temp2 = ett + två
    End If
  End If
  If ett = tre Then
    If temp1 = 0 Then
      temp1 = ett + tre
    Else
      temp2 = ett + tre
    End If
  End If
  If ett = fyra Then
    If temp1 = 0 Then
      temp1 = ett + fyra
    Else
      temp2 = ett + fyra
    End If
  End If
  If ett = fem Then
    If temp1 = 0 Then
      temp1 = ett + fem
    Else
      temp2 = ett + fem
    End If
  End If
  If två = tre Then
    If temp1 = 0 Then
      temp1 = två + tre
    Else
      temp2 = två + tre
    End If
  End If
  If två = fyra Then
    If temp1 = 0 Then
      temp1 = två + fyra
    Else
      temp2 = två + fyra
    End If
  End If
  If två = fem Then
    If temp1 = 0 Then
      temp1 = två + fem
    Else
      temp2 = två + fem
    End If
  End If
  If tre = fyra Then
    If temp1 = 0 Then
      temp1 = tre + fyra
    Else
      temp2 = tre + fyra
    End If
  End If
  If tre = fem Then
    If temp1 = 0 Then
      temp1 = tre + fem
    Else
      temp2 = tre + fem
    End If
  End If
  If fyra = fem Then
    If temp1 = 0 Then
      temp1 = fyra + fem
    Else
      temp2 = fyra + fem
    End If
  End If
  
    If temp1 > temp2 Then 'Det största paret väljs ut här
      lblEttPar = temp1
    ElseIf temp2 > temp1 Then
      lblEttPar = temp2
    End If
    
    If temp1 = 0 Then 'om du inte har några par
      svar = MsgBox("Vill du stryka ett par?", vbQuestion + vbYesNo, "STRYKA?")
      If svar = vbYes Then
        lblEttPar = 0
        lblEttPar.BackColor = &HC0E0FF
        KnappLås
      ElseIf svar = vbNo Then
        Exit Sub
      End If
    End If
    
lblEttPar.BackColor = &HC0E0FF
KnappLås
Omställ
RäknaTotSumma
End Sub

Private Sub lblFem_Click()
  If lblFem.BackColor = &H8080FF Then
    lblFem.BackColor = &HFF8080
  ElseIf lblFem.BackColor = &HFF8080 Then
    lblFem.BackColor = &H8080FF
  End If
End Sub

Private Sub lblFemmor_Click()
Dim svar As String
lblFemmor.Caption = "0"
  If ett = 5 Then
    lblFemmor = lblFemmor + 5
  End If
  If två = 5 Then
    lblFemmor = lblFemmor + 5
  End If
  If tre = 5 Then
    lblFemmor = lblFemmor + 5
  End If
  If fyra = 5 Then
    lblFemmor = lblFemmor + 5
  End If
  If fem = 5 Then
    lblFemmor = lblFemmor + 5
  End If
  
  If lblFemmor = 0 Then
    svar = MsgBox("Vill du stryka femmorna?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbNo Then
      lblFemmor = ""
      Exit Sub
    End If
  End If
  
lblFemmor.BackColor = &HC0E0FF
KnappLås
Omställ
RäknaSumma
End Sub

Private Sub lblFyra_Click()
  If lblFyra.BackColor = &H8080FF Then
    lblFyra.BackColor = &HFF8080
  ElseIf lblFyra.BackColor = &HFF8080 Then
    lblFyra.BackColor = &H8080FF
  End If
End Sub

Private Sub lblFyror_Click()
Dim svar As String
lblFyror.Caption = "0"
  If ett = 4 Then
    lblFyror = lblFyror + 4
  End If
  If två = 4 Then
    lblFyror = lblFyror + 4
  End If
  If tre = 4 Then
    lblFyror = lblFyror + 4
  End If
  If fyra = 4 Then
    lblFyror = lblFyror + 4
  End If
  If fem = 4 Then
    lblFyror = lblFyror + 4
  End If
  
  If lblFyror = 0 Then
    svar = MsgBox("Vill du stryka fyrorna?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbNo Then
      lblFyror = ""
      Exit Sub
    End If
  End If
  
lblFyror.BackColor = &HC0E0FF
KnappLås
Omställ
RäknaSumma
End Sub

Private Sub lblFyrtal_Click()
Dim svar As String
  If (ett = två And två = tre And tre = fyra) Then 'räknar ut fyrtal enligt samma princip som tretal
    lblFyrtal = ett + två + tre + fyra
    lblFyrtal.BackColor = &HC0E0FF
    KnappLås
    Omställ
    RäknaTotSumma
  ElseIf (ett = två And två = tre And tre = fem) Then
    lblFyrtal = ett + två + tre + fem
    lblFyrtal.BackColor = &HC0E0FF
    KnappLås
    Omställ
    RäknaTotSumma
  ElseIf (ett = två And två = fyra And fyra = fem) Then
    lblFyrtal = ett + två + fyra + fem
    lblFyrtal.BackColor = &HC0E0FF
    KnappLås
    Omställ
    RäknaTotSumma
  ElseIf (ett = tre And tre = fyra And fyra = fem) Then
    lblFyrtal = ett + tre + fyra + fem
    lblFyrtal.BackColor = &HC0E0FF
    KnappLås
    Omställ
    RäknaTotSumma
  ElseIf (två = tre And tre = fyra And fyra = fem) Then
    lblFyrtal = två + tre + fyra + fem
    lblFyrtal.BackColor = &HC0E0FF
    KnappLås
    Omställ
    RäknaTotSumma
  End If
  
  If lblFyrtal = "" Then
    svar = MsgBox("Vill du stryka fyrtal?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbYes Then
      lblFyrtal = 0
      lblFyrtal.BackColor = &HC0E0FF
      KnappLås
      Omställ
      RäknaSumma
    End If
  End If
End Sub

Private Sub lblKåk_Click()
Dim svar As String
  If ett = två And tre = fyra And fyra = fem Then 'Kåk kontrolleras enligt samma princip som tretal och fyrtal
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf ett = tre And två = fyra And fyra = fem Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf ett = fyra And två = tre And tre = fem Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf ett = fem And två = tre And tre = fyra Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf två = tre And ett = fyra And fyra = fem Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf två = fyra And ett = tre And tre = fem Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf två = fem And ett = tre And tre = fyra Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf tre = fyra And ett = två And två = fem Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf tre = fem And ett = två And två = fyra Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  ElseIf fyra = fem And ett = två And två = tre Then
    lblKåk = ett + två + tre + fyra + fem
    Omställ
    RäknaTotSumma
    lblKåk.BackColor = &HC0E0FF
    KnappLås
  End If
  
  If lblKåk = "" Then 'om du vill stryka kåk
    svar = MsgBox("Vill du stryka kåk?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbYes Then
      lblKåk = 0
      lblKåk.BackColor = &HC0E0FF
      KnappLås
      Omställ
      RäknaTotSumma
    End If
  End If
End Sub

Private Sub lblLS_Click()
Dim svar As String
  If (ett + två + tre + fyra + fem = 15) And ett < 6 And två < 6 And tre < 6 And fyra < 6 And fem < 6 And ett <> två And ett <> tre And ett <> fyra And ett <> fem And två <> tre And två <> fyra And två <> fem And tre <> fyra And tre <> fem And fyra <> fem Then
    lblLS = 15 'raden ovanför avgör genom uteslutningsmetoden om man har liten stege
    lblLS.BackColor = &HC0E0FF
    KnappLås
    Omställ
    RäknaTotSumma
  Else 'Om man skulle vilja stryka stegen
    svar = MsgBox("Vill du stryka Liten Stege?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbYes Then
      lblLS = 0
      Omställ
      lblLS.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
End Sub

Private Sub lblSexor_Click()
Dim svar As String
lblSexor.Caption = "0"
  If ett = 6 Then
    lblSexor = lblSexor + 6
  End If
  If två = 6 Then
    lblSexor = lblSexor + 6
  End If
  If tre = 6 Then
    lblSexor = lblSexor + 6
  End If
  If fyra = 6 Then
    lblSexor = lblSexor + 6
  End If
  If fem = 6 Then
    lblSexor = lblSexor + 6
  End If
  
  If lblSexor = 0 Then
    svar = MsgBox("Vill du stryka sexorna?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbNo Then
      lblSexor = ""
      Exit Sub
    End If
  End If
  
lblSexor.BackColor = &HC0E0FF
KnappLås
Omställ
RäknaSumma
End Sub

Private Sub lblSiffersumma_Change()
  RäknaTotSumma
End Sub

Private Sub lblSS_Click()
Dim svar As Integer
  If (ett + två + tre + fyra + fem = 20) And ett > 1 And två > 1 And tre > 1 And fyra > 1 And fem > 1 And ett <> två And ett <> tre And ett <> fyra And ett <> fem And två <> tre And två <> fyra And två <> fem And tre <> fyra And tre <> fem And fyra <> fem Then
    lblSS = 20 'denna kod fungerar precis likadant som den för liten stege
    lblSS.BackColor = &HC0E0FF
    KnappLås
    Omställ
    RäknaTotSumma
  Else
    svar = MsgBox("Vill du stryka Stor Stege?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbYes Then
      lblSS = 0
      Omställ
      lblSS.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
End Sub

Private Sub lblTotsumma_Click()

End Sub

Private Sub lblTre_Click()
  If lblTre.BackColor = &H8080FF Then
    lblTre.BackColor = &HFF8080
  ElseIf lblTre.BackColor = &HFF8080 Then
    lblTre.BackColor = &H8080FF
  End If
End Sub

Private Sub lblTreor_Click()
Dim svar As String
lblTreor.Caption = "0"
  If ett = 3 Then
    lblTreor = lblTreor + 3
  End If
  If två = 3 Then
    lblTreor = lblTreor + 3
  End If
  If tre = 3 Then
    lblTreor = lblTreor + 3
  End If
  If fyra = 3 Then
    lblTreor = lblTreor + 3
  End If
  If fem = 3 Then
    lblTreor = lblTreor + 3
  End If
  
  If lblTreor = 0 Then
    svar = MsgBox("Vill du stryka treorna?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbNo Then
      lblTreor = ""
      Exit Sub
    End If
  End If
  
lblTreor.BackColor = &HC0E0FF
KnappLås
Omställ
RäknaSumma
End Sub

Private Sub lblTretal_Click()
Dim svar As String
  If (ett = två And två = tre) Then 'denna svagt förvirrande if-sats märker om du får tretal och lägger då ihop rätt tal därefter
    lblTretal = ett + två + tre
  ElseIf (ett = två And två = fyra) Then
    lblTretal = ett + två + fyra
  ElseIf (ett = två And två = fem) Then
    lblTretal = ett + två + fem
  ElseIf (ett = tre And tre = fyra) Then
    lblTretal = ett + tre + fyra
  ElseIf (ett = tre And tre = fem) Then
    lblTretal = ett + tre + fem
  ElseIf (ett = fyra And fyra = fem) Then
    lblTretal = ett + fyra + fem
  ElseIf (två = tre And tre = fyra) Then
    lblTretal = två + tre + fyra
  ElseIf (två = tre And tre = fem) Then
    lblTretal = två + tre + fem
  ElseIf (två = fyra And fyra = fem) Then
    lblTretal = två + fyra + fem
  ElseIf (tre = fyra And fyra = fem) Then
    lblTretal = tre + fyra + fem
  End If
  
  If lblTretal = "" Then 'Om du vill stryka
    svar = MsgBox("Vill du stryka tretal?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbYes Then
      lblTretal = 0
    ElseIf svar = vbNo Then
      Exit Sub
    End If
  End If
  
lblTretal.BackColor = &HC0E0FF
KnappLås
Omställ
RäknaTotSumma
End Sub

Private Sub lblTvå_Click()
  If lblTvå.BackColor = &H8080FF Then
    lblTvå.BackColor = &HFF8080
  ElseIf lblTvå.BackColor = &HFF8080 Then
    lblTvå.BackColor = &H8080FF
  End If
End Sub

Private Sub lblTvåor_Click()
Dim svar As String
lblTvåor.Caption = "0"
  If ett = 2 Then
    lblTvåor = lblTvåor + 2
  End If
  If två = 2 Then
    lblTvåor = lblTvåor + 2
  End If
  If tre = 2 Then
    lblTvåor = lblTvåor + 2
  End If
  If fyra = 2 Then
    lblTvåor = lblTvåor + 2
  End If
  If fem = 2 Then
    lblTvåor = lblTvåor + 2
  End If
  
  If lblTvåor = 0 Then
    svar = MsgBox("Vill du stryka tvåorna?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbNo Then
      lblTvåor = ""
      Exit Sub
    End If
  End If
  
lblTvåor.BackColor = &HC0E0FF
KnappLås
Omställ
RäknaSumma
End Sub

Private Sub Omställ() 'denna subrutin åberopas varje gång du klickar i ett värde så att tärningarna ställs om
cmdKasta.Enabled = True
lblEtt.BackColor = &H8080FF
lblTvå.BackColor = &H8080FF
lblTre.BackColor = &H8080FF
lblFyra.BackColor = &H8080FF
lblFem.BackColor = &H8080FF
lblEtt.Enabled = False
lblTvå.Enabled = False
lblTre.Enabled = False
lblFyra.Enabled = False
lblFem.Enabled = False
klicknummer = 0
End Sub

Private Sub RäknaSumma()
Dim temp1 As Integer, temp2 As Integer, temp3 As Integer, temp4 As Integer, temp5 As Integer, temp6 As Integer
  lblSiffersumma = "0"
  If lblEttor = "" Then 'Eftersom VB inte kan räkna med "" så måste det göras om till 0
    temp1 = 0
  Else
    temp1 = lblEttor
  End If
  If lblTvåor = "" Then
    temp2 = 0
  Else
    temp2 = lblTvåor
  End If
  If lblTreor = "" Then
    temp3 = 0
  Else
    temp3 = lblTreor
  End If
  If lblFyror = "" Then
    temp4 = 0
  Else
    temp4 = lblFyror
  End If
  If lblFemmor = "" Then
    temp5 = 0
  Else
    temp5 = lblFemmor
  End If
  If lblSexor = "" Then
    temp6 = 0
  Else
    temp6 = lblSexor
  End If
  
  lblSiffersumma = temp1 + temp2 + temp3 + temp4 + temp5 + temp6
  If lblSiffersumma >= 63 Then
    lblBonus = "50"
  End If
End Sub

Private Sub lblTvåPar_Click()
Dim svar As String
  If ett = två And tre = fyra Then
    If ett = tre Then 'här undersöks det om paren har samma värde, vilket de ju inte får ha...
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + tre + fyra
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = två And tre = fem Then
    If ett = tre Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + tre + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = två And fyra = fem Then
    If ett = fyra Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = tre And två = fyra Then
    If ett = två Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + tre + fyra
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = tre And två = fem Then
    If ett = två Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + tre + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = tre And fyra = fem Then
    If ett = fyra Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + tre + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = fyra And två = tre Then
    If ett = två Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + tre + fyra
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = fyra And två = fem Then
    If ett = två Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = fyra And tre = fem Then
    If ett = tre Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + tre + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = fem And två = tre Then
    If ett = två Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + tre + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = fem And två = fyra Then
    If ett = två Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = tre And två = fyra Then
    If ett = två Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + tre + fyra
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = två And tre = fyra Then
    If ett = tre Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + två + tre + fyra
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If ett = fem And tre = fyra Then
    If ett = tre Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = ett + tre + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If två = tre And fyra = fem Then
    If två = fyra Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = två + tre + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If två = fyra And tre = fem Then
    If två = tre Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = två + tre + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  If två = fem And tre = fyra Then
    If två = tre Then
      MsgBox "De två paren får inte ha samma värde!", vbExclamation, "OBS!"
    Else
      lblTvåPar = två + tre + fyra + fem
      Omställ
      lblTvåPar.BackColor = &HC0E0FF
      KnappLås
      RäknaTotSumma
    End If
  End If
  
  If lblTvåPar = "" Then 'om du inte har några par
    svar = MsgBox("Vill du stryka två par?", vbQuestion + vbYesNo, "STRYKA?")
      If svar = vbYes Then
        lblTvåPar = 0
        Omställ
        lblTvåPar.BackColor = &HC0E0FF
        KnappLås
        RäknaTotSumma
      End If
  End If
End Sub

Private Sub lblYatzy_Click()
Dim svar As String
  If ett = två And två = tre And tre = fyra And fyra = fem Then
    lblYatzy = 50
    RäknaTotSumma
    lblYatzy.BackColor = &HC0E0FF
    KnappLås
    Omställ
  Else
    svar = MsgBox("Vill du stryka Yatzy?", vbQuestion + vbYesNo, "STRYKA?")
    If svar = vbYes Then
      lblYatzy = 0
      RäknaTotSumma
      lblYatzy.BackColor = &HC0E0FF
      KnappLås
      Omställ
    End If
  End If
End Sub

Private Sub RäknaTotSumma()
Dim tp1 As Integer, tp2 As Integer, tp3 As Integer, tp4 As Integer, tp5 As Integer, tp6 As Integer, tp7 As Integer, tp8 As Integer, tp9 As Integer, tp10 As Integer, tp11 As Integer
If lblEttor = "" Or lblTvåor = "" Or lblTreor = "" Or lblFyror = "" Or lblFemmor = "" Or lblSexor = "" Or lblEttPar = "" Or lblTvåPar = "" Or lblTretal = "" Or lblFyrtal = "" Or lblLS = "" Or lblSS = "" Or lblKåk = "" Or lblChans = "" Or lblYatzy = "" Then 'om allt INTE är ifyllt
  If lblSiffersumma = "" Then 'VB kan ju inte räkna med "" så här görs det om till 0
    tp1 = 0
  Else
    tp1 = lblSiffersumma
  End If
  If lblBonus = "" Then
    tp2 = 0
  Else
    tp2 = lblBonus
  End If
  If lblEttPar = "" Then
    tp3 = 0
  Else
    tp3 = lblEttPar
  End If
  If lblTvåPar = "" Then
    tp4 = 0
  Else
    tp4 = lblTvåPar
  End If
  If lblTretal = "" Then
    tp5 = 0
  Else
    tp5 = lblTretal
  End If
  If lblFyrtal = "" Then
    tp6 = 0
  Else
    tp6 = lblFyrtal
  End If
  If lblLS = "" Then
    tp7 = 0
  Else
    tp7 = lblLS
  End If
  If lblSS = "" Then
    tp8 = 0
  Else
    tp8 = lblSS
  End If
  If lblKåk = "" Then
    tp9 = 0
  Else
    tp9 = lblKåk
  End If
  If lblChans = "" Then
    tp10 = 0
  Else
    tp10 = lblChans
  End If
  If lblYatzy = "" Then
    tp11 = 0
  Else
    tp11 = lblYatzy
  End If
  lblTotsumma = tp1 + tp2 + tp3 + tp4 + tp5 + tp6 + tp7 + tp8 + tp9 + tp10 + tp11 'summan av de tillfälliga variablerna sätts ihop
Else 'om allt ÄR ifyllt
Dim svar As String
  If lblNamn = "" Then 'om du inte fyllt i något namn
    svar = MsgBox("Du fick " & lblTotsumma & " poäng! Vill du spela igen?", vbInformation + vbYesNo, "Spelet slut")
    If svar = vbYes Then 'ladda om allting
      LaddaOm
    ElseIf svar = vbNo Then 'avsluta
      End
    End If
  Else 'om du fyllt i ett namn
    svar = MsgBox(lblNamn & " fick " & lblTotsumma & " poäng! Vill du spela igen?", vbInformation + vbYesNo, "Spelet slut")
    If svar = vbYes Then 'ladda om allting
      LaddaOm
    ElseIf svar = vbNo Then 'avsluta
      End
    End If
  End If
End If
End Sub


Private Sub LaddaOm() 'när du startar om körs denna subrutin
  lblEttor = ""
  lblEttor.Enabled = True
  lblEttor.BackColor = &H80C0FF
  lblTvåor = ""
  lblTvåor.Enabled = True
  lblTvåor.BackColor = &H80C0FF
  lblTreor = ""
  lblTreor.Enabled = True
  lblTreor.BackColor = &H80C0FF
  lblFyror = ""
  lblFyror.Enabled = True
  lblFyror.BackColor = &H80C0FF
  lblFemmor = ""
  lblFemmor.Enabled = True
  lblFemmor.BackColor = &H80C0FF
  lblSexor = ""
  lblSexor.Enabled = True
  lblSexor.BackColor = &H80C0FF
  lblSiffersumma = ""
  lblBonus = ""
  lblEttPar = ""
  lblEttPar.Enabled = True
  lblEttPar.BackColor = &H80C0FF
  lblTvåPar = ""
  lblTvåPar.Enabled = True
  lblTvåPar.BackColor = &H80C0FF
  lblTretal = ""
  lblTretal.Enabled = True
  lblTretal.BackColor = &H80C0FF
  lblFyrtal = ""
  lblFyrtal.Enabled = True
  lblFyrtal.BackColor = &H80C0FF
  lblLS = ""
  lblLS.Enabled = True
  lblLS.BackColor = &H80C0FF
  lblSS = ""
  lblSS.Enabled = True
  lblSS.BackColor = &H80C0FF
  lblKåk = ""
  lblKåk.Enabled = True
  lblKåk.BackColor = &H80C0FF
  lblChans = ""
  lblChans.Enabled = True
  lblChans.BackColor = &H80C0FF
  lblYatzy = ""
  lblYatzy.Enabled = True
  lblYatzy.BackColor = &H80C0FF
  lblTotsumma = ""
End Sub

Private Sub KnappLås() 'den här subrutinen ser till så att alla knappar blir icke-klickbara
  lblEttor.Enabled = False
  lblTvåor.Enabled = False
  lblTreor.Enabled = False
  lblFyror.Enabled = False
  lblFemmor.Enabled = False
  lblSexor.Enabled = False
  lblEttPar.Enabled = False
  lblTvåPar.Enabled = False
  lblTretal.Enabled = False
  lblFyrtal.Enabled = False
  lblLS.Enabled = False
  lblSS.Enabled = False
  lblKåk.Enabled = False
  lblChans.Enabled = False
  lblYatzy.Enabled = False
End Sub
