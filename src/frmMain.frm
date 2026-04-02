VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frmNew"
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   15750
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   15750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmInfoBall 
      BackColor       =   &H00FFFFFF&
      Caption         =   "info kugla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   12360
      TabIndex        =   12
      Top             =   360
      Width           =   2535
      Begin VB.Label lblBojaCounter 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   435
         Index           =   5
         Left            =   1680
         TabIndex        =   22
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label lblBojaCounter 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Index           =   4
         Left            =   1680
         TabIndex        =   21
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label lblBojaCounter 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Index           =   3
         Left            =   1680
         TabIndex        =   20
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblBojaCounter 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   435
         Index           =   2
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblBojaCounter 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   1
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblInfBall 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PINK:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   360
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblInfBall 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PLAVA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblInfBall 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ZELENA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblInfBall 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ŽUTA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   360
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblInfBall 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CRVENA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "info"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Timer tmrAutoPlay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   3360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cheat"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8040
      Width           =   615
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "<-- UNDO (zadnji potez)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   333
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   7935
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label lblJockerHint 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "1)desni klik na kuglu za odabir jocker kugle! 2)desni klik za akctiv/deactiv jocker kugle!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblJocker 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JOCKER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Line lineGrid 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   120
      X2              =   720
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Shape shpKuglaBonus 
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   240
      Shape           =   3  'Circle
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpKugla 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   0
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblOrientation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_BOX_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblBodoviTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1800"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpBodoviTemp 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Shape           =   2  'Oval
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblGameMode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GAME_MODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label txtBodovi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label2 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label lblKugla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu mnuGameMode 
      Caption         =   "Game Mode"
      Begin VB.Menu mnuGameModeSelect 
         Caption         =   "&1. Standard"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGameModeSelect 
         Caption         =   "&2. Continuous"
         Index           =   2
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuGameModeSelect 
         Caption         =   "&3. Shifter"
         Index           =   3
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuGameModeSelect 
         Caption         =   "&4. MegaShift"
         Index           =   4
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuOrientation 
      Caption         =   "Orientation"
      Begin VB.Menu mnuOrientationSelect 
         Caption         =   "&1. Portrait        (11 x 12)"
         Index           =   1
      End
      Begin VB.Menu mnuOrientationSelect 
         Caption         =   "&2. Landscape   (15 x 8)"
         Index           =   2
      End
      Begin VB.Menu mnuOrientationSelect 
         Caption         =   "&3. Box              (14 x 14)"
         Index           =   3
      End
      Begin VB.Menu mnuOrientationSelect 
         Caption         =   "&4. XPERIA X1 Style (16 x 11)"
         Index           =   4
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsGrid 
         Caption         =   "&1. Grid"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuOptionsJocker 
         Caption         =   "&2. Jocker"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfoBall 
         Caption         =   "&3. INFO: Ball"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuJocker 
      Caption         =   "Jocker"
      Begin VB.Menu mnuJockerBall 
         Caption         =   "Jocker Ball"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''
'NOTACIJA:
'   imena funkcija      =>  jedanDva()
'   imena varijabli:
'           lokalne     =>  jedan_dva
'           globalne    =>  JEDAN_DVA


Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
'sirina table po X osi
Dim TABLA_DUZINA_X As Byte '= 19

'visina table po Y osi
Dim TABLA_DUZINA_Y As Byte '= 10

Dim TABLA() As Long

'kopija table - sluzi za UNDO
Dim TABLA_COPY() As Long


'*********************************************************************

'tabla sa bonusnim kuglama (za Continous i MegaShift GAME_MODE)
Dim TABLA_BONUS() As Long
Dim TABLA_BONUS_COPY() As Long

'duzina table
Dim TABLA_BONUS_DUZINA() As Byte
Dim TABLA_BONUS_DUZINA_COPY() As Byte

'trenutna tabla -> ona koja se koristi
Dim TABLA_BONUS_TRENUTNA As Long
Dim TABLA_BONUS_TRENUTNA_LAST As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''


'maksimalan broj boja
Const BOJA_MAX  As Byte = 5

'deklaracija boja
Dim BOJA(BOJA_MAX) As Long

'deklaracija brojaca boja
'brojac boja istih kugla
Dim BOJA_COUNTER(BOJA_MAX) As Integer

'boja odabranih kugli
Const BOJA_SELECTED As Long = vbBlack

'deklaracija jocker boja
Dim JOCKER_BOJA(1) As Long
    '0  -   jocker boja (ne aktivna)
    '1  -   jocker boja (aktivna)

''''''''''''''''''''''''''''''''''''''''''''''''''''
'sluzi za oznacavanje kugla iste boje u nizu(koje su spojene)
Dim NIZ_KUGLA() As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''
'osvojeni bodovi u igri
Dim BODOVI As Integer

'trenutni moguci osvojeni bodovi u igri
Dim BODOVI_TEMP As Integer
'trenutni osvojeni bodovi u igri
Dim BODOVI_TEMP_OSVOJENI As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''
'ukupno game modova
Const GAME_MODE As Byte = 4
            '(1) Standard - obicna igra
            '(2) Continuous - ubaciva se random Y kugla kada se pojavi praznina
            '(3) Shifter - pomjeraju se kugle u desno (ne moze biti rupa izmedju kolona)
            '(4) MegaShift - (Continuous + Shifter) skupa

'tekstovi za pojedini game mod
Dim GAME_MODE_TEXT(GAME_MODE) As String

'privremeno cuva novi odabrani game mod
Dim GAME_MODE_SELECTED_NEW As Byte

''''''''''''''''''''''''''''''''''''''''''''''''''''

'kada se sruse kugle provjerava da li je oborena kolona
Dim KOLONA_OBORENA As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ORIENTATION As Byte
''''''''''''''''''''''''''''''''''''''''''''''''''''

'trenutni jocker
Dim JOCKER As Boolean

'novi jocker koji ce se koristiti u iducoj partiji
Dim JOCKER_NEW As Boolean

'da li je jocker u upotrebi
Dim JOCKER_IN_USE As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''
'kugla koja se kliknula
Dim KUGLA_TRENUTNA As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''
'koristi se za preskakanje izvodjenja koda u lblKugla_click
Dim SKIP As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''
'oznacava tipku misa
'   1   -   Left Button
'   2   -   Right Button
Dim KLIK As Byte
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim PREKID_APP As Byte
Const PREKID_APP_MAX As Byte = 100

Public Sub inicTabla()
    'inicijalizacija table
    'postavljanje vrijednosti na boju kugle
    '
    'Empty => nema kugle
    '
    Dim br_sl As Integer
    Dim boja_sl As Integer
    
    Dim i As Integer
    Dim br_count(BOJA_MAX) As Integer
    For i = 1 To BOJA_MAX
        br_count(i) = 0
    Next i
 
    ''''radi sigurnosti pokretanja inicijalizacije table uvodi se sigurnosni brojac
    Const sigurnosni_brojac_max As Integer = 20000
    Dim sigurnosni_brojac As Integer
    
    
    'ubacivanje kugli u tablu
    'random kugla u random polje
    For i = getDuzinaTable To 1 Step -1
        sigurnosni_brojac = 0
        Do
            sigurnosni_brojac = sigurnosni_brojac + 1
            
            boja_sl = getSlucajniBroj(1, BOJA_MAX)
            If br_count(boja_sl) < BOJA_COUNTER(boja_sl) Then
                br_count(boja_sl) = br_count(boja_sl) + 1
                TABLA(getOsaX(i), getOsaY(i)) = BOJA(boja_sl)
            End If
            
            If sigurnosni_brojac >= sigurnosni_brojac_max Then Exit Do
        
        Loop Until TABLA(getOsaX(i), getOsaY(i)) <> Empty
        
        If sigurnosni_brojac >= sigurnosni_brojac_max Then Exit For
    Next i
    
    'ako je greska
    If sigurnosni_brojac >= sigurnosni_brojac_max Then
        PREKID_APP = PREKID_APP + 1
        Debug.Print "SYSTEM FAILURE" & "(" & PREKID_APP & ")"
        
        'ako je ponovi 100 puta
        If PREKID_APP >= PREKID_APP_MAX Then
            '// LOST number ;) //
            MsgBox "FATAL ERROR!" & vbCrLf & vbCrLf & "code: 4 8 15 16 23 42" & vbCrLf & vbCrLf & "SYSTEM FAILURE", vbCritical
            Unload Me
        End If
        
        Form_Load
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'rastrkavanje random kugle sa random drugom kuglom
    'broj iteracija je izmedju 15% duzine table i duzine tabel
    
    Dim djeljenik As Byte: djeljenik = getSlucajniBroj(2, 5)
    Dim iteracija As Byte: iteracija = getSlucajniBroj(0, (getDuzinaTable / djeljenik))
    Dim kugla1, kugla2 As Integer
    For i = 1 To iteracija
        kugla1 = getSlucajniBroj(1, getDuzinaTable)
        Do
            kugla2 = getSlucajniBroj(1, getDuzinaTable)
        Loop Until kugla1 <> kugla2
    
        swapKugla kugla1, kugla2
    Next i
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ogranicavanja maksimalnog broja kugla (spojenih)
    Dim kugle_min As Byte: kugle_min = getSlucajniBroj(5, 12)
    Dim kugle_max As Byte: kugle_max = getSlucajniBroj(5, 12)
    Dim max_kugla_skupa As Byte: max_kugla_skupa = getSlucajniBroj(kugle_min, kugle_max)
    
    'ogranicenje minimuma 2 => AKO JE 1 TREBA PUNOOOOO VREMENA RADI ALGORITMA:-(
    If max_kugla_skupa < 2 Then max_kugla_skupa = 2
    
    Dim brojac As Integer
    Dim j As Integer
    Dim PREKID As Byte
    
    Do
        PREKID = getDuzinaTable
        brojac = 0
        While brojac < getDuzinaTable
            brojac = brojac + 1
            setNizKugla (brojac)
            If getNizKuglaDuzina > max_kugla_skupa Then
                PREKID = PREKID - 1
                For j = max_kugla_skupa To getNizKuglaDuzina
                    Do
                        br_sl = getSlucajniBroj(1, getDuzinaTable)
                    Loop Until TABLA(getOsaX(br_sl), getOsaY(br_sl)) <> TABLA(getOsaX(getNizKugla(j)), getOsaY(getNizKugla(j)))
                    swapKugla getNizKugla(j), br_sl
                Next j
            End If
        Wend
    Loop Until PREKID = getDuzinaTable
End Sub

Public Sub showTabla()
    '
    'crtanje(prikaz) table
    '
    
    Dim i, j As Integer
    For i = 1 To TABLA_DUZINA_X
        For j = 1 To TABLA_DUZINA_Y
            If TABLA(i, j) = Empty Then
            'ako kugla nije inicijalizirana => HIDE
                shpKugla(getKuglaIndex(i, j)).Visible = False
                lblKugla(getKuglaIndex(i, j)).Visible = False
            Else
            'ako je inicijalizirana => SHOW
                shpKugla(getKuglaIndex(i, j)).Visible = True
                shpKugla(getKuglaIndex(i, j)).FillColor = TABLA(i, j)
                
                lblKugla(getKuglaIndex(i, j)).BackStyle = 0
                lblKugla(getKuglaIndex(i, j)).Visible = True
            End If
        Next j
    Next i
End Sub

Public Function getKuglaIndex(ByVal osa_x As Byte, ByVal osa_y As Byte) As Integer
    '
    'vraca broj kugle(Index)
    '
    
    getKuglaIndex = (((osa_y - 1) * TABLA_DUZINA_X) + osa_x)
End Function

Public Function getKuglaGore(ByVal kugla As Integer) As Long
    '
    'vraca stanje kugle iznad
    'ako ima kugla iznad    => broj boje kugle
    '           ako nema    => 0
    '
    Dim osa_x As Byte: osa_x = getOsaX(kugla)
    Dim osa_y As Byte: osa_y = getOsaY(kugla)
    
    If osa_y = 1 Then
        'ako je u prvom redu
        getKuglaGore = 0
    Else
        'ako nema kugle iznad a nije u prvom redu
        If TABLA(osa_x, (osa_y - 1)) = Empty Then
            getKuglaGore = 0
            Exit Function
        End If
        'ako ima kugle
        getKuglaGore = TABLA(osa_x, (osa_y - 1))
    End If
End Function

Public Function getKuglaGoreIndex(ByVal kugla As Integer) As Integer
    '
    'vraca index kugle iznad
    'ako ima kugla iznad    => index kugle
    '           ako nema    => 0
    '
    Dim osa_x As Byte: osa_x = getOsaX(kugla)
    Dim osa_y As Byte: osa_y = getOsaY(kugla)
    
    If osa_y = 1 Then
        'ako je u prvom redu
        getKuglaGoreIndex = 0
    Else
        'ako nema kugle iznad a nije u prvom redu
        If TABLA(osa_x, (osa_y - 1)) = Empty Then
            getKuglaGoreIndex = 0
            Exit Function
        End If
        'ako ima kugle
        getKuglaGoreIndex = getKuglaIndex(osa_x, (osa_y - 1))
    End If
End Function

Public Function getKuglaDole(ByVal kugla As Integer) As Long
    '
    'vraca stanje kugle ispod
    'ako ima kugla ispod    => broj boje kugle
    '           ako nema    => 0
    '
    Dim osa_x As Byte: osa_x = getOsaX(kugla)
    Dim osa_y As Byte: osa_y = getOsaY(kugla)
    
    If osa_y = TABLA_DUZINA_Y Then
        'ako je u zadnjem redu
        getKuglaDole = 0
    Else
        'ako nema kugle ispod a nije u zadnjem redu
        If TABLA(osa_x, (osa_y + 1)) = Empty Then
            getKuglaDole = 0
            Exit Function
        End If
        'ako ima kugle
        getKuglaDole = TABLA(osa_x, (osa_y + 1))
    End If
End Function

Public Function getKuglaDoleIndex(ByVal kugla As Integer) As Integer
    '
    'vraca index kugle ispod
    'ako ima kugla ispod    => broj indexa kugle
    '           ako nema    => 0
    '
    Dim osa_x As Byte: osa_x = getOsaX(kugla)
    Dim osa_y As Byte: osa_y = getOsaY(kugla)
    
    If osa_y = TABLA_DUZINA_Y Then
        'ako je u zadnjem redu
        getKuglaDoleIndex = 0
    Else
        'ako nema kugle ispod a nije u zadnjem redu
        If TABLA(osa_x, (osa_y + 1)) = Empty Then
            getKuglaDoleIndex = 0
            Exit Function
        End If
        'ako ima kugle
        getKuglaDoleIndex = getKuglaIndex(osa_x, (osa_y + 1))
    End If
End Function

Public Function getKuglaDesno(ByVal kugla As Integer) As Long
    '
    'vraca stanje kugle desno
    'ako ima kugla desno    => broj boje kugle
    '           ako nema    => 0
    '
    Dim osa_x As Byte: osa_x = getOsaX(kugla)
    Dim osa_y As Byte: osa_y = getOsaY(kugla)
    
    If osa_x = TABLA_DUZINA_X Then
        'ako je u zadnjoj koloni
        getKuglaDesno = 0
    Else
        'ako nema kugle desno a nije u zadnjoj koloni
        If TABLA((osa_x + 1), osa_y) = Empty Then
            getKuglaDesno = 0
            Exit Function
        End If
        'ako ima kugle
        getKuglaDesno = TABLA((osa_x + 1), osa_y)
    End If
End Function

Public Function getKuglaDesnoIndex(ByVal kugla As Integer) As Long
    '
    'vraca index kugle desno
    'ako ima kugla desno    => broj indexa kugle
    '           ako nema    => 0
    '
    Dim osa_x As Byte: osa_x = getOsaX(kugla)
    Dim osa_y As Byte: osa_y = getOsaY(kugla)
    
    If osa_x = TABLA_DUZINA_X Then
        'ako je u zadnjoj koloni
        getKuglaDesnoIndex = 0
    Else
        'ako nema kugle desno a nije u zadnjoj koloni
        If TABLA((osa_x + 1), osa_y) = Empty Then
            getKuglaDesnoIndex = 0
            Exit Function
        End If
        'ako ima kugle
        getKuglaDesnoIndex = getKuglaIndex((osa_x + 1), osa_y)
    End If
End Function

Public Function getKuglaLjevo(ByVal kugla As Integer) As Long
    '
    'vraca stanje kugle ljevo
    'ako ima kugla ljevo    => broj boje kugle
    '           ako nema    => 0
    '
    Dim osa_x As Byte: osa_x = getOsaX(kugla)
    Dim osa_y As Byte: osa_y = getOsaY(kugla)
    
    If osa_x = 1 Then
        'ako je u prvoj koloni
        getKuglaLjevo = 0
    Else
        'ako nema kugle ljevo a nije u prvoj koloni
        If TABLA((osa_x - 1), osa_y) = Empty Then
            getKuglaLjevo = 0
            Exit Function
        End If
        'ako ima kugle
        getKuglaLjevo = TABLA((osa_x - 1), osa_y)
    End If
End Function

Public Function getKuglaLjevoIndex(ByVal kugla As Integer) As Integer
    '
    'vraca index kugle ljevo
    'ako ima kugla ljevo    => broj indexa kugle
    '           ako nema    => 0
    '
    Dim osa_x As Byte: osa_x = getOsaX(kugla)
    Dim osa_y As Byte: osa_y = getOsaY(kugla)
    
    If osa_x = 1 Then
        'ako je u prvoj koloni
        getKuglaLjevoIndex = 0
    Else
        'ako nema kugle ljevo a nije u prvoj koloni
        If TABLA((osa_x - 1), osa_y) = Empty Then
            getKuglaLjevoIndex = 0
            Exit Function
        End If
        'ako ima kugle
        getKuglaLjevoIndex = getKuglaIndex((osa_x - 1), osa_y)
    End If
End Function

Public Function getOsaX(ByVal kugla As Integer) As Byte
     '
     ' => osu X od kugle
     '
     Dim pozicija As Byte: pozicija = 0
     pozicija = (kugla - ((getOsaY(kugla) - 1) * TABLA_DUZINA_X))
     getOsaX = pozicija
End Function

Public Function getOsaY(ByVal kugla As Integer) As Byte
     '
     ' => osu Y od kugle
     '
     Dim pozicija As Byte: pozicija = 0
     pozicija = Fix(kugla / TABLA_DUZINA_X)
     If kugla Mod TABLA_DUZINA_X <> 0 Then pozicija = pozicija + 1
     getOsaY = pozicija
End Function




Public Sub inicBojeKugla()
    'inicijaliziranje boja kugla
    '''''''''''''''''''''''''''''''''''''''
    BOJA(1) = &HC0&      'crvena
    BOJA(2) = &HC0C0&    'zuta
    BOJA(3) = &HC000&    'zelena
    BOJA(4) = &HC00000   'plava
    BOJA(5) = &H800080   'pink
    
    JOCKER_BOJA(0) = &H404040      'boja za JOCKER kuglu (NE aktivnu)
    JOCKER_BOJA(1) = &H80FF&       'boja za JOCKER kuglu (aktivnu)
    
    'brojac kugla iste boje
    'PRAVILO:   min. broj kugla iste boje je "istih_kugla_min",
    '           max. broj je "istih_kugla_max"
    '''''''''''''''''''''''''''''''''''''''
    
    Dim a, b, c, d As Integer
    a = Int(getDuzinaTable / 6.7)
    b = Int(getDuzinaTable / 5)
    c = Int(b + 1)
    d = Int(getDuzinaTable / 3.3)
    Dim istih_kugla_min As Integer: istih_kugla_min = getSlucajniBroj(a, b)
    Dim istih_kugla_max As Integer: istih_kugla_max = getSlucajniBroj(c, d)
    
    'pomocna varijabla za odredjivanje maximuma countera
    'odredjuje koliko kugla ima istu boju
    Dim boja_counter_max(BOJA_MAX) As Integer
    Dim boja_counter_min(BOJA_MAX) As Integer
    

    Dim i, j As Integer
    
    'resetiranje countera boja
    For i = 2 To BOJA_MAX
        BOJA_COUNTER(i) = 0
    Next i
    BOJA_COUNTER(1) = getSlucajniBroj(istih_kugla_min, istih_kugla_max)
    
    
    'pomocna varijabla sluzi za zbrajanje sracunatih countera
    Dim ostalo As Integer
    Dim maksimum As Integer
    Dim minimum As Integer
    
    'odredjivanje ranga minimuma da bi odgovarao zbroj kugli
    Dim min_max_prosjek As Integer
    min_max_prosjek = Fix(istih_kugla_max / 2) + Fix(istih_kugla_min / 2)
    
    For i = 2 To BOJA_MAX
        
        ostalo = 0
       
        For j = i To 1 Step -1
            ostalo = ostalo + BOJA_COUNTER(j)
        Next j
        
        boja_counter_max(i) = (getDuzinaTable - ((BOJA_MAX - i) * istih_kugla_min) - ostalo)
        'odredjivanje ranga maksimuma
        If boja_counter_max(i) > istih_kugla_max Then
            maksimum = istih_kugla_max
        Else
            maksimum = boja_counter_max(i)
        End If
        
        boja_counter_min(i) = Abs((getDuzinaTable - ((BOJA_MAX - i) * min_max_prosjek) - ostalo))
        'odredjivanje ranga minimuma
        If boja_counter_min(i) < istih_kugla_min Then
            minimum = istih_kugla_min
        Else
            minimum = boja_counter_min(i)
        End If
        
        'setovanje maksimuma
        BOJA_COUNTER(i) = getSlucajniBroj(minimum, maksimum)
    Next i
End Sub

Public Function getSlucajniBroj(ByVal min As Integer, ByVal max As Integer) As Integer
    '
    'vraca slucajni broj izmedju min i max
    '
    'PROVJERA: AKO JE MIN VECE OD MAX ONDA IH ZAMJENI
    If min > max Then
        Dim tmp As Integer: tmp = min
        min = max
        max = tmp
    End If
    
    Randomize
    getSlucajniBroj = Int((max - min + 1) * Rnd + min) 'formula ;-)
End Function

Public Sub swapKugla(ByVal kugla1 As Integer, ByVal kugla2 As Integer)
    '
    'mjenja pozicije dvije kugle
    'POTREBNO PONOVNO ISCRTAT TABLU!!!
    '
    Dim tmp_kugla As Long: tmp_kugla = 0
    
    tmp_kugla = TABLA(getOsaX(kugla1), getOsaY(kugla1))
    TABLA(getOsaX(kugla1), getOsaY(kugla1)) = TABLA(getOsaX(kugla2), getOsaY(kugla2))
    TABLA(getOsaX(kugla2), getOsaY(kugla2)) = tmp_kugla
End Sub

Public Sub inicNizKugla()
    '
    'Inicijalizacija niza kugla
    'niz sadrzi lokacije kugla koje su povezane (iste boje)
    '
    Dim i As Integer
    For i = 1 To getDuzinaTable
        NIZ_KUGLA(i) = 0
    Next i
End Sub

Public Sub putNizKugla(ByVal kugla As Integer)
    '
    'ubacuje broj kugle na zadnje mjesto
    '
    Dim i As Integer
    For i = 1 To getNizKuglaDuzina
        If getNizKugla(i) = kugla Then Exit Sub
    Next i
    NIZ_KUGLA(getNizKuglaDuzina + 1) = kugla
End Sub

Public Function getDuzinaTable() As Integer
    '
    'vraca broj polja na tabli
    '
    getDuzinaTable = (TABLA_DUZINA_X) * (TABLA_DUZINA_Y)
End Function

Public Function getNizKuglaDuzina() As Integer
    '
    'vraca duzinu niza kugla
    '
    Dim i As Integer
    Dim brojac As Integer: brojac = 0
    For i = 1 To getDuzinaTable
        If NIZ_KUGLA(i) <> 0 Then brojac = brojac + 1
    Next i
    
    getNizKuglaDuzina = brojac
End Function

Public Function getNizKugla(ByVal Index As Integer) As Integer
    '
    'vraca vrijednost indexa iz niza kugla
    '
    
    getNizKugla = NIZ_KUGLA(Index)
End Function

Private Sub cmdNew_Click()
    newGame GAME_MODE_SELECTED_NEW, ORIENTATION
End Sub

Private Sub cmdUndo_Click()
    '
    'VRACANJE PRETHODNOG STANJA TABLE
    '
    Dim i, j As Integer
    
    cmdUndo.Enabled = False
    setKuglaLostFocus
    
    'tabla
    For i = 1 To TABLA_DUZINA_X
        For j = 1 To TABLA_DUZINA_Y
            TABLA(i, j) = TABLA_COPY(i, j)
        Next j
    Next i
    
    'ako je game_mode = 2 - Contiouns
    If (GAME_MODE_SELECTED = 2 Or GAME_MODE_SELECTED = 4) Then
        'vracanje bonusne table
        'SAMO AKO JE PRIJE TOGA OBORENA KOLONA
        If KOLONA_OBORENA = True Then
            TABLA_BONUS_TRENUTNA = (TABLA_BONUS_TRENUTNA) - (TABLA_BONUS_TRENUTNA_LAST)
            showTablaBonus
        End If
    End If
    
    'set jockera
    If checkFirstMove = False Then JOCKER_IN_USE = False
    
    'setovanje bodova
    setBodovi (-BODOVI_TEMP_OSVOJENI)
    
    
    'prikaz table
    showTabla
End Sub



Private Sub Command1_Click()
    On Error GoTo kraj
    Dim X As Byte: X = InputBox("x")
    Dim Y As Byte: Y = InputBox("y")
    Dim b As Byte: b = InputBox("boja" + vbCrLf + "1. crvena" + vbCrLf + "2. zuta" + vbCrLf + "3. zelena" + vbCrLf + "4. plava" + vbCrLf + "5. ljubicasta")
    TABLA(X, Y) = BOJA(b)
    showTabla
kraj:
End Sub
Public Function getNizKuglaMin() As Integer
    '
    'vraca najmanji index kugle iz NizaKugla
    '
    Dim i As Integer
    Dim min As Integer: min = getDuzinaTable
    For i = 1 To getNizKuglaDuzina
        If getNizKugla(i) < min Then
            min = getNizKugla(i)
        End If
    Next i
    getNizKuglaMin = min
End Function

Public Sub inicBodovi()
    '
    'inicijaliziranje bodova
    'formula za bodove (n * (n - 1))
    '
    BODOVI = 0
    setBodovi (0)
    txtBodovi.Visible = True
    hideBodoviTemp
End Sub




Private Sub Form_Click()
    
    setKuglaLostFocus

End Sub




Private Sub Form_Initialize()
    
    'ako je aplikacija vec pokrenuta
    If App.PrevInstance = True Then
        MsgBox "vec pokrenut!!", vbExclamation, "TITLE"
        End
    End If
    
    PREKID_APP = 1

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 13
    
    inicBojeForme
    inicKontrole
    
    GAME_MODE_SELECTED_NEW = 1
    mnuGameModeSelect_Click (GAME_MODE_SELECTED_NEW)
    
    ORIENTATION = 3
    mnuOrientationSelect_Click (ORIENTATION)
    
    GRID = Not (True)
    mnuOptionsGrid_Click
    
    JOCKER_NEW = True
    
    newGame GAME_MODE_SELECTED_NEW, ORIENTATION
    
    Me.Visible = True
    Screen.MousePointer = 0
End Sub



Private Sub Form_Terminate()
    End
End Sub



Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub lblBodoviTemp_Click()
    
    If KLIK = 2 Then Exit Sub
    
    setKuglaLostFocus

End Sub

Private Sub lblBodoviTemp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    KLIK = Button
End Sub


Private Sub lblKugla_Click(Index As Integer)
    
    If SKIP = True Or (KLIK = 2) Then Exit Sub
    
    'ako je jocker ukljucen
    'mjenja jocker kuglu sa trenutnom kliknutom kuglom
    Dim j_ind As Integer: j_ind = getJockerIndex(False)
    If (JOCKER = True) Then
        TABLA(getOsaX(j_ind), getOsaY(j_ind)) = TABLA(getOsaX(Index), getOsaY(Index))
    End If
    
    
    setNizKugla (Index)
    showNizKugla
    showBodoviTemp (Index)
    
    'ako je jocker ukljucen
    'mjenja vraca kuglu u njezinu orginalnu boju
    If (JOCKER = True) Then
        TABLA(getOsaX(j_ind), getOsaY(j_ind)) = JOCKER_BOJA(1)
    End If
    
End Sub



Public Sub setNizKugla(ByVal kugla As Integer)
    '
    'upisue u niz sve povezane kugle iste boje
    '
    inicNizKugla
    putNizKugla (kugla)
            
    Dim brojac As Integer: brojac = 1
    Dim X As Long
    
    Do
        'gleda iznad i ako ima ista upisje u NizKugla
        ''''''''''''''''''''''''''''''''''''''''''''''''
        X = getKuglaGore(getNizKugla(brojac))
        If (X > 0) Then
            If (X = TABLA(getOsaX(getNizKugla(brojac)), getOsaY(getNizKugla(brojac)))) Then
                putNizKugla (getKuglaGoreIndex(getNizKugla(brojac)))
            End If
        End If
        
        'gleda ispod i ako ima ista upisje u NizKugla
        ''''''''''''''''''''''''''''''''''''''''''''''''
        X = getKuglaDole(getNizKugla(brojac))
        If (X > 0) Then
            If (X = TABLA(getOsaX(getNizKugla(brojac)), getOsaY(getNizKugla(brojac)))) Then
                putNizKugla (getKuglaDoleIndex(getNizKugla(brojac)))
            End If
        End If
        
        'gleda ljevo i ako ima ista upisje u NizKugla
        ''''''''''''''''''''''''''''''''''''''''''''''''
        X = getKuglaLjevo(getNizKugla(brojac))
        If (X > 0) Then
            If (X = TABLA(getOsaX(getNizKugla(brojac)), getOsaY(getNizKugla(brojac)))) Then
                putNizKugla (getKuglaLjevoIndex(getNizKugla(brojac)))
            End If
        End If
        
        'gleda desno i ako ima ista upisje u NizKugla
        ''''''''''''''''''''''''''''''''''''''''''''''''
        X = getKuglaDesno(getNizKugla(brojac))
        If (X > 0) Then
            If (X = TABLA(getOsaX(getNizKugla(brojac)), getOsaY(getNizKugla(brojac)))) Then
                putNizKugla (getKuglaDesnoIndex(getNizKugla(brojac)))
            End If
        End If
    
        brojac = brojac + 1
    Loop Until brojac > getNizKuglaDuzina
    
End Sub


Public Sub showNizKugla()
    '
    'prikaz niza kugla na tabli
    '
    '
    
    'resetiranje pozadine kugla
    clearKugle
    If getNizKuglaDuzina <= 1 Then Exit Sub
    
    'prikaz pozadine kugla
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'petlja j samo radi MISTERIOZNOG buga(2x se pomovi isti kod i bude sve ok)..LOL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i, j As Integer
    For j = 1 To 2
        For i = 1 To getNizKuglaDuzina
            lblKugla(getNizKugla(i)).BackColor = BOJA_SELECTED
            lblKugla(getNizKugla(i)).BackStyle = 1
        Next i
    Next j
End Sub

Public Sub inicBojeForme()
    'inicijalizacija boja forme
    Dim i As Integer
    For i = 1 To getDuzinaTable
        lblKugla(i).BackColor = vbBlack
    Next i
End Sub

Private Sub lblKugla_DblClick(Index As Integer)
    '
    'UNISTAVANJE KUGLA I SCROLL ZA DOLE
    '
    
    
    'ako je samo jedna kugla u nizu
    KOLONA_OBORENA = False
    
    If (getNizKuglaDuzina <= 1) Or (KLIK = 2) Then Exit Sub
    
    Screen.MousePointer = 13
        
    Dim i, j, n As Integer
    
    'INICIJALIZACIJA TIPKE UNDO
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To TABLA_DUZINA_X
        For j = 1 To TABLA_DUZINA_Y
            TABLA_COPY(i, j) = TABLA(i, j)
        Next j
    Next i
    cmdUndo.Enabled = True
    
    'postavljanje niza kugla na vrijednost 0
    For i = 1 To getNizKuglaDuzina
        TABLA(getOsaX(getNizKugla(i)), getOsaY(getNizKugla(i))) = 0
    Next i
    
    'SCROLL DOWN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = (TABLA_DUZINA_X + 1) To getDuzinaTable
        'ako naidje na prazno polje treba da skrola citavu kolonu za dole
        If TABLA(getOsaX(i), getOsaY(i)) = 0 Then
            scrollDownKugle (i)
        End If
    Next i
    


    'PROVJERA DA LI IMA PRAZNIH KOLONA (y)
    'ako ima pomjera ih s ljeva na desno
    'ako je GAME_MODE_SELECTED = 2 ili 4 onda ubacuje kolonu y u Empty kolonu
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    Dim br As Integer: br = 0
    'n - ako ima vise od jedne kolone da se treba pomaknuti
    For n = (TABLA_DUZINA_X) - 1 To 1 Step -1
        For i = (TABLA_DUZINA_X) To 1 Step -1
            'samo ako je kolona empty - spaja ih
            If TABLA(i, TABLA_DUZINA_Y) = 0 Then
                moveKoloneDesno (getOsaX(i))
                If (GAME_MODE_SELECTED = 2 Or GAME_MODE_SELECTED = 4) Then
                    'ako je GAME_MODE_SELECTED = 2 onda ubacuje kolonu y u Empty kolonu
                    KOLONA_OBORENA = True
                    'ako je prva kolona empty da ubaci bonus kolonu u 1. kolonu
                    If TABLA(1, TABLA_DUZINA_Y) = 0 Then
                        'ubacivanje bonusa u 1. kolonu
                        For j = TABLA_DUZINA_Y To ((TABLA_DUZINA_Y - TABLA_BONUS_DUZINA(TABLA_BONUS_TRENUTNA)) + 1) Step -1
                            TABLA(1, j) = TABLA_BONUS(TABLA_BONUS_TRENUTNA, j)
                        Next j
                        br = br + 1
                        TABLA_BONUS_TRENUTNA = TABLA_BONUS_TRENUTNA + 1
                        'ako dodje do potrebe za prosiriti niz
                        If (TABLA_BONUS_TRENUTNA) > UBound(TABLA_BONUS, 1) Then
                            setTablaBonus (TABLA_BONUS_TRENUTNA)
                        End If
                        showTablaBonus
                    End If
                End If
            End If
        Next i
    Next n
    
    TABLA_BONUS_TRENUTNA_LAST = br
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (GAME_MODE_SELECTED = 3 Or GAME_MODE_SELECTED = 4) Then
        'ako je GAME_MODE_SELECTED = 3 (Shift) onda pomjera kugle u desno
        'NE MOZE IZMEDJU KOLONA BITI EMPTY KUGLA npr.:
        '
        '1  5   1                                       0   1   1
        '2  5   2   =>(ako se obore dvije petice) =>    0   2   2
        '3  1   4                                       3   1   4
        
        Dim iteration As Integer
        Dim iteration_i_start As Integer
        Dim iteration_j_start As Integer
        
        If KOLONA_OBORENA = True Then
            iteration = TABLA_DUZINA_X - 1
            iteration_i_start = TABLA_DUZINA_Y
            iteration_j_start = TABLA_DUZINA_X
        Else
            iteration = (getNizKuglaMaxX - getNizKuglaMinX) + 1
            iteration_i_start = getNizKuglaMaxY
            iteration_j_start = getNizKuglaMaxX
        End If
        
        
        For i = iteration_i_start To 1 Step -1
            For n = 1 To iteration
                For j = iteration_j_start To 2 Step -1
                    'shiftovanje kugla u desno ako na poziciji nema kugle
                    If TABLA(j, i) = 0 Then shiftKugleDesno j, i
                Next j
            Next n
        Next i
        
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'prikaz table i bodova
    hideBodoviTemp
    BODOVI_TEMP_OSVOJENI = BODOVI_TEMP
    setBodovi (BODOVI_TEMP_OSVOJENI)
    showTabla
    
    
    'PROVJERA DA LI JE GAME OVER (tj. da li ima vishe istih kugla skupa)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim kraj As Byte: kraj = 1
    
    For i = 1 To getDuzinaTable
        'alocira niz spojenih kugla
        setNizKugla (i)
        
        'ako nema bar 2 kugle u nizu onda je GAME OVER
        If getNizKuglaDuzina >= 2 Then
            kraj = 0
            Exit For
        End If
    Next i
   
    If kraj = 1 Then showGameOver
   
    Screen.MousePointer = 0
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub lblKugla_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    KLIK = Button
    
    'najbolja putanja
    '   1)  GORNJA
    '   2)  DONJA
    '   3)  LJEVA
    '   4)  DESNA
    Dim the_best As Byte
    
    'duzina najbolje putanje
    Dim the_best_duzina As Integer

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ako je kliknuta jocker kugla i jocker kugla je aktivirana
    If (Button = 1) And Index = getJockerIndex(False) Then
        
        the_best = 0
        the_best_duzina = 0
        
        'provjera za gornjeg        => 1
        If getKuglaGoreIndex(Index) > 0 Then
            TABLA(getOsaX(Index), getOsaY(Index)) = TABLA(getOsaX(getKuglaGoreIndex(Index)), getOsaY(getKuglaGoreIndex(Index)))
            setNizKugla (Index)
            If getNizKuglaDuzina > the_best_duzina Then
                the_best_duzina = getNizKuglaDuzina
                the_best = 1
            End If
        End If
            
        'provjera za donjeg         => 2
        If getKuglaDoleIndex(Index) > 0 Then
            TABLA(getOsaX(Index), getOsaY(Index)) = TABLA(getOsaX(getKuglaDoleIndex(Index)), getOsaY(getKuglaDoleIndex(Index)))
            setNizKugla (Index)
            If getNizKuglaDuzina > the_best_duzina Then
                the_best_duzina = getNizKuglaDuzina
                the_best = 2
            End If
        End If
        
        'provjera za ljevog         => 3
        If getKuglaLjevoIndex(Index) > 0 Then
            TABLA(getOsaX(Index), getOsaY(Index)) = TABLA(getOsaX(getKuglaLjevoIndex(Index)), getOsaY(getKuglaLjevoIndex(Index)))
            setNizKugla (Index)
            If getNizKuglaDuzina > the_best_duzina Then
                the_best_duzina = getNizKuglaDuzina
                the_best = 3
            End If
        End If
        
        'provjera za desnog         => 4
        If getKuglaDesnoIndex(Index) > 0 Then
            TABLA(getOsaX(Index), getOsaY(Index)) = TABLA(getOsaX(getKuglaDesnoIndex(Index)), getOsaY(getKuglaDesnoIndex(Index)))
            setNizKugla (Index)
            If getNizKuglaDuzina > the_best_duzina Then
                the_best_duzina = getNizKuglaDuzina
                the_best = 4
            End If
        End If
        
        'prikaz najbolje putanje
        If the_best = 1 Then
            'za putanju GORE
            TABLA(getOsaX(Index), getOsaY(Index)) = TABLA(getOsaX(getKuglaGoreIndex(Index)), getOsaY(getKuglaGoreIndex(Index)))
            setNizKugla (getKuglaGoreIndex(Index))
            showBodoviTemp (getKuglaGoreIndex(Index))
        
        ElseIf the_best = 2 Then
            'za putanju DOLE
            TABLA(getOsaX(Index), getOsaY(Index)) = TABLA(getOsaX(getKuglaDoleIndex(Index)), getOsaY(getKuglaDoleIndex(Index)))
            setNizKugla (getKuglaDoleIndex(Index))
            showBodoviTemp (getKuglaDoleIndex(Index))
        
        ElseIf the_best = 3 Then
           'za putanju LJEVO
           TABLA(getOsaX(Index), getOsaY(Index)) = TABLA(getOsaX(getKuglaLjevoIndex(Index)), getOsaY(getKuglaLjevoIndex(Index)))
           setNizKugla (getKuglaLjevoIndex(Index))
           showBodoviTemp (getKuglaLjevoIndex(Index))
        
        ElseIf the_best = 4 Then
            'za putanju DESNO
            TABLA(getOsaX(Index), getOsaY(Index)) = TABLA(getOsaX(getKuglaDesnoIndex(Index)), getOsaY(getKuglaDesnoIndex(Index)))
            setNizKugla (getKuglaDesnoIndex(Index))
            showBodoviTemp (getKuglaDesnoIndex(Index))
        End If
        
        showNizKugla

        'mjenja - vraca kuglu u njezinu orginalnu boju
        TABLA(getOsaX(Index), getOsaY(Index)) = JOCKER_BOJA(1)
    
        SKIP = True
        Exit Sub
    End If
    
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ako je kliknuta jocker kugla da se moze RightButton aktivirati/deaktivirati
    Dim i As Integer
    
    If (Button = 2) And Index = getJockerIndex(True) Then
        
        'ako su kugle selektirane & ako se u selektiranim kuglama nalazi jocker kugla
        Dim u_nizu As Boolean: u_nizu = False
        Dim TMP_j_ind As Integer
        
        
        If shpKugla(Index).FillColor = JOCKER_BOJA(0) Then
            
            
            shpKugla(Index).FillColor = JOCKER_BOJA(1)
            TABLA(getOsaX(Index), getOsaY(Index)) = JOCKER_BOJA(1)
            
            If (getSelectedKugle = True) And (getJockerInNiz = True) Then
                Dim j_ind As Integer: j_ind = getJockerIndex(False)

                TABLA(getOsaX(j_ind), getOsaY(j_ind)) = TABLA(getOsaX(getNizKugla(1)), getOsaY(getNizKugla(1)))

                'ako su iste boje kugla iznad i kugle u nizu
                If checkKuglaColor(getKuglaGoreIndex(Index), getNizKugla(1)) = True Then
                    setNizKugla (getKuglaGoreIndex(Index))
                    showBodoviTemp (getKuglaGoreIndex(Index))
                
                'ako su iste boje kugla dole i kugle u nizu
                ElseIf checkKuglaColor(getKuglaDoleIndex(Index), getNizKugla(1)) = True Then
                    setNizKugla (getKuglaDoleIndex(Index))
                    showBodoviTemp (getKuglaDoleIndex(Index))
                
                'ako su iste boje kugla ljevo i kugle u nizu
                ElseIf checkKuglaColor(getKuglaLjevoIndex(Index), getNizKugla(1)) = True Then
                    setNizKugla (getKuglaLjevoIndex(Index))
                    showBodoviTemp (getKuglaLjevoIndex(Index))
                
                'ako su iste boje kugla desno i kugle u nizu
                ElseIf checkKuglaColor(getKuglaDesnoIndex(Index), getNizKugla(1)) = True Then
                    setNizKugla (getKuglaDesnoIndex(Index))
                    showBodoviTemp (getKuglaDesnoIndex(Index))
                End If
                
                showNizKugla
                
                TABLA(getOsaX(j_ind), getOsaY(j_ind)) = JOCKER_BOJA(1)
                SKIP = True
                Exit Sub
            
            End If
        
        Else
            
            shpKugla(Index).FillColor = JOCKER_BOJA(0)
            TABLA(getOsaX(Index), getOsaY(Index)) = JOCKER_BOJA(0)
            
            'AKO SE ISKLJUCI JOCKER KUGLA DOK SU KUGLE SELEKTIRANE
            'A U SELEKTIRANIM KUGLAMA SE NALAZI JOCKER KUGLA
            'TREBA DA SE:
            '   SELEKTIRAJU ONE KUGLE IZ VEC SELEKTIRANOG NIZA KOJIH IMA NAJVISE
            '   (BEZ JOCKER KUGLE)
            
            TMP_j_ind = getJockerIndex(True)
            For i = 1 To getNizKuglaDuzina
                If getNizKugla(i) = TMP_j_ind Then u_nizu = True
            Next i
            
            If (getSelectedKugle = True) And (u_nizu = True) Then
                
                the_best = 0
                the_best_duzina = 0
       
                'ako kugla GORE iste boje kao kugla iz niza
                If checkKuglaColor(getNizKugla(1), getKuglaGoreIndex(Index)) = True Then
                    setNizKugla (getKuglaGoreIndex(Index))
                    If (getNizKuglaDuzina > the_best_duzina) Then
                        the_best_duzina = getNizKuglaDuzina
                        the_best = 1
                    End If
                End If
                
                'ako kugla DOLE iste boje kao kugla iz niza
                If checkKuglaColor(getNizKugla(1), getKuglaDoleIndex(Index)) = True Then
                    setNizKugla (getKuglaDoleIndex(Index))
                    If (getNizKuglaDuzina > the_best_duzina) Then
                        the_best_duzina = getNizKuglaDuzina
                        the_best = 2
                    End If
                End If
                
                'ako kugla LJEVO iste boje kao kugla iz niza
                If checkKuglaColor(getNizKugla(1), getKuglaLjevoIndex(Index)) = True Then
                    setNizKugla (getKuglaLjevoIndex(Index))
                    If (getNizKuglaDuzina > the_best_duzina) Then
                        the_best_duzina = getNizKuglaDuzina
                        the_best = 3
                    End If
                End If
                
                'ako kugla DESNO iste boje kao kugla iz niza
                If checkKuglaColor(getNizKugla(1), getKuglaDesnoIndex(Index)) = True Then
                    setNizKugla (getKuglaDesnoIndex(Index))
                    If (getNizKuglaDuzina > the_best_duzina) Then
                        the_best_duzina = getNizKuglaDuzina
                        the_best = 4
                    End If
                End If
                
                If the_best = 1 Then
                    'za UP putanju
                    setNizKugla (getKuglaGoreIndex(Index))
                    showBodoviTemp (getKuglaGoreIndex(Index))
                
                ElseIf the_best = 2 Then
                    'za DOWN putanju
                    setNizKugla (getKuglaDoleIndex(Index))
                    showBodoviTemp (getKuglaDoleIndex(Index))
                
                ElseIf the_best = 3 Then
                    'za LEFT putanju
                    setNizKugla (getKuglaLjevoIndex(Index))
                    showBodoviTemp (getKuglaLjevoIndex(Index))
                
                ElseIf the_best = 4 Then
                    'za RIGHT putanju
                    setNizKugla (getKuglaDesnoIndex(Index))
                    showBodoviTemp (getKuglaDesnoIndex(Index))
                End If
                
                'prikaz niza kugla
                showNizKugla
            End If
        End If
        
        SKIP = True
        Exit Sub
    End If
    
    
    
    SKIP = False
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ako je stisnut RightButton tipka i jocker je upaljen da prikaze PopUp Menu
    If JOCKER = True Then
        If (Button = 2) And (JOCKER_IN_USE = False) Then
            setKuglaLostFocus
            
            KUGLA_TRENUTNA = Index
    
            PopupMenu mnuJocker
        End If
    End If
    

End Sub

Private Sub lblKugla_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = _
    "X= " & getOsaX(Index) & vbCrLf & _
    "Y= " & getOsaY(Index) & vbCrLf & _
    Index
End Sub



Public Sub setBodovi(ByVal BOD As Integer)
    '
    'setovanje i prikazivanje bodova
    '
    BODOVI = BODOVI + BOD
    txtBodovi.Caption = BODOVI
End Sub

Public Sub showBodoviTemp(ByVal kugla As Integer)
    '
    'procedura za priakaz privremenih mogucih bodova
    '
    
    'ako nema u nizu vishe od jedne kugle
    If getNizKuglaDuzina <= 1 Then
        hideBodoviTemp
        Exit Sub
    End If

    'ako ima vishe od jedne kugle u nizu
    Dim min As Integer: min = getNizKuglaMin
    
    If (((getOsaX(min)) = 1) And ((getOsaY(min)) = 1)) Then
        'ako je u prvoj koloni i u prvom redu
        shpBodoviTemp.Top = shpKugla(min).Top - 150
        shpBodoviTemp.Left = shpKugla(min).Left - 150
        
        lblBodoviTemp.Top = shpBodoviTemp.Top + 100
        lblBodoviTemp.Left = shpBodoviTemp.Left
    
    ElseIf (getOsaX(min)) = 1 Then
        'ako je prva kugla u nizu u prvoj koloni bodovi se pokazuju iznad
        shpBodoviTemp.Top = shpKugla(min).Top - 650
        shpBodoviTemp.Left = shpKugla(min).Left
        
        lblBodoviTemp.Top = shpBodoviTemp.Top + 100
        lblBodoviTemp.Left = shpBodoviTemp.Left
    
    ElseIf (getOsaY(min)) = 1 Then
        'ako je prva kugla u nizu u prvom redu bodovi se pokazuju ljevo
        shpBodoviTemp.Top = shpKugla(min).Top
        shpBodoviTemp.Left = shpKugla(min).Left - 650
        
        lblBodoviTemp.Top = shpBodoviTemp.Top + 100
        lblBodoviTemp.Left = shpBodoviTemp.Left
    
    Else
        'bodovi se pokazuju iznad ljevo
        
        shpBodoviTemp.Top = shpKugla(min).Top - 650
        shpBodoviTemp.Left = shpKugla(min).Left - 600
        
        lblBodoviTemp.Top = shpBodoviTemp.Top + 100
        lblBodoviTemp.Left = shpBodoviTemp.Left
    End If
    
    'prikaz shape-a sa bodovima
    shpBodoviTemp.BackColor = shpKugla(kugla).FillColor
    shpBodoviTemp.Visible = True
    
    'prikaz labele sa bodovima
    lblBodoviTemp.Caption = getBodoviTemp(getNizKuglaDuzina)
    lblBodoviTemp.Visible = True

End Sub

Public Function getBodoviTemp(ByVal n As Integer) As Integer
    '
    'racuna broj bodova u odnosu na broj kugla
    '
    BODOVI_TEMP = (n) * (n - 1)
    getBodoviTemp = BODOVI_TEMP
End Function

Public Sub hideBodoviTemp()
    lblBodoviTemp.Visible = False
    shpBodoviTemp.Visible = False
End Sub

Public Sub scrollDownKugle(ByVal Index As Integer)
    '
    'Scrolla kugla za dole
    '
    '
    'primjer:
    '5          (Empty) = 0
    '6          5
    '21     =>  6
    '22         21
    '(index)    22
    
    Dim i As Integer
    Dim osa_x As Integer: osa_x = getOsaX(Index)
    Dim y_min As Integer: y_min = 1
    
    'odredjiva u kojem Y-nu je nula
    For i = 1 To (getOsaY(Index) - 1)
        If TABLA(osa_x, i) = 0 Then
            y_min = i
            Exit For
        End If
    Next i
    
    
    'scrola kuglu po kuglu odozgo prema dole
    For i = (getOsaY(Index) - 1) To y_min Step -1
        TABLA(osa_x, i + 1) = TABLA(osa_x, i)
        If i = y_min Then TABLA(osa_x, i) = 0
    Next i
End Sub

Public Sub moveKoloneDesno(ByVal kolonaX As Integer)
    '
    'pomjera kolone na desno
    '
    'PRIMJER:
    '1  2   3   kolonaX     4           0   1   2   3   4
    '4  4   1   kolonaX     4       =>  0   4   4   1   4
    '1  2   3   kolonaX     2           0   1   2   3   2
    
    Dim i, j As Integer
    
    For i = kolonaX To 1 Step -1
        For j = 1 To TABLA_DUZINA_Y
            
            'ako prebaciva kolonu 1 (X=1) onda oznacava kolonu 1 (X=1) kao Empty
            If i = 1 Then
                TABLA(i, j) = 0
            Else
                TABLA(i, j) = TABLA(i - 1, j)
            End If
            
        Next j
    Next i
End Sub

Public Sub showGameOver()
    cmdUndo.Enabled = False
    MsgBox "game over", vbInformation
End Sub

Public Sub inicKontrole()
    cmdNew.Enabled = True
    cmdUndo.Enabled = False
    
    lblGameMode.Caption = ""
    lblOrientation.Caption = ""
    txtBodovi.Caption = ""
    
    lblJocker.Caption = ""
    mnuJocker.Visible = False
End Sub

Public Sub clearKugle()
    '
    'BOJI POZADINU KUGLA U DEFAULTNU
    '
    
    Dim i As Integer
    For i = 1 To getDuzinaTable
        'lblKugla(i).BackColor = frmMain.BackColor
        lblKugla(i).BackStyle = 0
    Next i
End Sub

Public Sub inicGameMode()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'INICIJALIZACIJA GAME MODOVA
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim i As Integer
    
    inicGameModeText
    
    If (GAME_MODE_SELECTED = 2) Or (GAME_MODE_SELECTED = 4) Then
        
        'ako je game mode 2 - Continuous
        KOLONA_OBORENA = False
        inicTablaBonus
        showTablaBonus
    ElseIf (GAME_MODE_SELECTED = 1) Or (GAME_MODE_SELECTED = 3) Then
        
        'ako je game mode 1 - Standard ili 3 - Shifter
        For i = 1 To TABLA_DUZINA_Y
            shpKuglaBonus(i).Visible = False
        Next i
    End If
End Sub

Public Sub inicGameModeText()
    '
    'inicijalizacija game mode textova i prikazivanje
    '
    GAME_MODE_TEXT(1) = "Standard"
    GAME_MODE_TEXT(2) = "Continuous"
    GAME_MODE_TEXT(3) = "Shifter"
    GAME_MODE_TEXT(4) = "MegaShift"
    
    lblGameMode.Caption = GAME_MODE_TEXT(GAME_MODE_SELECTED)
End Sub

Public Sub inicTablaBonus()
    '
    'inicijalizacija bonusnih tabli
    '
    Dim i, j As Integer
    
    TABLA_BONUS_TRENUTNA = 1
    TABLA_BONUS_TRENUTNA_LAST = 1
    
    ReDim TABLA_BONUS(TABLA_DUZINA_X, TABLA_DUZINA_Y)
    ReDim TABLA_BONUS_DUZINA(TABLA_DUZINA_X)
    
    'resetiranje bonus table na nulu
    For i = 1 To UBound(TABLA_BONUS, 1)
        For j = 1 To UBound(TABLA_BONUS, 2)
            TABLA_BONUS(i, j) = 0
        Next j
    Next i
    
    'setovaje tabela
    For i = 1 To UBound(TABLA_BONUS, 1)
        setTablaBonus (i)
    Next i
    
End Sub



Public Sub showTablaBonus()
    '
    'prikaz bonus table
    '
    Dim i As Integer
    
    'setovanje shapova da budu nevidljivi
    For i = 1 To TABLA_DUZINA_Y
        shpKuglaBonus(i).Visible = False
    Next i
    
   
    'bojenje shapeova
    For i = 1 To TABLA_BONUS_DUZINA(TABLA_BONUS_TRENUTNA)
        shpKuglaBonus(i).FillColor = TABLA_BONUS(TABLA_BONUS_TRENUTNA, (TABLA_DUZINA_Y + 1) - i)
        shpKuglaBonus(i).Visible = True
    Next i

End Sub


Public Sub setTablaBonus(ByVal kolonaX As Byte)
    'primjer:
    '
    '0  3   0   ... X->
    '0  1   0       .
    '0  4   0       .
    '3  6   0       .
    '2  8   0       .
    '3  1   0       . |
    '3  1   1   ... Y ¢
        
    Dim i, j As Integer
    
    'AKO JE DOSLO DO PROSIRENJA BONUS NIZA
        '1) kopirati trenutni niz u copy niz...ok
        '2) prosiriti orginalni niz...ok
        '3) vratiti iz copy u orginalni...ok
        '
        '4) kopirati trenutni niz duzine u copy niz duzine...ok
        '5) prosiriti orginalni niz duzine...ok
        '6) vratiti iz copy u orginalni...ok
    
    If kolonaX > UBound(TABLA_BONUS, 1) Then
        
        '1)
            ReDim TABLA_BONUS_COPY(UBound(TABLA_BONUS, 1), UBound(TABLA_BONUS, 2))
        
            For i = 1 To (UBound(TABLA_BONUS, 1))
                For j = 1 To (UBound(TABLA_BONUS, 2))
                    TABLA_BONUS_COPY(i, j) = TABLA_BONUS(i, j)
                Next j
            Next i
        '2)
            ReDim TABLA_BONUS(UBound(TABLA_BONUS_COPY, 1) + 1, UBound(TABLA_BONUS_COPY, 2))
        '3)
            For i = 1 To (UBound(TABLA_BONUS, 1) - 1)
                For j = 1 To (UBound(TABLA_BONUS, 2))
                    TABLA_BONUS(i, j) = TABLA_BONUS_COPY(i, j)
                Next j
            Next i
        '4)
            ReDim TABLA_BONUS_DUZINA_COPY(UBound(TABLA_BONUS_DUZINA, 1))
            For i = 1 To (UBound(TABLA_BONUS_DUZINA, 1))
                TABLA_BONUS_DUZINA_COPY(i) = TABLA_BONUS_DUZINA(i)
            Next i
        '5)
            ReDim TABLA_BONUS_DUZINA(UBound(TABLA_BONUS_DUZINA_COPY, 1) + 1)
        '6)
            For i = 1 To (UBound(TABLA_BONUS_DUZINA, 1) - 1)
                TABLA_BONUS_DUZINA(i) = TABLA_BONUS_DUZINA_COPY(i)
            Next i
    End If
    
    '
    'setuje random duzinu bonus table
    '
    Dim min As Byte: min = getSlucajniBroj(1, Fix(TABLA_DUZINA_Y / 2))
    Dim max As Byte: max = getSlucajniBroj(1, TABLA_DUZINA_Y)
    
    TABLA_BONUS_DUZINA(kolonaX) = getSlucajniBroj(min, max)
    
    'inicijalizacija boja i duzina bonus tabela
    'vishe bonus tabela se koristi da se u slucaju vise umetnutih tabla moze vratiti UNDO
    For j = UBound(TABLA_BONUS, 2) To ((UBound(TABLA_BONUS, 2) - TABLA_BONUS_DUZINA(kolonaX)) + 1) Step -1
        TABLA_BONUS(kolonaX, j) = BOJA(getSlucajniBroj(1, BOJA_MAX))
    Next j
End Sub
Public Function getNizKuglaMinX() As Integer
    '
    'vraca najmanji broj kolone (osaX) u kojoj se nalazi niz kugla
    'Reminder: NIZ KUGLA CUVA INDEX-e SHAPOVA U KOJIMA SU SVE KUGLE U NIZU (SKUPA)
    '
    Dim osa_x As Byte: osa_x = TABLA_DUZINA_X
    Dim i As Integer
    
    For i = 1 To getNizKuglaDuzina
        If getOsaX(getNizKugla(i)) < osa_x Then osa_x = getOsaX(getNizKugla(i))
    Next i
    
    getNizKuglaMinX = osa_x
End Function
Public Function getNizKuglaMaxX() As Integer
    '
    'vraca najveci broj kolone (osaX) u kojoj se nalazi niz kugla
    'Reminder: NIZ KUGLA CUVA INDEX-e SHAPOVA U KOJIMA SU SVE KUGLE U NIZU (SKUPA)
    '
    Dim osa_x As Integer: osa_x = 1
    Dim i As Integer
    
    For i = 1 To getNizKuglaDuzina
        If getOsaX(getNizKugla(i)) > osa_x Then osa_x = getOsaX(getNizKugla(i))
    Next i
    
    getNizKuglaMaxX = osa_x
End Function

Public Function getNizKuglaMaxY() As Integer
    '
    'vraca najveci broj reda (osaY) u kojoj se nalazi niz kugla
    'Reminder: NIZ KUGLA CUVA INDEX-e SHAPOVA U KOJIMA SU SVE KUGLE U NIZU (SKUPA)
    '
    Dim osa_y As Integer: osa_y = 1
    Dim i As Integer
    
    For i = 1 To getNizKuglaDuzina
        If getOsaY(getNizKugla(i)) > osa_y Then osa_y = getOsaY(getNizKugla(i))
    Next i
    
    getNizKuglaMaxY = osa_y
End Function




Public Sub shiftKugleDesno(ByVal x_osa As Byte, ByVal y_osa As Byte)
    '
    'pomjera kugle u desno do x_osa pozicije (na y_osi)
    'example:
    '
    '
    '   5   1   5   x_osa   =>  0   5   1   5
    '
    'POTREBNO PONOVNO ISCRTAT TABLU!!!
    
    Dim i As Integer
    
    For i = x_osa To 2 Step -1
        TABLA(i, y_osa) = TABLA(i - 1, y_osa)
    Next i
    
    'oznacava pocetne kugle kao prazne
    If getKuglaMinOsaY(y_osa) = 0 Then
        TABLA(getKuglaMinOsaY(y_osa) + 1, y_osa) = 0
    Else
        TABLA(getKuglaMinOsaY(y_osa), y_osa) = 0
    End If
End Sub

Public Function getKuglaMinOsaY(ByVal osa_y As Byte) As Integer
    '
    'vraca do kojeg reda (osaX) su prazne kugle u odredjenoj koloni (osaY)
    '
    Dim i As Integer
    Dim brojac As Integer: brojac = 0
    
    For i = 1 To TABLA_DUZINA_X
        If TABLA(i, osa_y) <> 0 Then
            brojac = (i - 1)
            Exit For
        End If
    Next i
    
    getKuglaMinOsaY = brojac
End Function

Public Sub newGame(ByVal mod_igre As Byte, ByVal orijentacija_igre As Byte)
    '
    'starta novu igru
    '
    
    On Error Resume Next
    
    Screen.MousePointer = 13
    
    GAME_MODE_SELECTED = mod_igre
    ORIENTATION_SELECTED = orijentacija_igre
    
    cmdNew.Enabled = False
    cmdUndo.Enabled = False
    
    inicOrientation
    inicDynVar
   
    inicBodovi
    inicBojeKugla
    inicGameMode
    inicTabla
    inicLokacijaKugla
    inicJocker
    inicGrid
    
    inicNizKugla 'treba da vrati niz kugla na nulu
    
    showTabla
    showBojeCounter
    showGrid (GRID)
    
    
    cmdNew.Enabled = True
    cmdNew.SetFocus
    
    
    Screen.MousePointer = 0

End Sub







Public Sub inicLokacijaKugla()
    '
    'inicijalizacija (pozicija) kugla na ekranu
    '
    
    Dim i, j As Integer
    
    'setovanje nevidljivosti svih shapeova i labela
    For i = 1 To getDuzinaTable
        shpKugla(i).Visible = False
        lblKugla(i).Visible = False
    Next i
    
    Const kugla_top As Integer = 360
    Const kugla_left As Integer = 1560
    Const kugla_height As Integer = 715 '855
    Const kugla_width As Integer = 715 '735
    
    'inicijalizacija pocetne (prve) kugle (ALL)
    shpKugla(1).Top = kugla_top
    shpKugla(1).Left = kugla_left
    shpKugla(1).Height = kugla_height
    shpKugla(1).Width = kugla_width

    lblKugla(1).Top = kugla_top
    lblKugla(1).Left = kugla_left
    lblKugla(1).Height = kugla_height
    lblKugla(1).Width = kugla_width
    
    'inicijalizacija svih ostalih kugla (Width & Height)
    For i = 2 To getDuzinaTable
        shpKugla(i).Height = kugla_height
        shpKugla(i).Width = kugla_width
        
        lblKugla(i).Height = kugla_height
        lblKugla(i).Width = kugla_width
    Next i
    
    'inicijalizacija kugla u prvom redu (Top & Left)
    For i = 2 To TABLA_DUZINA_X
        shpKugla(i).Left = shpKugla(i - 1).Left + shpKugla(i - 1).Width
        shpKugla(i).Top = kugla_top
    
        lblKugla(i).Left = lblKugla(i - 1).Left + lblKugla(i - 1).Width
        lblKugla(i).Top = kugla_top
    Next i
    
    'inicijalizacija kugla u prvoj koloni (Top & Left)
    For i = 2 To TABLA_DUZINA_Y
        shpKugla(getKuglaIndex(1, i)).Left = kugla_left
        shpKugla(getKuglaIndex(1, i)).Top = (shpKugla(getKuglaGoreIndex(getKuglaIndex(1, i))).Top) + (shpKugla(getKuglaGoreIndex(getKuglaIndex(1, i))).Height)
    
        lblKugla(getKuglaIndex(1, i)).Left = kugla_left
        lblKugla(getKuglaIndex(1, i)).Top = (lblKugla(getKuglaGoreIndex(getKuglaIndex(1, i))).Top) + (lblKugla(getKuglaGoreIndex(getKuglaIndex(1, i))).Height)
    Next i
    
    'inicijalizacija kugla u ostalim redovima i kolonama (Top & Left)
    For i = 2 To TABLA_DUZINA_X
        For j = 2 To TABLA_DUZINA_Y
            shpKugla(getKuglaIndex(i, j)).Left = (shpKugla(getKuglaIndex(i - 1, j)).Left) + (shpKugla(getKuglaIndex(i - 1, j)).Width)
            shpKugla(getKuglaIndex(i, j)).Top = (shpKugla(getKuglaGoreIndex(getKuglaIndex(i, j))).Top) + (shpKugla(getKuglaGoreIndex(getKuglaIndex(i, j))).Height)
            
            lblKugla(getKuglaIndex(i, j)).Left = (lblKugla(getKuglaIndex(i - 1, j)).Left) + (lblKugla(getKuglaIndex(i - 1, j)).Width)
            lblKugla(getKuglaIndex(i, j)).Top = (lblKugla(getKuglaGoreIndex(getKuglaIndex(i, j))).Top) + (lblKugla(getKuglaGoreIndex(getKuglaIndex(i, j))).Height)
        Next j
    Next i
    
    'vidljivost kugla
    For i = 1 To getDuzinaTable
        lblKugla(i).Visible = True
        shpKugla(i).Visible = True
    Next i
    
    '''''''''''''''''''''''''''''''''''''''
    'lokacija i dimenzije bonus kugla
    
    'postavlja Top 1. kugle u pravcu zadnje kugle u Y osi
    Dim kugla_bonus_top As Integer
    kugla_bonus_top = shpKugla(getKuglaIndex(1, TABLA_DUZINA_Y)).Top - shpKugla(getKuglaIndex(1, TABLA_DUZINA_Y)).Height
    
    'postavlja Left 1. kugle u pravcu 1. kugle i odmaknuto za sirinu kugle
    Dim kugla_bonus_left As Integer: kugla_bonus_left = shpKugla(1).Left - shpKugla(1).Width
    
    Const kugla_bonus_height As Integer = 615
    Const kugla_bonus_width As Integer = 495
    
     
    'inicijalizacija svih bonus kugla (ALL)
    For i = 1 To TABLA_DUZINA_Y
        
        If (i = 1) Then
            shpKuglaBonus(i).Top = kugla_bonus_top
        ElseIf (i >= 2) Then
            shpKuglaBonus(i).Top = shpKuglaBonus(i - 1).Top - shpKuglaBonus(i - 1).Height
        End If
        
        shpKuglaBonus(i).Left = kugla_bonus_left
        shpKuglaBonus(i).Width = kugla_bonus_width
        shpKuglaBonus(i).Height = kugla_bonus_height
    Next i
End Sub



Private Sub mnuGameModeSelect_Click(Index As Integer)
    Dim i As Byte
    For i = 1 To 4
        mnuGameModeSelect(i).Checked = False
    Next i
    
    mnuGameModeSelect(Index).Checked = True
    
    GAME_MODE_SELECTED_NEW = Index
End Sub

Private Sub mnuInfoBall_Click()
    '
    'setovanje Info Ball-a
    '
    'DEFAULT: ON
    
    If mnuInfoBall.Checked = True Then
        mnuInfoBall.Checked = False
        frmInfoBall.Visible = False
    Else
        mnuInfoBall.Checked = True
        frmInfoBall.Visible = True
    End If

End Sub

Private Sub mnuJockerBall_Click()
    '
    'potvrda jocker kugle
    '
    
    If checkFirstMove = False Then
        MsgBox "MORA SE ODIGRATI BAR JEDAN POTEZ!!", vbInformation
        Exit Sub
    End If
    
    
    Dim pitanje
    shpKugla(KUGLA_TRENUTNA).FillColor = JOCKER_BOJA(1)
    
    pitanje = MsgBox("SIGURNO??", vbQuestion + vbYesNo, "Jocker")
    If pitanje = vbYes Then
        'yes
        TABLA(getOsaX(KUGLA_TRENUTNA), getOsaY(KUGLA_TRENUTNA)) = JOCKER_BOJA(1)
        showTabla
        JOCKER_IN_USE = True
    Else
        'no
        shpKugla(KUGLA_TRENUTNA).FillColor = TABLA(getOsaX(KUGLA_TRENUTNA), getOsaY(KUGLA_TRENUTNA))
    End If
End Sub

Private Sub mnuOptionsGrid_Click()
    '
    'setovanje GRID-a
    '
    If GRID = True Then
        mnuOptionsGrid.Checked = False
        GRID = False
    Else
        mnuOptionsGrid.Checked = True
        GRID = True
    End If
    
    showGrid (GRID)
End Sub

Private Sub mnuOptionsJocker_Click()
    '
    'setovanje JOCKER-a
    '
    If JOCKER_NEW = True Then
        mnuOptionsJocker.Checked = False
        JOCKER_NEW = False
    Else
        mnuOptionsJocker.Checked = True
        JOCKER_NEW = True
    End If
    

End Sub


Private Sub mnuOrientationSelect_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To 4
        mnuOrientationSelect(i).Checked = False
    Next i
    
    mnuOrientationSelect(Index).Checked = True
    ORIENTATION = Index
End Sub


Public Sub inicOrientation()
    '
    'Inicijalizacija Orjentacije igre
    '
    If ORIENTATION_SELECTED = 1 Then
        'ako je odabran PORTRAIT
        TABLA_DUZINA_X = 11
        TABLA_DUZINA_Y = 12
        lblOrientation.Caption = "PORTRAIT"
    
    ElseIf ORIENTATION_SELECTED = 2 Then
        'ako je odabran LANDSCAPE
        TABLA_DUZINA_X = 15
        TABLA_DUZINA_Y = 8
        lblOrientation.Caption = "LANDSCAPE"
    
    ElseIf ORIENTATION_SELECTED = 3 Then
        'ako je BOX
        TABLA_DUZINA_X = 14
        TABLA_DUZINA_Y = 14
        lblOrientation.Caption = "BOX"
    ElseIf ORIENTATION_SELECTED = 4 Then
        'ako je BOX
        TABLA_DUZINA_X = 11
        TABLA_DUZINA_Y = 16
        lblOrientation.Caption = "X1"
    
    Else
        'ako je error
        MsgBox "NIJE INICIJALIZIRANA DIMENZIJA TABLE!", vbCritical
        Unload Me
        End
    End If
End Sub


Public Sub inicDynVar()
    
    ReDim TABLA(TABLA_DUZINA_X, TABLA_DUZINA_Y) As Long
    
    ReDim TABLA_COPY(TABLA_DUZINA_X, TABLA_DUZINA_Y) As Long
    
    ReDim NIZ_KUGLA(TABLA_DUZINA_X * TABLA_DUZINA_Y) As Integer
    
    GenerateControlArray shpKugla, (getDuzinaTable)
    GenerateControlArray lblKugla, (getDuzinaTable)
    
    'ako je game mode 2 ili 4 da kreira bonus shapeove
    If (GAME_MODE_SELECTED = 2 Or GAME_MODE_SELECTED = 4) Then
        GenerateControlArray shpKuglaBonus, TABLA_DUZINA_Y
    ElseIf (GAME_MODE_SELECTED = 1 Or GAME_MODE_SELECTED = 3) Then
        GenerateControlArray shpKuglaBonus, 1
    End If
End Sub
Public Sub GenerateControlArray(kontrola As Variant, ByVal count As Integer)
    '
    ' kreira niz kontrola
    '
    
    'ako je vec kreirano "count" kontrola
    If count = (kontrola.ubound) Then Exit Sub
    
    Dim i As Integer

    On Error Resume Next

    'clean up
    For i = 1 To (kontrola.ubound)
        Unload kontrola(i)
    Next i
    
    'create new
    For i = 1 To (count)
        Load kontrola(i)
    Next i
End Sub


Public Function getScreenResolution(ByVal parametar As String) As Integer
    '
    'vraca rezoluciju ekrana (sirinu ili visinu - ovisno od parametra)
    '
    
    Dim sirina As Integer:  sirina = (Screen.Width / Screen.TwipsPerPixelY)
    Dim visina As Integer:  visina = (Screen.Height / Screen.TwipsPerPixelX)
    
    If (UCase(parametar) = "S") Then
        
        'vraca sirinu
        getScreenResolution = sirina
    
    ElseIf (UCase(parametar) = "V") Then
        
        'vraca visinu
        getScreenResolution = visina
    
    ElseIf (UCase(parametar) = "SIV") Then
        
        'vraca sirina + visina
        getScreenResolution = (sirina + visina)
    
    Else
        
        'vraca NULU
        getScreenResolution = 0
    End If
End Function


Public Sub showGrid(ByVal prikaz As Boolean)
    '
    'iscrtavanje linija (Grid-a) na tabli
    '
    
    On Error Resume Next
    
    Dim i As Byte
    
    'vidljivost grida
    For i = 1 To (TABLA_DUZINA_X + TABLA_DUZINA_Y + 2)
        
        If prikaz = True Then
            lineGrid(i).Visible = True
        
        ElseIf prikaz = False Then
            lineGrid(i).Visible = False
        End If
    
    Next i
End Sub



Private Sub tmrAutoPlay_Timer()
    newGame getSlucajniBroj(1, 4), getSlucajniBroj(1, 3)
End Sub

Public Sub setJocker(ByVal kugla As Integer)
    '
    'setuje Jocker Kuglu
    '
    If JOCKER = False Then
        TABLA(getOsaX(kugla), getOsaY(kugla)) = JOCKER_BOJA(1)
        JOCKER = True
    End If
End Sub

Public Function getJockerIndex(ByVal boja_0 As Boolean) As Integer
    '
    'vraca Index Jocker Kugle
    '
    'ako je parametar TRUE provjerava jocker (ne aktivnu) kuglu
    
    Dim i, j As Integer
    For i = 1 To TABLA_DUZINA_X
        For j = 1 To TABLA_DUZINA_Y
            
            'ako treba da provjerava NEaktiviranog jockera
            If boja_0 = True Then
                If TABLA(i, j) = JOCKER_BOJA(0) Then
                    getJockerIndex = getKuglaIndex(i, j)
                    Exit Function
                End If
            End If
            
            If TABLA(i, j) = JOCKER_BOJA(1) Then
                getJockerIndex = getKuglaIndex(i, j)
                Exit Function
            End If
        Next j
    Next i
    getJockerIndex = 0
End Function

Public Sub inicJocker()
    'inicijalizacija jocker kugle
    JOCKER_IN_USE = False
    JOCKER = JOCKER_NEW
    
    If JOCKER = True Then
        mnuOptionsJocker.Checked = True
        lblJocker.Caption = "JOCKER: ON"
        lblJocker.ForeColor = vbBlue
        lblJockerHint.Visible = True
    Else
        mnuOptionsJocker.Checked = False
        lblJocker.Caption = "JOCKER: OFF"
        lblJocker.ForeColor = vbRed
        lblJockerHint.Visible = False
    End If
    lblJocker.Visible = True
End Sub

Public Sub inicGrid()
    '
    'inicijalizacija GRID-a
    '
    On Error Resume Next

    'generiranje kontrola grid-a
    GenerateControlArray lineGrid, (TABLA_DUZINA_X + TABLA_DUZINA_Y + 2)
    
    
    '''''''''''''''''''''''''''''''''''''''''
    'isrcravanje bocnih linija
    '''''''''''''''''''''''''''''''''''''''''
    
    Dim i As Byte
    
    'debljina bocnih linija
    For i = 1 To 4
        lineGrid(i).BorderWidth = 2
    Next i
    
    'iscrtavanje gornje linije
    lineGrid(1).X1 = (shpKugla(1).Left)
    lineGrid(1).Y1 = (shpKugla(1).Top)
    lineGrid(1).X2 = (shpKugla(getKuglaIndex(TABLA_DUZINA_X, 1)).Left + shpKugla(1).Width)
    lineGrid(1).Y2 = (shpKugla(1).Top)
    
    
    'iscrtavanje donje linije
    lineGrid(2).X1 = (shpKugla(1).Left)
    lineGrid(2).Y1 = (shpKugla(getKuglaIndex(1, TABLA_DUZINA_Y)).Top + shpKugla(1).Height)
    lineGrid(2).X2 = (shpKugla(getKuglaIndex(TABLA_DUZINA_X, TABLA_DUZINA_Y)).Left + shpKugla(1).Width)
    lineGrid(2).Y2 = (shpKugla(getKuglaIndex(TABLA_DUZINA_X, TABLA_DUZINA_Y)).Top + shpKugla(1).Height)
    
    'iscrtavanje ljeve linije
    lineGrid(3).X1 = (shpKugla(1).Left)
    lineGrid(3).Y1 = (shpKugla(1).Top)
    lineGrid(3).X2 = (shpKugla(1).Left)
    lineGrid(3).Y2 = (shpKugla(getKuglaIndex(1, TABLA_DUZINA_Y)).Top + shpKugla(1).Height)
    
    'iscrtavanje desne linije
    lineGrid(4).X1 = (shpKugla(getKuglaIndex(TABLA_DUZINA_X, TABLA_DUZINA_Y)).Left + shpKugla(1).Width)
    lineGrid(4).Y1 = (shpKugla(1).Top)
    lineGrid(4).X2 = (shpKugla(getKuglaIndex(TABLA_DUZINA_X, TABLA_DUZINA_Y)).Left + shpKugla(1).Width)
    lineGrid(4).Y2 = (shpKugla(getKuglaIndex(TABLA_DUZINA_X, TABLA_DUZINA_Y)).Top + shpKugla(1).Height)
    
    
    '''''''''''''''''''''''''''''''''''''''''
    'isrcravanje ostalih linija
    '''''''''''''''''''''''''''''''''''''''''

    'debljina linija
    For i = 5 To TABLA_DUZINA_X + TABLA_DUZINA_Y + 2
        lineGrid(i).BorderWidth = 1
    Next i
    '
    'crtanje horizontalnih linija
    For i = 2 To TABLA_DUZINA_Y
        lineGrid(i + 3).X1 = (shpKugla(getKuglaIndex(1, i)).Left)
        lineGrid(i + 3).Y1 = (shpKugla(getKuglaIndex(1, i)).Top)
        lineGrid(i + 3).X2 = (shpKugla(getKuglaIndex(TABLA_DUZINA_X, i)).Left + shpKugla(1).Width)
        lineGrid(i + 3).Y2 = (shpKugla(getKuglaIndex(TABLA_DUZINA_X, i)).Top)
    Next i

    'crtanje vertikalnih linija
    For i = 2 To TABLA_DUZINA_X
        lineGrid(i + 2 + TABLA_DUZINA_Y).X1 = (shpKugla(i).Left)
        lineGrid(i + 2 + TABLA_DUZINA_Y).Y1 = (shpKugla(i).Top)
        lineGrid(i + 2 + TABLA_DUZINA_Y).X2 = (shpKugla(getKuglaIndex(i, TABLA_DUZINA_Y)).Left)
        lineGrid(i + 2 + TABLA_DUZINA_Y).Y2 = (shpKugla(getKuglaIndex(i, TABLA_DUZINA_Y)).Top + shpKugla(getKuglaIndex(i, TABLA_DUZINA_Y)).Height)
    Next i
End Sub

Public Function getSelectedKugle() As Boolean
    '
    'vraca TRUE ako ima vishe od jedne kugle u nizu
    '
    If getNizKuglaDuzina >= 2 Then
        getSelectedKugle = True
    ElseIf getNizKuglaDuzina < 2 Then
        getSelectedKugle = False
    End If
End Function

Public Function checkKuglaColor(ByVal kugla1 As Integer, ByVal kugla2 As Integer) As Boolean
    '
    'poredi da li ja kugla1 == kugla 2
    'vraca TRUE ako su kugle iste boje
    '
            
    If TABLA(getOsaX(kugla1), getOsaY(kugla1)) = TABLA(getOsaX(kugla2), getOsaY(kugla2)) Then
        checkKuglaColor = True
        Exit Function
    End If
    
    checkKuglaColor = False
End Function

Public Function getJockerInNiz() As Boolean
    '
    'ako je jocker kugla vezana za niz kugla vraca TRUE
    '
    
    Dim i As Integer
    Dim j_ind As Integer: j_ind = getJockerIndex(True)
    
    For i = 1 To getNizKuglaDuzina
        
        If getKuglaGoreIndex(getNizKugla(i)) = (j_ind) Then
            getJockerInNiz = True
            Exit Function
        
        ElseIf getKuglaDoleIndex(getNizKugla(i)) = (j_ind) Then
            getJockerInNiz = True
            Exit Function
        
        ElseIf getKuglaLjevoIndex(getNizKugla(i)) = (j_ind) Then
            getJockerInNiz = True
            Exit Function
        
        ElseIf getKuglaDesnoIndex(getNizKugla(i)) = (j_ind) Then
            getJockerInNiz = True
            Exit Function
        End If
    Next i
    
    getJockerInNiz = False
End Function

Public Sub setKuglaLostFocus()
    '
    ' kada kugla izgubi focus
    '
    inicNizKugla
    clearKugle
    hideBodoviTemp
End Sub

Public Function checkFirstMove() As Boolean
    '
    'vraca TRUE ako je prvi potez odigran
    '
    Dim i As Byte
    Dim br As Byte: br = 0
    
    For i = 1 To TABLA_DUZINA_X
        If TABLA(getOsaX(i), 1) <> 0 Then br = br + 1
    Next i
    
    If br = TABLA_DUZINA_X Then
        checkFirstMove = False
        Exit Function
    End If
    
    checkFirstMove = True
End Function

Public Sub showBojeCounter()
    '
    '   prikaziva koliko ima koje kugle (BOJA)
    '
    
    Dim i, t, t_i As Byte: t = 0: t_i = 0
    
    
    For i = 1 To 5
        lblBojaCounter(i).Caption = BOJA_COUNTER(i)
        lblBojaCounter(i).FontBold = False
        lblBojaCounter(i).FontSize = 12
        lblInfBall(i).FontSize = 12
        lblInfBall(i).FontBold = False
      
        
        If BOJA_COUNTER(i) > t Then
            t = BOJA_COUNTER(i)
            t_i = i
        End If
    Next i
        
    lblBojaCounter(t_i).FontBold = True
    lblBojaCounter(t_i).FontSize = 14
    
    lblInfBall(t_i).FontSize = 14
    lblInfBall(t_i).FontBold = True
    
        
End Sub
