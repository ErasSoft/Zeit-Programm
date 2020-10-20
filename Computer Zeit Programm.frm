VERSION 5.00
Begin VB.Form frm_Zeit_Programm 
   BackColor       =   &H0000FF00&
   Caption         =   "Computer Zeit-Programm"
   ClientHeight    =   4845
   ClientLeft      =   11415
   ClientTop       =   4740
   ClientWidth     =   3840
   ForeColor       =   &H00808080&
   Icon            =   "Computer Zeit Programm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   3840
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame6 
      BackColor       =   &H0000FF00&
      Height          =   2535
      Left            =   7320
      TabIndex        =   49
      Top             =   0
      Width           =   2412
      Begin VB.Label lbl_Version 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "v.1.1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   55
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lbl_Programm 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Zeit-Programm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   54
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lbl_copyright_by 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lbl_copyright_Tino 
         BackStyle       =   0  'Transparent
         Caption         =   "Tino Schuldt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl_copyright_Eras 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "alias Eras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1440
         TabIndex        =   51
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl_copyright_Datum 
         BackStyle       =   0  'Transparent
         Caption         =   "25.03.2008"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Image img_Eras_Logo 
         Height          =   1065
         Left            =   840
         Picture         =   "Computer Zeit Programm.frx":030A
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.OptionButton opt_7 
      BackColor       =   &H0000FF00&
      Caption         =   "Auswahl 7"
      Enabled         =   0   'False
      Height          =   252
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":11EF
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   48
      Top             =   4320
      Width           =   1092
   End
   Begin VB.CommandButton cmd_7 
      Caption         =   "Eigene Anwendung"
      Height          =   492
      Left            =   240
      MouseIcon       =   "Computer Zeit Programm.frx":1341
      MousePointer    =   99  'Benutzerdefiniert
      OLEDropMode     =   1  'Manuell
      TabIndex        =   46
      Top             =   3720
      Width           =   1332
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   3360
   End
   Begin VB.TextBox txt_Minute 
      Height          =   288
      Left            =   2880
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox txt_Stunde 
      Height          =   288
      Left            =   2160
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CheckBox ckb_Stoppuhr 
      Caption         =   "Check1"
      Height          =   200
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":1493
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   19
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   200
   End
   Begin VB.CheckBox ckb_Uhrzeit 
      Caption         =   "Check1"
      Height          =   200
      Left            =   2880
      MouseIcon       =   "Computer Zeit Programm.frx":15E5
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   20
      Top             =   120
      Width           =   200
   End
   Begin VB.CommandButton cmd_ausführen 
      Caption         =   "Ausführen"
      Height          =   372
      Left            =   2800
      MouseIcon       =   "Computer Zeit Programm.frx":1737
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   11
      Top             =   4440
      Width           =   1000
   End
   Begin VB.TextBox txt_d 
      Height          =   288
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   372
   End
   Begin VB.TextBox txt_s 
      Height          =   288
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   372
   End
   Begin VB.TextBox txt_min 
      Height          =   288
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   372
   End
   Begin VB.TextBox txt_h 
      Height          =   288
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   372
   End
   Begin VB.OptionButton opt_6 
      BackColor       =   &H0000FF00&
      Caption         =   "Auswahl 6"
      Height          =   252
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":1889
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   9
      Top             =   4080
      Width           =   1092
   End
   Begin VB.OptionButton opt_5 
      BackColor       =   &H0000FF00&
      Caption         =   "Auswahl 5"
      Height          =   252
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":19DB
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   8
      Top             =   3840
      Width           =   1092
   End
   Begin VB.OptionButton opt_4 
      BackColor       =   &H0000FF00&
      Caption         =   "Auswahl 4"
      Height          =   252
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":1B2D
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   7
      Top             =   3600
      Width           =   1092
   End
   Begin VB.OptionButton opt_3 
      BackColor       =   &H0000FF00&
      Caption         =   "Auswahl 3"
      Height          =   252
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":1C7F
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   6
      Top             =   3360
      Width           =   1092
   End
   Begin VB.OptionButton opt_2 
      BackColor       =   &H0000FF00&
      Caption         =   "Auswahl 2"
      Height          =   252
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":1DD1
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   5
      Top             =   3120
      Width           =   1092
   End
   Begin VB.OptionButton opt_1 
      BackColor       =   &H0000FF00&
      Caption         =   "Auswahl 1"
      Height          =   252
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":1F23
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   4
      Top             =   2880
      Width           =   1092
   End
   Begin VB.OptionButton opt_8 
      BackColor       =   &H0000FF00&
      Caption         =   "Auswahl 8"
      Height          =   252
      Left            =   1680
      MouseIcon       =   "Computer Zeit Programm.frx":2075
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   10
      Top             =   4560
      Value           =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmd_1 
      Caption         =   "Sperren"
      Height          =   492
      Left            =   240
      MouseIcon       =   "Computer Zeit Programm.frx":21C7
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   12
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton cmd_2 
      Caption         =   "Abmelden"
      Height          =   492
      Left            =   240
      MouseIcon       =   "Computer Zeit Programm.frx":2319
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   13
      Top             =   720
      Width           =   1332
   End
   Begin VB.CommandButton cmd_3 
      Caption         =   "Neu Starten"
      Height          =   492
      Left            =   240
      MouseIcon       =   "Computer Zeit Programm.frx":246B
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   14
      Top             =   1320
      Width           =   1332
   End
   Begin VB.CommandButton cmd_4 
      Caption         =   "Herunterfahren"
      Height          =   492
      Left            =   240
      MouseIcon       =   "Computer Zeit Programm.frx":25BD
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   15
      Top             =   1920
      Width           =   1332
   End
   Begin VB.CommandButton cmd_5 
      Caption         =   "Herunterfahren"
      Height          =   492
      Left            =   240
      MouseIcon       =   "Computer Zeit Programm.frx":270F
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   16
      Top             =   2520
      Width           =   1332
   End
   Begin VB.CommandButton cmd_6 
      Caption         =   "Herunterfahren PC ausschalten"
      Height          =   492
      Left            =   240
      MouseIcon       =   "Computer Zeit Programm.frx":2861
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   17
      Top             =   3120
      Width           =   1332
   End
   Begin VB.CommandButton cmd_8 
      Caption         =   "Beenden"
      Height          =   492
      Left            =   240
      MouseIcon       =   "Computer Zeit Programm.frx":29B3
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   18
      Top             =   4320
      Width           =   1332
   End
   Begin VB.Label lbl_7 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   47
      Top             =   3840
      Width           =   252
   End
   Begin VB.Label lbl_Minute_hilfe 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "( 0 - 59 )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2640
      TabIndex        =   45
      Top             =   2040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lbl_Stunde_hilfe 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "( 0 - 23 )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1800
      TabIndex        =   44
      Top             =   2040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lbl_s_hilfe 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "( 0 - 59 )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2760
      TabIndex        =   43
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lbl_min_hilfe 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "( 0 - 59 )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2760
      TabIndex        =   42
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label lbl_h_hilfe 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "( 0 - 23 )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2760
      TabIndex        =   41
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label lbl_d_hilfe 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "( 0 - ... )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2640
      TabIndex        =   40
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Label lbl_Minute 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Minute:"
      Height          =   252
      Left            =   2760
      TabIndex        =   39
      Top             =   1440
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label lbl_Stunde 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Stunde:"
      Height          =   252
      Left            =   1920
      TabIndex        =   38
      Top             =   1440
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label lbl_Doppelpunkt1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2640
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label lbl_Stoppuhr 
      BackStyle       =   0  'Transparent
      Caption         =   "Stoppuhr"
      Height          =   252
      Left            =   1920
      TabIndex        =   36
      Top             =   120
      Width           =   732
   End
   Begin VB.Label lbl_Uhrzeit 
      BackStyle       =   0  'Transparent
      Caption         =   "Uhrzeit"
      Height          =   252
      Left            =   3120
      TabIndex        =   35
      Top             =   120
      Width           =   612
   End
   Begin VB.Label lbl_s 
      BackStyle       =   0  'Transparent
      Caption         =   "Sekunden"
      Height          =   252
      Left            =   2160
      TabIndex        =   34
      Top             =   2400
      Width           =   732
   End
   Begin VB.Label lbl_min 
      BackStyle       =   0  'Transparent
      Caption         =   "Minuten"
      Height          =   252
      Left            =   2160
      TabIndex        =   33
      Top             =   2040
      Width           =   732
   End
   Begin VB.Label lbl_h 
      BackStyle       =   0  'Transparent
      Caption         =   "Stunden"
      Height          =   252
      Left            =   2160
      TabIndex        =   32
      Top             =   1680
      Width           =   732
   End
   Begin VB.Label lbl_d 
      BackStyle       =   0  'Transparent
      Caption         =   "Tage"
      Height          =   252
      Left            =   2160
      TabIndex        =   31
      Top             =   1320
      Width           =   732
   End
   Begin VB.Label lbl_Erklärung 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Gebe eine Zeit an, nachdem das Programm ausgeführt wird:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1680
      TabIndex        =   30
      Top             =   480
      Width           =   2052
   End
   Begin VB.Label lbl_8 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   29
      Top             =   4440
      Width           =   252
   End
   Begin VB.Label lbl_6 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   28
      Top             =   3240
      Width           =   252
   End
   Begin VB.Label lbl_5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   27
      Top             =   2640
      Width           =   252
   End
   Begin VB.Label lbl_4 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   26
      Top             =   2040
      Width           =   252
   End
   Begin VB.Label lbl_3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   25
      Top             =   1440
      Width           =   252
   End
   Begin VB.Label lbl_2 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   24
      Top             =   840
      Width           =   252
   End
   Begin VB.Label lbl_1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   23
      Top             =   240
      Width           =   252
   End
End
Attribute VB_Name = "frm_Zeit_Programm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Stopp, Auswahl, Bestätigung, Bug_Anwendung As Integer
Dim Tage, Stunden, Minuten, Sekunden As Integer
Dim N_Stunden, N_Minuten As Integer
Dim Datei As String
Dim Programm_Verzeichnis, Laufwerksbuchstabe As String
Dim comdlg32 As Integer

Dim Result As Long
Private Declare Function LockWorkStation Lib "user32.dll" () As Long

Private Declare Function ExitWindowsEx Lib "user32" _
  (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Const EWX_LOGOFF = 0

Private Declare Function GetSystemDirectory Lib "kernel32" _
    Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

Function GetSysDir() As String
  Dim strDir  As String
  Dim nLen    As Long

  strDir = Space(255)
  nLen = GetSystemDirectory(strDir, 255)
  strDir = Left(strDir, nLen)

  If Right$(strDir, 1) <> "\" Then strDir = strDir & "\"

  GetSysDir = strDir
End Function

Public Function Neu()
Timer1.Enabled = False
Bestätigung = 0
Stopp = 0
cmd_ausführen.Caption = "Ausführen"
ckb_Stoppuhr.Value = 1
Ckb_Stoppuhr_Click
txt_d.Locked = False
txt_h.Locked = False
txt_min.Locked = False
txt_s.Locked = False
txt_Stunde.Locked = False
txt_Minute.Locked = False
txt_d.Text = ""
txt_h.Text = ""
txt_min.Text = ""
txt_s.Text = ""
txt_Stunde.Text = ""
txt_Minute.Text = ""
txt_d.SetFocus
ckb_Stoppuhr.Enabled = True
ckb_Uhrzeit.Enabled = True
lbl_Stoppuhr.ForeColor = RGB(0, 0, 0)
lbl_Uhrzeit.ForeColor = RGB(0, 0, 0)
lbl_Erklärung.ForeColor = RGB(0, 0, 0)
lbl_d.ForeColor = RGB(0, 0, 0)
lbl_h.ForeColor = RGB(0, 0, 0)
lbl_min.ForeColor = RGB(0, 0, 0)
lbl_s.ForeColor = RGB(0, 0, 0)
lbl_d_hilfe.ForeColor = RGB(0, 0, 0)
lbl_h_hilfe.ForeColor = RGB(0, 0, 0)
lbl_min_hilfe.ForeColor = RGB(0, 0, 0)
lbl_s_hilfe.ForeColor = RGB(0, 0, 0)
lbl_1.ForeColor = RGB(0, 0, 0)
lbl_2.ForeColor = RGB(0, 0, 0)
lbl_3.ForeColor = RGB(0, 0, 0)
lbl_4.ForeColor = RGB(0, 0, 0)
lbl_5.ForeColor = RGB(0, 0, 0)
lbl_6.ForeColor = RGB(0, 0, 0)
lbl_7.ForeColor = RGB(0, 0, 0)
lbl_8.ForeColor = RGB(0, 0, 0)
lbl_Stunde.ForeColor = RGB(0, 0, 0)
lbl_Minute.ForeColor = RGB(0, 0, 0)
lbl_Doppelpunkt1.ForeColor = RGB(0, 0, 0)
lbl_Stunde_hilfe.ForeColor = RGB(0, 0, 0)
lbl_Minute_hilfe.ForeColor = RGB(0, 0, 0)
opt_8.Value = True
opt_1.Enabled = True
opt_2.Enabled = True
opt_3.Enabled = True
opt_4.Enabled = True
opt_5.Enabled = True
opt_6.Enabled = True
opt_8.Enabled = True
cmd_1.Enabled = True
cmd_2.Enabled = True
cmd_3.Enabled = True
cmd_4.Enabled = True
cmd_5.Enabled = True
cmd_6.Enabled = True
cmd_7.Enabled = True
cmd_8.Enabled = True

End Function




Private Sub Form_Load()
'Programm_Verzeichnis = GetSysDir     'Gibt aktuelles Verzeichnis an
'Laufwerksbuchstabe = Mid(Programm_Verzeichnis, 1, 1)  'Filtert erstes Zeichen (Laufwerkbuchstaben) herraus
'If Dir(Laufwerksbuchstabe & ":\Windows\System32\comdlg32.ocx") <> "" Then
'comdlg32 = 1
'Else
'comdlg32 = 0
'End If

'If (comdlg32 = 0) Then
'cmd_7.Enabled = False
'End If

    On Error Resume Next
    
    Dim objWMI As Object
    
    'Prüfen ob WMI installiert ist
    Set objWMI = GetObject("WinMgmts:")
    
    If Err.Number <> 0 Then
        Err.Clear
    End If
    
opt_7.Enabled = False
End Sub

Private Sub cmd_1_click()
If (Bestätigung = 0) Then
  If MsgBox("Möchten Sie den Computer wirklich sperren?", _
         vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If
End If

Call Neu
LockWorkStation
End Sub

Private Sub cmd_2_click()
If (Bestätigung = 0) Then
  If MsgBox("Möchten Sie den Benutzer wirklich abmelden?", _
         vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If
End If
Call Neu
Result = ExitWindowsEx(EWX_LOGOFF, 0)
End Sub

Private Sub cmd_3_click()
    'Einfache Variante
    On Error Resume Next
    
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object

If (Bestätigung = 0) Then
  If MsgBox("Möchten Sie den Computer wirklich neu Starten?", _
         vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If
End If
Call Neu
    
    'WMI-Objekt erstellen und Abfragen ausführen
    Set objWMI = GetObject("WinMgmts:{impersonationLevel=impersonate, (Shutdown)}!/root/cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    
    
        'Windows neu starten
        For Each objItem In colItems
            objItem.Reboot
        Next objItem
End Sub

Private Sub cmd_4_click()
    'Einfache Variante
    On Error Resume Next
    
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object

If (Bestätigung = 0) Then
  If MsgBox("Möchten Sie den Computer wirklich Herunterfahren?", _
         vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If
End If
Call Neu
    
    'WMI-Objekt erstellen und Abfragen ausführen
    Set objWMI = GetObject("WinMgmts:{impersonationLevel=impersonate, (Shutdown)}!/root/cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")

        'Windows herunterfahren
        For Each objItem In colItems
            objItem.Shutdown
        Next objItem
End Sub

Private Sub cmd_5_click()
    'Komplexe Variante
    On Error Resume Next
    
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object
    
    Dim bytShutdownFlag As Byte

If (Bestätigung = 0) Then
  If MsgBox("Möchten Sie den Computer wirklich Herunterfahren?", _
         vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If
End If
Call Neu
    
    'WMI-Objekt erstellen und Abfragen ausführen
    Set objWMI = GetObject("WinMgmts:{impersonationLevel=impersonate, (Shutdown)}!/root/cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")

    'Windows herunterfahren
    bytShutdownFlag = 1
    
    'Funktion erzwingen
        bytShutdownFlag = bytShutdownFlag + 4
    
    'Auswahl ausführen
    For Each objItem In colItems
        objItem.Win32Shutdown (bytShutdownFlag)
    Next objItem
End Sub

Private Sub cmd_6_click()
    'Komplexe Variante
    On Error Resume Next
    
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object
    
    Dim bytShutdownFlag As Byte

If (Bestätigung = 0) Then
  If MsgBox("Möchten Sie den Computer wirklich Herunterfahren und das der PC sich ausschaltet?", _
         vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If
End If
Call Neu
    
    'WMI-Objekt erstellen und Abfragen ausführen
    Set objWMI = GetObject("WinMgmts:{impersonationLevel=impersonate, (Shutdown)}!/root/cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    
    'Windows herunterfahren - PC ausschalten
    bytShutdownFlag = 8
    
    'Funktion erzwingen
        bytShutdownFlag = bytShutdownFlag + 4
    
    'Auswahl ausführen
    For Each objItem In colItems
        objItem.Win32Shutdown (bytShutdownFlag)
    Next objItem
End Sub

Private Sub cmd_7_Click()
On Error GoTo Fehler

With frm_CommonDialog.CommonDialog1
    .Filter = ".exe|*.exe"
    .ShowOpen
    Datei = .FileName
  End With
  
If Dir(Datei) <> "" Then
  opt_7.Enabled = True
Else
  opt_7.Enabled = False
End If

If (Datei = "*.exe") Then
  opt_7.Enabled = False
Else
  opt_7.Enabled = True
End If

Fehler:
End Sub

Private Sub cmd_8_click()
If (Bestätigung = 0) Then
  If MsgBox("Möchten Sie das Programm wirklich beenden?", _
         vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If
End If
Call Neu
    'Beenden
    Unload Me
End Sub

Private Sub Ckb_Stoppuhr_Click()
If (ckb_Stoppuhr.Value = 0) Then
Auswahl = 1
ckb_Uhrzeit.Value = 1
lbl_Erklärung.Caption = "Gebe eine Uhrzeit an, an dem das Programm ausgeführt wird:"
txt_d.Visible = False
txt_h.Visible = False
txt_min.Visible = False
txt_s.Visible = False
lbl_d.Visible = False
lbl_h.Visible = False
lbl_min.Visible = False
lbl_s.Visible = False
lbl_d_hilfe.Visible = False
lbl_h_hilfe.Visible = False
lbl_min_hilfe.Visible = False
lbl_s_hilfe.Visible = False
txt_Stunde.Visible = True
txt_Minute.Visible = True
lbl_Stunde.Visible = True
lbl_Minute.Visible = True
lbl_Doppelpunkt1.Visible = True
lbl_Stunde_hilfe.Visible = True
lbl_Minute_hilfe.Visible = True
txt_Stunde.SetFocus

ElseIf (ckb_Stoppuhr.Value = 1) Then
Auswahl = 0
ckb_Uhrzeit.Value = 0
lbl_Erklärung.Caption = "Gebe eine Zeit an, nachdem das Programm ausgeführt wird:"
txt_Stunde.Visible = False
txt_Minute.Visible = False
lbl_Stunde.Visible = False
lbl_Minute.Visible = False
lbl_Doppelpunkt1.Visible = False
lbl_Stunde_hilfe.Visible = False
lbl_Minute_hilfe.Visible = False
txt_d.Visible = True
txt_h.Visible = True
txt_min.Visible = True
txt_s.Visible = True
lbl_d.Visible = True
lbl_h.Visible = True
lbl_min.Visible = True
lbl_s.Visible = True
lbl_d_hilfe.Visible = True
lbl_h_hilfe.Visible = True
lbl_min_hilfe.Visible = True
lbl_s_hilfe.Visible = True
txt_d.SetFocus
Else
End If
End Sub

Private Sub Ckb_Uhrzeit_Click()
If (ckb_Uhrzeit.Value = 0) Then
Auswahl = 0
ckb_Stoppuhr.Value = 1
lbl_Erklärung.Caption = "Gebe eine Zeit an, nachdem das Programm ausgeführt wird:"
txt_Stunde.Visible = False
txt_Minute.Visible = False
lbl_Stunde.Visible = False
lbl_Minute.Visible = False
lbl_Doppelpunkt1.Visible = False
lbl_Stunde_hilfe.Visible = False
lbl_Minute_hilfe.Visible = False
txt_d.Visible = True
txt_h.Visible = True
txt_min.Visible = True
txt_s.Visible = True
lbl_d.Visible = True
lbl_h.Visible = True
lbl_min.Visible = True
lbl_s.Visible = True
lbl_d_hilfe.Visible = True
lbl_h_hilfe.Visible = True
lbl_min_hilfe.Visible = True
lbl_s_hilfe.Visible = True
ElseIf (ckb_Uhrzeit.Value = 1) Then
Auswahl = 1
ckb_Stoppuhr.Value = 0
lbl_Erklärung.Caption = "Gebe eine Uhrzeit an, an dem das Programm ausgeführt wird:"
txt_d.Visible = False
txt_h.Visible = False
txt_min.Visible = False
txt_s.Visible = False
lbl_d.Visible = False
lbl_h.Visible = False
lbl_min.Visible = False
lbl_s.Visible = False
lbl_d_hilfe.Visible = False
lbl_h_hilfe.Visible = False
lbl_min_hilfe.Visible = False
lbl_s_hilfe.Visible = False
txt_Stunde.Visible = True
txt_Minute.Visible = True
lbl_Stunde.Visible = True
lbl_Minute.Visible = True
lbl_Doppelpunkt1.Visible = True
lbl_Stunde_hilfe.Visible = True
lbl_Minute_hilfe.Visible = True
Else
End If
End Sub

Private Sub cmd_ausführen_Click()

If (Stopp = 0) Then

If (Auswahl = 0) Then
Tage = Val(txt_d)
Stunden = Val(txt_h)
Minuten = Val(txt_min)
Sekunden = Val(txt_s)

If (Tage < 0) Then
Tage = 0
txt_d.Text = Tage
End If
If (Stunden < 0) Then
Stunden = 0
txt_h.Text = Stunden
End If
If (Minuten < 0) Then
Minuten = 0
txt_min.Text = Minuten
End If
If (Sekunden < 1) Then
Sekunden = 1
txt_s.Text = Sekunden
End If

If (Stunden > 23) Then
Stunden = 23
txt_h.Text = Stunden
End If
If (Minuten > 59) Then
Minuten = 59
txt_min.Text = Minuten
End If
If (Sekunden > 59) Then
Sekunden = 59
txt_s.Text = Sekunden
End If
ElseIf (Auswahl = 1) Then
Stunden = Val(txt_Stunde)
Minuten = Val(txt_Minute)
If (Stunden < 0) Then
Stunden = 0
End If
If (Stunden > 23) Then
Stunden = 23
End If
If (Minuten < 0) Then
Minuten = 0
End If
If (Minuten > 59) Then
Minuten = 59
End If
Else
End If
txt_Stunde.Text = Stunden
txt_Minute.Text = Minuten
Bestätigung = 1
Stopp = 1
cmd_ausführen.Caption = "Stopp"
Timer1.Enabled = True
txt_d.Locked = True
txt_h.Locked = True
txt_min.Locked = True
txt_s.Locked = True
txt_Stunde.Locked = True
txt_Minute.Locked = True
ckb_Stoppuhr.Enabled = False
ckb_Uhrzeit.Enabled = False
lbl_Stoppuhr.ForeColor = RGB(160, 160, 160)
lbl_Uhrzeit.ForeColor = RGB(160, 160, 160)
lbl_Erklärung.ForeColor = RGB(160, 160, 160)
lbl_d.ForeColor = RGB(160, 160, 160)
lbl_h.ForeColor = RGB(160, 160, 160)
lbl_min.ForeColor = RGB(160, 160, 160)
lbl_s.ForeColor = RGB(160, 160, 160)
lbl_d_hilfe.ForeColor = RGB(160, 160, 160)
lbl_h_hilfe.ForeColor = RGB(160, 160, 160)
lbl_min_hilfe.ForeColor = RGB(160, 160, 160)
lbl_s_hilfe.ForeColor = RGB(160, 160, 160)
lbl_1.ForeColor = RGB(160, 160, 160)
lbl_2.ForeColor = RGB(160, 160, 160)
lbl_3.ForeColor = RGB(160, 160, 160)
lbl_4.ForeColor = RGB(160, 160, 160)
lbl_5.ForeColor = RGB(160, 160, 160)
lbl_6.ForeColor = RGB(160, 160, 160)
lbl_7.ForeColor = RGB(160, 160, 160)
lbl_8.ForeColor = RGB(160, 160, 160)
lbl_Stunde.ForeColor = RGB(160, 160, 160)
lbl_Minute.ForeColor = RGB(160, 160, 160)
lbl_Doppelpunkt1.ForeColor = RGB(160, 160, 160)
lbl_Stunde_hilfe.ForeColor = RGB(160, 160, 160)
lbl_Minute_hilfe.ForeColor = RGB(160, 160, 160)
opt_1.Enabled = False
opt_2.Enabled = False
opt_3.Enabled = False
opt_4.Enabled = False
opt_5.Enabled = False
opt_6.Enabled = False

If (opt_7.Enabled = False) Then
Bug_Anwendung = 0
Else
Bug_Anwendung = 1
End If
opt_7.Enabled = False

opt_8.Enabled = False
cmd_1.Enabled = False
cmd_2.Enabled = False
cmd_3.Enabled = False
cmd_4.Enabled = False
cmd_5.Enabled = False
cmd_6.Enabled = False
cmd_7.Enabled = False
cmd_8.Enabled = False
ElseIf (Stopp = 1) Then
Bestätigung = 0
Stopp = 0
cmd_ausführen.Caption = "Ausführen"
Timer1.Enabled = False
txt_d.Locked = False
txt_h.Locked = False
txt_min.Locked = False
txt_s.Locked = False
txt_Stunde.Locked = False
txt_Minute.Locked = False
ckb_Stoppuhr.Enabled = True
ckb_Uhrzeit.Enabled = True
lbl_Stoppuhr.ForeColor = RGB(0, 0, 0)
lbl_Uhrzeit.ForeColor = RGB(0, 0, 0)
lbl_Erklärung.ForeColor = RGB(0, 0, 0)
lbl_d.ForeColor = RGB(0, 0, 0)
lbl_h.ForeColor = RGB(0, 0, 0)
lbl_min.ForeColor = RGB(0, 0, 0)
lbl_s.ForeColor = RGB(0, 0, 0)
lbl_d_hilfe.ForeColor = RGB(0, 0, 0)
lbl_h_hilfe.ForeColor = RGB(0, 0, 0)
lbl_min_hilfe.ForeColor = RGB(0, 0, 0)
lbl_s_hilfe.ForeColor = RGB(0, 0, 0)
lbl_1.ForeColor = RGB(0, 0, 0)
lbl_2.ForeColor = RGB(0, 0, 0)
lbl_3.ForeColor = RGB(0, 0, 0)
lbl_4.ForeColor = RGB(0, 0, 0)
lbl_5.ForeColor = RGB(0, 0, 0)
lbl_6.ForeColor = RGB(0, 0, 0)
lbl_7.ForeColor = RGB(0, 0, 0)
lbl_8.ForeColor = RGB(0, 0, 0)
lbl_Stunde.ForeColor = RGB(0, 0, 0)
lbl_Minute.ForeColor = RGB(0, 0, 0)
lbl_Doppelpunkt1.ForeColor = RGB(0, 0, 0)
lbl_Stunde_hilfe.ForeColor = RGB(0, 0, 0)
lbl_Minute_hilfe.ForeColor = RGB(0, 0, 0)
opt_1.Enabled = True
opt_2.Enabled = True
opt_3.Enabled = True
opt_4.Enabled = True
opt_5.Enabled = True
opt_6.Enabled = True

If (Bug_Anwendung = 0) Then
opt_7.Enabled = False
Else
opt_7.Enabled = True
End If

opt_8.Enabled = True
cmd_1.Enabled = True
cmd_2.Enabled = True
cmd_3.Enabled = True
cmd_4.Enabled = True
cmd_5.Enabled = True
cmd_6.Enabled = True
cmd_7.Enabled = True
cmd_8.Enabled = True

Else
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
If (Auswahl = 0) Then       'Stoppuhr
Tage = Val(txt_d)
Stunden = Val(txt_h)
Minuten = Val(txt_min)
Sekunden = Val(txt_s)
Sekunden = Sekunden - 1
    If (Sekunden < 0) Then
        If (Minuten > 0) Then
        Minuten = Minuten - 1
        Sekunden = 59
        Else
            If (Stunden > 0) Then
                Stunden = Stunden - 1
                Minuten = 59
                Sekunden = 59
            Else
                If (Tage > 0) Then
                Tage = Tage - 1
                Stunden = 23
                Minuten = 59
                Sekunden = 59
                Else
                'Ausführen
                End If
            End If
        End If
    End If
txt_d.Text = Tage
txt_h.Text = Stunden
txt_min.Text = Minuten
txt_s.Text = Sekunden

If (Tage = 0) Then
    If (Stunden = 0) Then
        If (Minuten = 0) Then
            If (Sekunden = 0) Then

                If (opt_1 = True) Then
                cmd_1_click
                ElseIf (opt_2 = True) Then
                cmd_2_click
                ElseIf (opt_3 = True) Then
                cmd_3_click
                ElseIf (opt_4 = True) Then
                cmd_4_click
                ElseIf (opt_5 = True) Then
                cmd_5_click
                ElseIf (opt_6 = True) Then
                cmd_6_click
                ElseIf (opt_7 = True) Then
                Shell (Datei), vbNormalFocus
                Call Neu
                Unload Me
                ElseIf (opt_8 = True) Then
                cmd_8_click
                Else
                End If
            End If
        End If
    End If
End If

ElseIf (Auswahl = 1) Then   'Uhrzeit
Stunden = Val(txt_Stunde)
Minuten = Val(txt_Minute)

N_Stunden = Hour(Now)
N_Minuten = Minute(Now)

If (Stunden = N_Stunden) Then
    If (Minuten = N_Minuten) Then
        If (opt_1 = True) Then
        cmd_1_click
        ElseIf (opt_2 = True) Then
        cmd_2_click
        ElseIf (opt_3 = True) Then
        cmd_3_click
        ElseIf (opt_4 = True) Then
        cmd_4_click
        ElseIf (opt_5 = True) Then
        cmd_5_click
        ElseIf (opt_6 = True) Then
        cmd_6_click
        ElseIf (opt_7 = True) Then
        Shell (Datei), vbNormalFocus
        Call Neu
        Unload Me
        ElseIf (opt_8 = True) Then
        cmd_8_click
        Else
        End If
    End If
End If

Else
End If
End Sub
