VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form B_Rex 
   Appearance      =   0  '2D
   BackColor       =   &H00C0C0C0&
   Caption         =   "B_Rex"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15810
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "B_rex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   10635
   ScaleWidth      =   15810
   Visible         =   0   'False
   WindowState     =   2  'Maximiert
   Begin VB.HScrollBar Horizontal 
      Height          =   240
      Left            =   6840
      SmallChange     =   30
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.VScrollBar Vertikal 
      Height          =   672
      LargeChange     =   4
      Left            =   7560
      SmallChange     =   30
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox FuKurve 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1032
      Left            =   4680
      ScaleHeight     =   1008
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   948
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   972
      Begin VB.Line Auflegedehnung 
         Visible         =   0   'False
         X1              =   60.19
         X2              =   780.47
         Y1              =   659.964
         Y2              =   659.964
      End
      Begin VB.Line Auflegedehnunganzeiger 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   60.19
         X2              =   780.47
         Y1              =   480.43
         Y2              =   480.43
      End
   End
   Begin VB.PictureBox Eigenschaftsleiste 
      Align           =   4  'Rechts ausrichten
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   10125
      Left            =   12075
      ScaleHeight     =   10095
      ScaleWidth      =   3705
      TabIndex        =   3
      Top             =   510
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   2220
         Picture         =   "B_rex.frx":000C
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   28
         Tag             =   "A#03182#"
         ToolTipText     =   "Schwingungsrechnung"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   1920
         Picture         =   "B_rex.frx":05E2
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   25
         Tag             =   "A#03178#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   1680
         Picture         =   "B_rex.frx":0BB8
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   22
         Tag             =   "A#03175#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   240
         Picture         =   "B_rex.frx":118E
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   21
         Tag             =   "#03166#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   540
         Picture         =   "B_rex.frx":1764
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   20
         Tag             =   "#03167#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   0
         Picture         =   "B_rex.frx":1D3A
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   19
         Tag             =   "A#03165#"
         Top             =   0
         Width           =   330
      End
      Begin VB.TextBox Eingabe 
         Appearance      =   0  '2D
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   855
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.ComboBox Auswahl 
         Appearance      =   0  '2D
         Height          =   288
         Left            =   120
         TabIndex        =   11
         Text            =   "Auswahl"
         Top             =   1200
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   3360
         Picture         =   "B_rex.frx":2310
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   10
         Tag             =   "E#03164#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   3060
         Picture         =   "B_rex.frx":28E6
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   9
         Tag             =   "A#03163#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   2760
         Picture         =   "B_rex.frx":2EBC
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   8
         Tag             =   "E#03162#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1440
         Picture         =   "B_rex.frx":3492
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   7
         Tag             =   "A#03332#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   1200
         Picture         =   "B_rex.frx":3A68
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   6
         Tag             =   "A#03331#"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox EigButton 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   960
         Picture         =   "B_rex.frx":403E
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   5
         Tag             =   "E#03330#"
         Top             =   0
         Width           =   330
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Eiged 
         Height          =   492
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   2052
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   12632256
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         MergeCells      =   2
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin PicClip.PictureClip Elbilder 
      Left            =   6840
      Top             =   3420
      _ExtentX        =   10239
      _ExtentY        =   12568
      _Version        =   393216
      Picture         =   "B_rex.frx":4614
   End
   Begin VB.PictureBox Kopfleiste 
      Align           =   1  'Oben ausrichten
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1052
      TabIndex        =   0
      Top             =   0
      Width           =   15810
      Begin VB.PictureBox Button 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   2640
         Picture         =   "B_rex.frx":8B62A
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   27
         Tag             =   "A"
         ToolTipText     =   "Zweischeibenantrieb"
         Top             =   0
         Width           =   750
      End
      Begin VB.ComboBox Beispielanlagen 
         Height          =   315
         ItemData        =   "B_rex.frx":8C934
         Left            =   6720
         List            =   "B_rex.frx":8C936
         Style           =   2  'Dropdown-Liste
         TabIndex        =   26
         Top             =   60
         Width           =   4275
      End
      Begin VB.TextBox Trumlänge 
         CausesValidation=   0   'False
         Height          =   240
         Left            =   10800
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   4
         Left            =   1500
         Picture         =   "B_rex.frx":8C938
         ToolTipText     =   "alle berechnen"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   11
         Left            =   6000
         Picture         =   "B_rex.frx":8D4BA
         Tag             =   "A#03177#"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   10
         Left            =   5520
         Picture         =   "B_rex.frx":8E03C
         Tag             =   "A#03176#"
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Trumlängeneinheit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Truml. [mm]:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   9
         Left            =   1980
         Picture         =   "B_rex.frx":8EBBE
         Tag             =   "A#03110#"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   8
         Left            =   4860
         Picture         =   "B_rex.frx":8F740
         Tag             =   "A#03161#"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   7
         Left            =   4500
         Picture         =   "B_rex.frx":902C2
         Tag             =   "A#03160"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   6
         Left            =   4140
         Picture         =   "B_rex.frx":90E44
         Tag             =   "A#03159#"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   5
         Left            =   3780
         Picture         =   "B_rex.frx":919C6
         Tag             =   "E#03158#"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   3
         Left            =   1020
         Picture         =   "B_rex.frx":92548
         Tag             =   "105"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   2
         Left            =   675
         Picture         =   "B_rex.frx":930CA
         Tag             =   "103"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   1
         Left            =   420
         Picture         =   "B_rex.frx":93C4C
         Tag             =   "102"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Datei 
         Height          =   450
         Index           =   0
         Left            =   0
         Picture         =   "B_rex.frx":947CE
         Tag             =   "101"
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.CommandButton Defaultbutton 
      Caption         =   "Defaultbutton"
      Default         =   -1  'True
      Height          =   252
      Left            =   4320
      TabIndex        =   2
      Top             =   60
      Width           =   852
   End
   Begin VB.PictureBox Konstruktion 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2352
      Left            =   4440
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   3135
      Begin VB.PictureBox Bildspeicher 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   432
         Left            =   4200
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   43
         TabIndex        =   16
         Top             =   1740
         Visible         =   0   'False
         Width           =   672
      End
      Begin VB.PictureBox Elementleiste 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         FillStyle       =   0  'Ausgefüllt
         ForeColor       =   &H80000008&
         Height          =   1452
         Left            =   0
         ScaleHeight     =   95
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   328
         TabIndex        =   13
         Top             =   0
         Width           =   4944
         Begin VB.Image Element 
            Height          =   480
            Index           =   1
            Left            =   60
            Picture         =   "B_rex.frx":95350
            Tag             =   "001"
            Top             =   300
            Width           =   480
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   2
            Left            =   480
            Picture         =   "B_rex.frx":95B92
            Tag             =   "002"
            Top             =   300
            Width           =   480
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   3
            Left            =   900
            Picture         =   "B_rex.frx":963D4
            Tag             =   "003"
            Top             =   300
            Width           =   480
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   4
            Left            =   1320
            Picture         =   "B_rex.frx":96C16
            Tag             =   "005"
            Top             =   300
            Width           =   480
         End
         Begin VB.Image Element 
            Height          =   240
            Index           =   9
            Left            =   1800
            Picture         =   "B_rex.frx":97858
            Tag             =   "101"
            Top             =   360
            Width           =   1920
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   11
            Left            =   3360
            Picture         =   "B_rex.frx":9909A
            Tag             =   "103"
            Top             =   360
            Width           =   1920
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   7
            Left            =   900
            Picture         =   "B_rex.frx":9C0DC
            Tag             =   "201"
            Top             =   720
            Width           =   480
         End
         Begin VB.Image Element 
            Height          =   390
            Index           =   12
            Left            =   1800
            Picture         =   "B_rex.frx":9CD1E
            Tag             =   "104"
            Top             =   720
            Width           =   1920
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   6
            Left            =   480
            Picture         =   "B_rex.frx":9F460
            Tag             =   "204"
            Top             =   720
            Width           =   480
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   5
            Left            =   60
            Picture         =   "B_rex.frx":A00A2
            Tag             =   "205"
            Top             =   720
            Width           =   480
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   8
            Left            =   1320
            Picture         =   "B_rex.frx":A0CE4
            Tag             =   "206"
            Top             =   720
            Width           =   480
         End
         Begin VB.Image Element 
            Height          =   480
            Index           =   10
            Left            =   3360
            Picture         =   "B_rex.frx":A1926
            Tag             =   "102"
            Top             =   720
            Width           =   1920
         End
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Undurchsichtig
         Height          =   492
         Left            =   1320
         Top             =   1680
         Visible         =   0   'False
         Width           =   552
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   200
         X2              =   235
         Y1              =   165
         Y2              =   165
      End
      Begin VB.Shape Rahmen 
         BorderColor     =   &H80000002&
         BorderStyle     =   3  'Punkt
         Height          =   492
         Left            =   600
         Top             =   1680
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Image GrKl 
         Height          =   120
         Index           =   1
         Left            =   2220
         MousePointer    =   9  'Größenänderung W O
         Picture         =   "B_rex.frx":A4968
         Top             =   1920
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Image GrKl 
         Height          =   120
         Index           =   0
         Left            =   2040
         MousePointer    =   9  'Größenänderung W O
         Picture         =   "B_rex.frx":A4A6A
         Top             =   1920
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Image Element 
         Height          =   480
         Index           =   0
         Left            =   0
         Tag             =   "001"
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.TextBox Fehlerliste 
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      Height          =   1032
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "B_Rex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Dummy As String 'abfallstring
'Private MKE(10) As Boolean 'konfiguriert die eigenschaftsanzeige
'einbuchstabige variablennamen nur noch lokal zulassen!
Private X1 As Integer 'Altposition eines Elements merken bei Lageveränderung
Private Y1 As Integer
Private X2 As Integer 'merkt sich eine clickposition
Private Y2 As Integer
Private X3 As Integer 'Übergabevariable für Clickposition im Steuerelement
Private Y3 As Integer
Private Rasterungx As Integer ' Möglichkeit zur variablen Rasterung über Programm

Private Rasterungy As Integer
Private oGrenze As Integer ' Bereich wird zur Kollisions/Berührungsüberwachung definiert
Private uGrenze As Integer
Private lGrenze As Integer
Private rGrenze As Integer
'Private Markiert As Single 'enthält Nummer des derzeit markierten Objekts
Private Mark As Boolean 'merken, ob träger rechts oder links verändert wird
Private Träger As Integer
Private Zahlen$ 'as string
Private LastIndex As Integer
Private Anzeigemodus As Integer
'Private Anz(5) As Boolean
Private Noresize As Boolean
Private Markierte_Anzeigen As Boolean
Private DragDrop As Boolean
Private ActTypZeile As Integer
Private Anhalten As Boolean

'Private VerfügHöhe As Integer
'1 anlage
'2 tabelle
'3 umfangskraft
'4 fehler
'5 priorität, nicht benutzt

Private AktZeile As Integer 'merkt sich die zeile des textfeldes
'träger sind in wirklichkeit nach unten breiter, huckepacks nach oben
'ihre position dort wird in der sys-variable gespeichert, aber beim aufruf leicht verfälscht
Private Trägergrkl As Boolean 'bestimmt, ob rechter, oder linker rand verändert wird
Private Sub Beispielanlagen_Click()
Dim i As Integer
On Error GoTo Errorhandler
    Bspanlagen.MoveFirst
    Do Until Bspanlagen("bezeichnung") = Beispielanlagen Or Bspanlagen("bezeichnung_en") = Beispielanlagen Or Bspanlagen("bezeichnung_jp") = Beispielanlagen Or i > 1000 Or Bspanlagen.EOF
         i = i + 1
         Bspanlagen.MoveNext
    Loop
    If Bspanlagen("bezeichnung") = Beispielanlagen Or Bspanlagen("bezeichnung_en") = Beispielanlagen Or Bspanlagen("bezeichnung_jp") = Beispielanlagen Then
        Call Dateiverwaltung.Undo(4)
    End If
Errorhandler:
End Sub

Private Sub Button_Click(Index As Integer)
'Antrieb Extremultus, gleichmaessiger Betrieb

If Zweischeiben = True Then 'hat schon zweischeiben, will nur ansicht aendern
    Call CodeDraw.Alleelementeverbinden
Else
    If Gespeichert = False Then
        Dummy$ = MsgBox(Lang_Res(149), vbYesNo + vbQuestion + vbDefaultButton2)  'Die aktuelle Anlage wird gelöscht. Wollen Sie trotzdem eine neue Anlage konstruieren?
        If Dummy$ = vbNo Then Exit Sub
    End If
    
    Bspanlagen.MoveFirst
    Do
         If Bspanlagen("bezeichnung") = "Antrieb Extremultus, gleichmaessiger Betrieb" Then Exit Do
         i = i + 1
         Bspanlagen.MoveNext
    Loop Until Bspanlagen.EOF Or i > 1000
    Call Mother.Neue_Knopfverwaltung(Button(Index))
    
    Call Dateiverwaltung.Undo(4) 'macht er vielleicht nicht

End If

    If left(Button(0).Tag, 1) = "E" Then
        Elementleiste.Visible = False
        Beispielanlagen.Visible = False
    Else
        Elementleiste.Visible = True
        Beispielanlagen.Visible = True
    End If
'End If
End Sub

Private Sub Datei_DblClick(Index As Integer)
'ziel der übung: nur eine info soll groß dargestellt werden
Dim i As Integer
If Index < 5 Or Index > 8 Then Exit Sub

For i = 4 To 8 'alle schließen
    If left(Datei(i).Tag, 1) = "E" Then Call Datei_Click(i)
Next i
DoEvents
Call Datei_Click(Index) 'einen öffnen
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Anhalten = True

    'f9 zur neuaufnahme in Datenbank
        If KeyCode = 120 Then Call Dateiverwaltung.Undo(3)
    
    'f10 löschen aus der Datenbank
        If KeyCode = 121 Then Call Dateiverwaltung.Undo(5)
End Sub

Public Sub Form_Load()
Dim i, j As Integer
    Alleelementedargestellt = False
    On Local Error Resume Next
    Sys(1).Element = "Band" 'damits aufgerufen wird, ob was drinsteht oder nicht
    Sys(2).Element = "Band" 'damits mit gespeichert wird
    Sys(1).Tag = "301" 'damits aufgerufen wird, ob was drinsteht oder nicht
    Noresize = True
    B_Rexgeladen = True
    
    Dateneingabe = True
    Reversieren = False
    Maxelementindex = 10
    
    Kopfleiste.Height = Screen.TwipsPerPixelY * 35
    Elementleiste.Top = -1
    Elementleiste.Width = 425
    Elementleiste.Height = 75
    
    Element(1).left = 5 'primäre Antriebsscheibe
    Element(1).Top = 2 'primäre Antriebsscheibe
    Element(2).left = 40 'Antriebsscheibe
    Element(2).Top = 2
    Element(3).left = 75 'umlenkscheibe
    Element(3).Top = 2
    Element(4).left = 110 'messerkante
    Element(4).Top = 2
    
    Element(5).left = 5 'abweiser
    Element(5).Top = 37
    Element(6).left = 40 'stau
    Element(6).Top = 37
    Element(7).left = 75 'transportgut
    Element(7).Top = 37
    Element(8).left = 110 'freie umfangskraft
    Element(8).Top = 37
    
    Element(9).left = 155 'tisch
    Element(9).Top = 7
    Element(9).Width = 128
    Element(9).Height = 16
    
    Element(11).left = 285 'rollenbahn
    Element(11).Top = 7
    Element(11).Width = 128
    Element(11).Height = 20
    
    Element(12).left = 155 'freie Umfangskraft
    Element(12).Top = 42
    Element(12).Width = 128
    Element(12).Height = 27
    
    Element(10).left = 285 'tragrollenbahn
    Element(10).Top = 42
    Element(10).Width = 128
    Element(10).Height = 20
    

    For i = 0 To 11
        'If I <> 4 Then
            Datei(i).Top = 2
            Datei(i).Height = 30
            Datei(i).Width = 31
        'End If
    Next i
    i = 31 'dateibreite
    j = 6 'breite zwischen dateigruppen
    Datei(0).left = 2 'neu
    Datei(1).left = 2 + i 'laden
    Datei(2).left = 2 + 2 * i 'speichern
    Datei(3).left = 2 + 3 * i  'drucken
    Datei(4).left = 2 + 1 * j + 4 * i  'alle rechnen
    Datei(9).left = 2 + 2 * j + 5 * i 'reversieren
    
    Button(0).left = 2 + 3 * j + 6 * i
    Button(0).Top = 2
    
    Datei(5).left = 2 + 4 * j + 8 * i  'zeichnung
    Datei(6).left = 2 + 4 * j + 9 * i 'eigenschaftstabelle
    Datei(7).left = 2 + 4 * j + 10 * i 'Fu-Kurve
    Datei(8).left = 2 + 4 * j + 11 * i 'Fehlerliste
    Datei(10).left = 2 + 5 * j + 12 * i 'eins zurück
    Datei(11).left = 2 + 5 * j + 13 * i 'eins vorwärts
    
    Trumlängeneinheit.left = Datei(11).left + Datei(11).Width + 4
    Trumlängeneinheit.Top = 7
    Trumlänge.left = Trumlängeneinheit.left + Trumlängeneinheit.Width + 4
    Trumlänge.Top = 6
    Beispielanlagen.left = Datei(11).left + Datei(11).Width + 4
    Beispielanlagen.Top = 7


    Noresize = False
    
    Call Dateiverwaltung.Beispielanlagen_einrichten
End Sub

Private Sub Eiged_RowColChange()
    
    If Eingabe.Visible = True Or Auswahl.Visible = True Then  'sonst verrutscht die aktuelle zeile in der tabelle
        DoEvents
    End If
End Sub
Private Sub Auswahl_Click()
    On Error Resume Next
    Dim P As Integer
    'hat er überhaupt was geändert?
    If Eiged.TextMatrix(AktZeile, 8) = Auswahl Then Exit Sub
    
    If Auswahl <> "-" Then
        Eiged.TextMatrix(AktZeile, 8) = Auswahl
        Eiged.TextMatrix(AktZeile, 7) = Auswahl
        Select Case Eiged.TextMatrix(AktZeile, 3)
            Case 41, 42  'bandkontaktfläche zum element
                If Auswahl = El(-4).Eigenschaft Or Auswahl = Sys(1).S(4) Then 'LS
                    Sys(Eiged.TextMatrix(AktZeile, 2)).E(Eiged.TextMatrix(AktZeile, 3)) = 1
                Else 'sonst ts
                    Sys(Eiged.TextMatrix(AktZeile, 2)).E(Eiged.TextMatrix(AktZeile, 3)) = 2
                End If
            Case Else
                'erste eigenschaft füllen
                P = 0
                Do
                    P = P + 1
                Loop Until (LCase(Kst(P).Bezeichnung) = LCase(Eiged.TextMatrix(AktZeile, 8)) And Kst(P).zuEigenschaft = Eiged.TextMatrix(AktZeile, 3)) Or Kst(P).Bezeichnung = "letzter Satz"
                
                'ev. weitere füllen
                If Kst(P).Bezeichnung <> "letzter Satz" Then
                    Sys(Eiged.TextMatrix(AktZeile, 2)).E(Eiged.TextMatrix(AktZeile, 3)) = P
                    If Eiged.TextMatrix(AktZeile, 3) = 36 Then 'gehört noch eine zweite einstellung ins schlepp
                        P = 0
                        Do
                            P = P + 1
                        Loop Until (LCase(Kst(P).Bezeichnung) = LCase(Eiged.TextMatrix(AktZeile, 8)) And Kst(P).zuEigenschaft = 61) Or Kst(P).Bezeichnung = "letzter Satz"
                        Sys(Eiged.TextMatrix(AktZeile, 2)).E(61) = P
                    End If
                    If Eiged.TextMatrix(AktZeile, 3) = 47 Then 'gehört noch eine zweite einstellung ins schlepp, e-modul
                        P = 0
                        Do
                            P = P + 1
                        Loop Until (LCase(Kst(P).Bezeichnung) = LCase(Eiged.TextMatrix(AktZeile, 8)) And Kst(P).zuEigenschaft = 103) Or Kst(P).Bezeichnung = "letzter Satz"
                        Sys(Eiged.TextMatrix(AktZeile, 2)).E(103) = P
                    End If
                End If
        End Select
        
        'besonderheit, wird bereits vor der rechnung bei Vollstaendigkeit ausgeführt
        If Eiged.TextMatrix(AktZeile, 3) = 47 Then
            Call Eigschaftsverr.Massenträgheitsmoment
            TabelleEig_ausfuellen
            Mother.H = Lang_Res(180)  'manuelle Überlast wurde auf 0 gesetzt
        End If
        
        'automatische überlastvorgabe, wenn diese ungleich 0 ist
        If Eiged.TextMatrix(AktZeile, 3) = 60 And Sys(Eiged.TextMatrix(AktZeile, 2)).E(60) <> 22 Then
            Sys(Eiged.TextMatrix(AktZeile, 2)).E(59) = 0 'dann braucht er die manuelle je nicht mehr
            TabelleEig_ausfuellen  'nur weil hier auch nummerische daten geändert werden
        End If
        
    End If
    Call Dateiverwaltung.Undo(0) 'regelt undo und aktualität

    Auswahl.Visible = False
    Call CodeCalc.Rechnungssteuerung("C") 'bricht selbst ab, wenn nicht endlos
    
    If Endlos = True Then Call TabelleEig_ausfuellen

End Sub
Private Sub Auswahl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Auswahl_KeyPress(KeyCode) 'pfeil 40 nach unten, als wäre return gedrückt worden
    KeyCode = 0 'sonst wird die markierung gelöscht
End Sub
Private Sub Auswahl_KeyPress(KeyAscii As Integer)
    Eiged.Col = 8
    If KeyAscii = 40 Then KeyAscii = 13 'pfeil runter = return
    If KeyAscii = 13 Then 'return, eingabe abschliessen
        
        Call Auswahl_Click 'dort ist die ganze behandlung einmal zusammengefaßt
        
        If Eiged.Row < Eiged.Rows - 1 Then
            Eiged.Row = Eiged.Row + 1
            Do Until InStr(Eiged.TextMatrix(Eiged.Row, 0), "3") = 0 Or Eiged.Row >= Eiged.Rows - 1
                Eiged.Row = Eiged.Row + 1
            Loop
        End If
        Call Eiged_MouseDown(1, 0, Eiged.Width, 0) 'eingabe oder auswahl neu setzen
        Exit Sub
    End If
    If KeyAscii = 38 Then 'pfeil hoch, eingabe abschliessen
        
        Call Auswahl_Click 'dort ist die ganze behandlung einmal zusammengefaßt
        
        If Eiged.Row > 1 Then
            Eiged.Row = Eiged.Row - 1
            Do Until InStr(Eiged.TextMatrix(Eiged.Row, 0), "3") = 0 Or Eiged.Row <= 2
                Eiged.Row = Eiged.Row - 1
            Loop
        End If
        Call Eiged_MouseDown(1, 0, Eiged.Width, 0) 'irgendein element neu setzen
        Exit Sub
    End If
    KeyAscii = 0
End Sub
Private Sub Eingabe_Wertuebernehmen()
On Error GoTo Errorhandler
    'auswahlen werden woanders abgehandelt
    
    'plausibilitätsprüfung
    Select Case left(Eiged.TextMatrix(AktZeile, 1), 4)
        Case "text"
            If Eiged.TextMatrix(AktZeile, 8) <> Eingabe Then
                Eiged.TextMatrix(AktZeile, 8) = Eingabe
                Sys(Eiged.TextMatrix(AktZeile, 2)).S(Abs(Eiged.TextMatrix(AktZeile, 3))) = Eiged.TextMatrix(AktZeile, 8)
                Eiged.TextMatrix(AktZeile, 7) = Eingabe
                Call Dateiverwaltung.Undo(0)
                Call TabelleEig_ausfuellen  'z.b. tragseitentest/laufseitentext kommt in der tabelle mehrfach vor
            End If
        Case "zahl"
            'element, eigenschaft, alte einstellung, neue einstellung
            If Eingabe = "" Then Eingabe = "0" 'sonst stürzt cdbl ab
            If CDbl(Eingabe) <> CDbl(Eiged.TextMatrix(AktZeile, 8)) Then
                Call Eigschaftsverr.Verrechnung(Eiged.TextMatrix(AktZeile, 2), Eiged.TextMatrix(AktZeile, 3), CDbl(Eingabe), Eiged.TextMatrix(AktZeile, 8))
                If Abbruch = False Then Call TabelleEig_ausfuellen
                'undofunktionen werden dort ausgeführt
            End If
    End Select
    Eingabe.Visible = False
Exit Sub

Errorhandler:
    Beep
    Mother.H = Lang_Res(148)  'unzulässige eingabe
    Eingabe = ""
    Eingabe.Visible = False
End Sub
Private Sub Eingabe_KeyDown(KeyCode As Integer, Shift As Integer)
    Eingabe_KeyPress (KeyCode) 'pfeil 40 nach unten, als wäre return gedrückt worden
    KeyCode = 0 'sonst wird die markierung gelöscht
End Sub
Private Sub Eingabe_KeyPress(KeyAscii As Integer)
Dim a$
Dim Merk As Boolean
    
    'bei komplett ausgefüllter anlage nullwerte verhindern!!!
    If KeyAscii = 40 Then KeyAscii = 13 'pfeil runter = return
    If KeyAscii = 13 Then 'return, eingabe abschliessen
        Call Eingabe_Wertuebernehmen
        If Eiged.Row < Eiged.Rows - 1 Then
            Eiged.Row = Eiged.Row + 1
            Do Until InStr(Eiged.TextMatrix(Eiged.Row, 0), "3") = 0 Or Eiged.Row >= Eiged.Rows - 1 'kein ergebnis und nicht am ende
                Eiged.Row = Eiged.Row + 1
            Loop
        End If
        Call Eiged_MouseDown(0, 0, Eiged.Width, 0) 'irgendein element neu setzen, so als wäre rechts geklickt worden
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 38 Then 'pfeil hoch, eingabe abschliessen
        Call Eingabe_Wertuebernehmen
        If Eiged.Row > 1 Then
            Eiged.Row = Eiged.Row - 1
            Do Until InStr(Eiged.TextMatrix(Eiged.Row, 0), "3") = 0 Or Eiged.Row >= Eiged.Rows - 1 'kein ergebnis und nicht am ende
            'Do Until Eiged.TextMatrix(Eiged.Row, 0) = "1" Or Eiged.TextMatrix(Eiged.Row, 0) = "2" Or Eiged.Row <= 2
                Eiged.Row = Eiged.Row - 1
            Loop
        End If
        Call Eiged_MouseDown(0, 0, Eiged.Width, 0) 'irgendein element neu setzen, so als wäre rechts geklickt worden
        KeyAscii = 0
        Exit Sub
    End If
    
    If Eiged.TextMatrix(AktZeile, 1) = "text" Then Exit Sub 'eingeben, was man will
    
    'nur zahleneingabe
    If (KeyAscii > 47 And KeyAscii < 58) Then Merk = True
    If KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 46 Then Merk = True
    If Merk = False Then
        KeyAscii = 0
        Exit Sub
    End If
    
    'Länderanpassung komma / punkt
    If KeyAscii = 44 Or KeyAscii = 46 Then
        If Val(2.5) = 2.5 Then 'der benutzt punkte
            If KeyAscii = 44 Then KeyAscii = 46 ' aus komma mach punkt
        Else
            If KeyAscii = 46 Then KeyAscii = 44 'aus punkt mach komma
        End If
    End If
    
End Sub
Private Sub Eiged_Scroll()
    'sonst werden folgende elemente nicht mit versetzt
    If Eingabe.Visible = True Then Call Eingabe_Wertuebernehmen
    If Auswahl.Visible = True Then Call Auswahl_Click
    Eingabe.Visible = False
    Auswahl.Visible = False
End Sub
Private Sub Eiged_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim P As Integer
Dim i As Double, j As Double

    DoEvents
    Eiged.Col = 8
    If Button > 0 Then 'wurde von maus, nicht als unterprogramm aufgerufen
        If Eingabe.Visible = True And Eingabe <> Eiged Then Call Eingabe_Wertuebernehmen
        If Auswahl.Visible = True And Auswahl <> Eiged Then Call Auswahl_Click
    End If
    
    'damit man die auswahl auch mal wieder los wird
    If x < Eiged.Width / 5 * 3 Then
        Eingabe.Visible = False
        Auswahl.Visible = False
        If Eiged.TextMatrix(Eiged.Row, 2) = "" Then Exit Sub
        Eiged.Col = 6
        If Eiged.CellForeColor = QBColor(12) Then
            Eiged.CellForeColor = QBColor(0)
            Sys(Eiged.TextMatrix(Eiged.Row, 2)).B(Eiged.TextMatrix(Eiged.Row, 3)) = False
        Else
            Eiged.CellForeColor = QBColor(12)
            Sys(Eiged.TextMatrix(Eiged.Row, 2)).B(Eiged.TextMatrix(Eiged.Row, 3)) = True
        End If
        Call Dateiverwaltung.Undo(0) 'regelt undo und aktualität
        Exit Sub
    End If

    If Eiged.CellBackColor = vbWhite Then 'weiß
        If left(Eiged.TextMatrix(Eiged.Row, 1), 4) = "zahl" Or left(Eiged.TextMatrix(Eiged.Row, 1), 4) = "text" Then
            
            Eingabe.Top = Eiged.CellTop + Eiged.Top - 50
            Eingabe.left = Eiged.CellLeft + Eiged.left - 20
            Eingabe.Width = Eiged.CellWidth + 20
            'Eingabe.Height = Eiged.CellHeight
            Eingabe = Eiged
            Eingabe.Visible = True
            AktZeile = Eiged.Row
            Auswahl.Visible = False
            If Eingabe.Visible = True Then
                Eingabe.SetFocus
                Eingabe.SelStart = 0
                Eingabe.SelLength = Len(Eingabe)
            End If
        Else 'liste
            Auswahl.Top = Eiged.CellTop + Eiged.Top - 50
            Auswahl.left = Eiged.CellLeft + Eiged.left - 20
            Auswahl.Visible = True
            AktZeile = Eiged.Row
            Auswahl.SetFocus
            Auswahl.Clear
            Auswahl = Eiged
            Eingabe.Visible = False
        
            Select Case Eiged.TextMatrix(Eiged.Row, 3)
                Case 41 To 42 'bandkontaktfläche zum element
                    If Sys(1).S(4) = "" Then
                        Auswahl.AddItem El(-4).Eigenschaft
                    Else
                        Auswahl.AddItem Sys(1).S(4) 'ls
                    End If
                    If Sys(1).S(3) = "" Then
                        Auswahl.AddItem El(-3).Eigenschaft
                    Else
                        Auswahl.AddItem Sys(1).S(3) 'ts
                    End If
                Case Else
                    P = 0
                    Do
                        P = P + 1
                            If Kst(P).zuEigenschaft = Eiged.TextMatrix(Eiged.Row, 3) Then Auswahl.AddItem Kst(P).Bezeichnung
                    Loop Until Kst(P).zuEigenschaft = 0
            End Select
            For P = 0 To Auswahl.ListCount - 1
                i = B_Rex.TextWidth(Auswahl.List(P))
                If i > j Then j = i * 1.1 'den längsten text einsammeln
            Next P
            
            Auswahl.Top = Eiged.CellTop + Eiged.Top - 50
            Auswahl.Width = Eiged.CellWidth + 20
            Auswahl.left = Eiged.CellLeft + Eiged.left - 20
            If j > Auswahl.Width Then
                Auswahl.left = Auswahl.left - (j - Auswahl.Width)
                Auswahl.Width = j
            End If
            Auswahl.Visible = True
        End If
    End If
    
End Sub
'ende der eigenschaftsüberwachung


Private Sub Defaultbutton_Click()
    'eigenschaftseingabe
    If Eingabe.Visible = True Then
        Call Eingabe_KeyPress(13)
        Exit Sub
    End If
    If Auswahl.Visible = True Then
        Call Auswahl_KeyPress(13)
        Exit Sub
    End If
    
    'trumlängeneingabe
    'die trumlänge wird bei beiden beteiligten elementen eingetragen
    If Sys(E3).Verb(1, 1) = E4 And Sys(E3).Verb(1, 2) = EA3 Then
        Sys(E3).Verb(1, 3) = Val(Trumlänge)
    Else
        Sys(E3).Verb(2, 3) = Val(Trumlänge)
    End If
    If Sys(E4).Verb(1, 1) = E3 And Sys(E4).Verb(1, 2) = EA4 Then
        Sys(E4).Verb(1, 3) = Val(Trumlänge)
    Else
        Sys(E4).Verb(2, 3) = Val(Trumlänge)
    End If
    E3 = 0
    E4 = 0
    Trumlänge.Visible = False
    Trumlängeneinheit.Visible = False
    Beispielanlagen.Visible = True
    ModusCalc = ""
    Call CodeDraw.Alleelementeverbinden
    Call Eigschaftsverr.Bandmindestlängenberechnung(24)
    Call TabelleEig_ausfuellen
    Call CodeCalc.Rechnungssteuerung("C")
    Call Dateiverwaltung.Undo(0)
End Sub

Private Sub Elementleiste_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Trägergrkl = True Then Call Konstruktion_dragdrop(Element(0), 1, 1) 'unsinnige übergabe, will er aber so, es wird ohnehin nur die länge verändert
End Sub
Public Sub Datei_Click(Index As Integer)
Dim i, K, j, Indize, Merk, Dumm As Integer
Dim Rev As Boolean
Dim Dummy$, B_Rex_Version$

    On Error GoTo Errorhandler
    Select Case Index
        Case 0 'datei löschen
            If Gespeichert = False Then
                Dummy$ = MsgBox(Lang_Res(149), vbYesNo + vbQuestion + vbDefaultButton2)  'Die aktuelle Anlage wird gelöscht. Wollen Sie trotzdem eine neue Anlage konstruieren?
                If Dummy$ = vbNo Then Exit Sub
            End If
            Dateioffen = ""
            For i = 0 To 20 'userdaten löschen
                Sam(i) = ""
            Next i
            
            GrKl(0).Visible = False
            GrKl(1).Visible = False
            For i = 0 To Maxelementezahl 'sicher ist sicher
                Sys(i) = Del 'falls jemand sys(0) vehunzt hat, das eigentlich hierfür ist
            Next i
            Konstruktion.Cls
            
            Aktuel = False
            Gespeichert = True
            
            Sys(1).Element = "Band" 'damits aufgerufen wird, ob was drinsteht oder nicht
            Sys(2).Element = "Band" 'damits mit gespeichert wird
            Sys(1).Tag = "301" 'damits aufgerufen wird, ob was drinsteht oder nicht
            
            Lastaktel = 0
            Call Tabelle_ausfuellen(0)
            
            Exit Sub
        Case 1 'laden
            i = 9
            Do
                i = i + 1
                If Sys(i).Element <> "" Then
                    If Gespeichert = False Then
                        Dummy$ = MsgBox(Lang_Res(150), vbYesNo + vbQuestion + vbDefaultButton2)  'Die aktuelle Anlage wird gelöscht. Laden trotzdem durchführen?
                        If Dummy$ = vbNo Then Exit Sub
                    End If
                    Gespeichert = True
                    Call Datei_Click(0) 'rekursiver aufruf
                End If
            Loop Until i = Maxelementindex Or Sys(i).Element <> ""
            
            Maxelementindex = 10
            Merk = FreeFile

            Input #Merk, B_Rex_Version$
            If B_Rex_Version$ = "2.1" Then
                
                'tabelle leeren
                Lastaktel = 0
                Call Tabelle_ausfuellen(0)
                
                Input #Merk, Indize 'anzahl zu erwartender elemente
                For K = 1 To Indize
                    Input #Merk, i
                    Input #Merk, Sys(i).Height
                    Input #Merk, Sys(i).Width
                    Input #Merk, Sys(i).Top
                    Input #Merk, Sys(i).left
                    Input #Merk, Sys(i).Element 'kann später ersetzt werden, wird e bei sprachsteuerung überschrieben
                    Input #Merk, Sys(i).Tag
                    Input #Merk, Sys(i).Vollstaendig
                    Input #Merk, Sys(i).Zugehoerigkeit
                    Input #Merk, Sys(i).Verb(1, 1) 'wird nicht richtig übertragen
                    Input #Merk, Sys(i).Verb(2, 1) 'wird nicht richtig übertragen
                    Input #Merk, Sys(i).Verb(1, 3) 'wird nicht richtig übertragen
                    Input #Merk, Sys(i).Verb(2, 3) 'wird nicht richtig übertragen
                    For j = 1 To 100
                        Input #Merk, Sys(i).E(j)
                    Next j
                    For j = 1 To 10
                        Input #Merk, Sys(i).S(j) 'texteigenschaften
                    Next j
                    If Maxelementindex < i Then Maxelementindex = i
                Next K
                For i = 0 To 20
                    Input #Merk, Sam(i)
                Next i
                
                'nicht bei der neueren variante
                'wird nicht weiter verwendet
                If B_Rex_Version$ = "2.0" Then
                    For i = 1 To 2
                        For j = 0 To 200
                            Input #Merk, Dumm
                        Next j
                    Next i
                End If
                
                Input #Merk, Aktuel
                Input #Merk, Rev
                If Rev <> Reversieren Then
                    Call Datei_Click(9)
                End If
                
                'sys(x).B!! enthält zu jeder eigenschaft einen boolean-ausdruck
                If B_Rex_Version$ <> "2.0" Then 'faktisch nur höhere versionen
                    For i = 1 To Maxelementindex
                        If Sys(i).Element <> "" Then
                            For j = -TextEigenschaftszahl To Eigenschaftszahl
                                Input #Merk, Sys(i).B(j)
                                'If Sys(I).B(J) = True Then Stop'testen
                            Next j
                        End If
                    Next i
                End If
                Close Merk
                Call Bandmindestlängenberechnung(0) 'zur sicherheit
    
                Call CodeCalc.Rechnungssteuerung("E") 'um endlos = true rauszufinden
                'zeichnen und laufrichtung rausfinden
                Call CodeDraw.Alleelementeverbinden 'richtig zeichnen und träger auf rechts = true oder false stellen
                'mit richtiger laufrichtung rechnen
                Call CodeCalc.Rechnungssteuerung("VC") 'die fehlenden teile nachholen
                'muß zwar zweimal rechnen, dauert aber nicht lange und ist praktisch alles unter einem dach bei der Vollstaendigkeitskontrolle
                
                Lastaktel = 0
                Call Tabelle_ausfuellen(0)
                Call Code1.B_Rex_Uebersetzen
                Gespeichert = True

                SaveSetting Init_SettingDir, "Startup", "Directory", CurDir() 'immer mit letztem verzeichnis öffnen, auch nach neustart
                Call Dateiverwaltung.Undo(0) 'schnell noch auf die neue tour mitprotokollieren, damit das undo nicht zickt

            Else
                'ist ein anderes dateiformat
                'entweder hoffnungslos alt
                If left(B_Rex_Version$, 1) = "1" Then
                    Beep
                    Mother.H = Lang_Res(410)  'Datei kann nicht gelesen werden, falsche Version
                    Exit Sub
                End If
                
                'oder was viel besseres
                Close Merk 'wieder zu, es kommt ein neuer anlauf
            End If


        Case 2 'speichern
            
        
        Case 3 'b_rex_Druck
            'sonst wird nicht alles gezeichnet (markierung des bandes ausschalten)
            ModusCalc = ""
            E3 = 0
            E4 = 0

            If Druckerfehlt = True Then
                Mother.H = Lang_Res(152) 'Sie haben keinen Drucker eingerichtet
                Exit Sub
            End If
            Dummy$ = MsgBox(Lang_Res(3004), vbYesNo + vbQuestion + vbDefaultButton1)  'Bitte bestätigen Sie den Druck. Drucken?
            If Dummy$ = vbNo Then Exit Sub
            
            LetztesFenster = 3
            Mother.SeitenZahl.Visible = False
            
        Case 4 'alle typen auslegen
            Call CodeCalc.Rechnungssteuerung("E")
            If Endlos = False Then
                Mother.H = "die Anlage ist nicht endlos"
                Exit Sub
            End If
            Call CodeCalc.Rechnungssteuerung("V") 'wobei man hier noch zwischen band und anderen unterscheiden könnte
            If Vollstaendig = False Then
                Mother.H = "vervollstaendigen Sie bitte erst die Angaben zur Anlage"
                Exit Sub
            End If
            
            Anhalten = False 'mit esc taste wird der vorgang beendet
            Eingabe.Visible = False
            Auswahl.Visible = False
            i = Typ.AbsolutePosition

            Mother.H = "Stop = [Esc]"
            B_Rex_AutoLauf = True
            Anhalten = False
            
            'alte sichern und n paar sachen verbiegen
                Call Dateiverwaltung.Undo(0)
                Sys(1).E(1) = 0
                Sys(1).E(54) = Sys(1).E(54) 'spannkraft auf scheibe
                Sys(1).E(55) = Sys(1).E(55) 'gewicht an scheibe
            
            
            On Local Error Resume Next

            DoEvents
            Typ.MoveFirst


            Mother.H = ""
            B_Rex_AutoLauf = False
            Typ.AbsolutePosition = i 'den anfangswert wieder einstellen
            Call Dateiverwaltung.Undo(1)

        
        Case 5 To 8 'fensterverwaltung
            
            'typentabelle und anlagenzeichnung schliessen sich einander aus:
            
            Call Mother.Knopfverwaltung(Index, "GrosserKnopf", "Button", "B_Rex")
            
            Call Eigenschaftsleiste_Resize
            Call Markieren
        Case 9 'reversieren
            B_Rex.Konstruktion.Cls
            Call Mother.Knopfverwaltung(Index, "GrosserKnopf", "Button", "B_Rex")
            If left(Datei(9).Tag, 1) = "E" Then
                Reversieren = True
            Else
                Reversieren = False
            End If
            
            Call Dateiverwaltung.Undo(0)
            Call CodeDraw.Alleelementeverbinden
            
            Call CodeCalc.Rechnungssteuerung("C")
            
            Call TabelleEig_ausfuellen
        Case 10, 11
            Call Dateiverwaltung.Undo(Index - 9)
    End Select
    
    Exit Sub

Errorhandler:
    Mother.H = Lang_Res(413)  'abgebrochen / keine gültige Datei
    Call CodeDraw.Alleelementeverbinden
    Exit Sub
End Sub
Private Sub Element_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Shape1.Visible = False
    NeuEl = 0
End Sub
Private Sub Element_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'wird nur gebraucht, wenn das erste mal ein element eingerichtet wird
    X3 = x / Screen.TwipsPerPixelX
    Y3 = y / Screen.TwipsPerPixelY
    NeuEl = Index
    AktEl = 0
    Element(0).Top = Element(NeuEl).Top + Vertikal.Value '- (32 - element(aktel).Height)
    Element(0).left = Element(NeuEl).left + Horizontal.Value 'rechnersiche korrektur, als elementleiste links war
    Element(0).Height = Element(NeuEl).Height - 2
    Element(0).Width = Element(NeuEl).Width - 2 '?ist eben so
    Element(0).Drag vbBeginDrag
    ModusCalc = ""
    
End Sub
Private Sub element_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Element(0).Drag vbEndDrag
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Trägergrkl = True Then Call Konstruktion_dragdrop(Element(0), 1, 1) 'unsinnige übergabe, will er aber so, es wird ohnehin nur die länge verändert
End Sub

Private Sub GrKl_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    'für den seltenen fall, daß ein element genau auf eines der beiden markierungsquadrate fällt
    Call Konstruktion_dragdrop(Source, GrKl(Index).left + x / Screen.TwipsPerPixelX, GrKl(Index).Top + y / Screen.TwipsPerPixelY) 'unsinnige übergabe, will er aber so, es wird ohnehin nur die länge verändert
End Sub
Private Sub GrKl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If left(Sys(AktEl).Tag, 1) <> "1" Then Exit Sub 'ist kein förderer
    GrKl(Index).Drag vbBeginDrag
    Trägergrkl = True
    If Index = 0 Then 'vergrössern oder verkleinern
        Mark = True 'als boolean merken, ob rechts oder links verändert wird
    Else
        Mark = False
    End If
    Rahmen.Top = Sys(AktEl).Top
    Rahmen.left = Sys(AktEl).left
    Rahmen.Width = GrKl(1).left - GrKl(0).left
    Rahmen.Height = Sys(AktEl).Height
    Rahmen.Visible = True
End Sub
Private Sub GrKl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If left(Sys(AktEl).Tag, 1) = "1" Then
        GrKl(Index).MousePointer = 9
    Else
        GrKl(Index).MousePointer = 0
    End If
End Sub


Private Sub Konstruktion_DblClick()
    If E3 > 0 Then AktEl = 1 'damit die eigenschaften des bandes dargestellt werden
    If AktEl = 0 Then
        If left(Datei(6).Tag, 1) = "E" Then Call Datei_Click(6) 'eigenschaftsliste raus
        If left(Datei(7).Tag, 1) = "E" Then Call Datei_Click(7) 'Umfangskraftkurve raus
        If left(Datei(8).Tag, 1) = "E" Then Call Datei_Click(8) 'fehlerliste raus
    Else
        If left(Datei(6).Tag, 1) = "A" Then Call Datei_Click(6)
    End If
End Sub
Private Sub Konstruktion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If ModusCalc = "bandlöschen" Then Call Steuerung_Click(0)
        If AktEl = 1 Then Call Steuerung_Click(0)
     
        If ModusCalc = "elementlöschen" Then Call Steuerung_Click(3)
        If ModusCalc = "" And AktEl > 9 Then Call Steuerung_Click(3)
    End If
    ModusCalc = ""
    LastIndex = 0
    
End Sub
Private Sub Konstruktion_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    'falls er was angefangen hat, dann sichern
    If Eingabe.Visible = True Then Call Eingabe_Wertuebernehmen
    If Auswahl.Visible = True Then Call Auswahl_Click
    Eingabe.Visible = False
    Auswahl.Visible = False

    Mother.H = ""
    Abbruch = False 'ist noch unentschieden, obs das element schon gibt
    AktEl = 9 'ab 10 beginnen die elemente aus fleisch und blut
    Do 'rausfinden, wo hingeklickt wurde und gegf. drag auslösen
        AktEl = AktEl + 1
        i = Sys(AktEl).Height
        If left(Sys(AktEl).Tag, 1) = "1" Then i = 18
        If x > Sys(AktEl).left And x < Sys(AktEl).left + Sys(AktEl).Width And y > Sys(AktEl).Top And y < Sys(AktEl).Top + i Then
            'element wurde gefunden
            If ModusCalc = "elementlöschen" Then
                'AktEl = AktEl
                Call Steuerung_Click(3) 'element löschen
                Exit Sub
            End If
            If E3 > 0 Then 'es gibt einen markierten bandabschnitt, der kann wieder normal dargestellt werden
                ModusCalc = ""
                
                
                Trumlänge.Visible = False
                Trumlängeneinheit.Visible = False
                Beispielanlagen.Visible = True
                E3 = 0 'diese 4 zeilen vielleicht mal als unterprogramm
                E4 = 0
                
                Call CodeDraw.Alleelementeverbinden
            End If
            If Button = 2 Then ModusCalc = "bandauflegen" 'rechte maustaste
            If ModusCalc = "bandauflegen" And left(Sys(AktEl).Tag, 1) <> "2" Then 'kein transportgut
                If (Sys(AktEl).Verb(1, 1) = 0 Or Sys(AktEl).Verb(2, 1) = 0) Then  'noch mind. eine Verbindung frei
                    'Markiert = AktEl
                    Line1.X1 = Sys(AktEl).left + Sys(AktEl).Width / 2
                    Line1.Y1 = Sys(AktEl).Top + Sys(AktEl).Height / 2
                    Line1.X2 = x
                    Line1.Y2 = y
                    Line1.Visible = True
                End If
                Exit Sub
            Else
                If ModusCalc = "bandauflegen" Then Exit Sub 'empirisch gefunden
                Element(0).Top = Sys(AktEl).Top
                If Sys(AktEl).Tag = "201" Then Element(0).Top = Element(0).Top + 15
                Element(0).left = Sys(AktEl).left
                If left(Sys(AktEl).Tag, 1) = "1" Then
                    Element(0).Height = 20 'elementhöhe entspricht nicht der systemhöhe
                Else
                    Element(0).Height = Sys(AktEl).Height
                End If
                Element(0).Width = Sys(AktEl).Width
                X3 = x - Sys(AktEl).left 'position innerhalb des elements
                Y3 = y - Sys(AktEl).Top
                X1 = Sys(AktEl).left 'merken, um falsche Positionierung aufzuheben
                Y1 = Sys(AktEl).Top
                ModusCalc = ""
                LastIndex = 1
                
                If Trägergrkl = False Then Element(0).Drag vbBeginDrag
                Abbruch = True 'entscheidet, ob neu erstellt, oder nur ersetzt wird
                'Markiert = AktEl
            End If
        End If
    Loop Until AktEl = Maxelementindex + 1 Or Abbruch = True
    If AktEl > Maxelementindex Then AktEl = 0 'wurde nix markiert
    
    'falls er was angefangen hat, dann sichern
    If Eingabe.Visible = True Then Call Eingabe_Wertuebernehmen
    If Auswahl.Visible = True Then Call Auswahl_Click

    Call Markieren 'gegebenenfalls wird auch die tabelle ausgefüllt
    
End Sub
Private Sub Markieren()
    'falls ein element bloss markiert werden sollte
    If AktEl > 9 Then 'And Left(Sys(AktEl).Tag, 1) = "1" Then'alle können markiert werden
        Select Case left(Sys(AktEl).Tag, 1)
            Case "1" 'träger
                GrKl(0).Top = Sys(AktEl).Top + 7
                GrKl(1).Top = Sys(AktEl).Top + 7
            Case "2"
                Select Case Sys(AktEl).Tag
                    Case "201" 'transportgut
                        GrKl(0).Top = Sys(AktEl).Top + 30
                        GrKl(1).Top = Sys(AktEl).Top + 30
                    Case Else
                        GrKl(0).Top = Sys(AktEl).Top + 13
                        GrKl(1).Top = Sys(AktEl).Top + 13
                End Select
            Case Else 'einzelne
                GrKl(0).Top = Sys(AktEl).Top + Sys(AktEl).Height / 2 - 4
                GrKl(1).Top = Sys(AktEl).Top + Sys(AktEl).Height / 2 - 4
        End Select
        GrKl(0).left = Sys(AktEl).left
        GrKl(1).left = Sys(AktEl).left + Sys(AktEl).Width - GrKl(1).Width
        GrKl(0).Visible = True
        GrKl(1).Visible = True
    Else
        GrKl(0).Visible = False
        GrKl(1).Visible = False
    End If
    Call Tabelle_ausfuellen(AktEl)
End Sub
Private Sub Konstruktion_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim L As Integer
    
    
    If ModusCalc = "bandauflegen" Then
        L = 9
        Do 'rausfinden, worüber die maus sich bewegt und entsprechenden Mauszeiger anschalten
            L = L + 1
            If x > Sys(L).left And x < Sys(L).left + Sys(L).Width And y > Sys(L).Top And y < Sys(L).Top + Sys(L).Height Then
                If left(Sys(L).Tag, 1) <> "2" And L <> AktEl Then
                    If Sys(L).Verb(1, 1) = 0 Or Sys(L).Verb(2, 1) = 0 Then 'noch mind. eine Verbindung frei
                        If B_Rex.MousePointer <> 14 Then
                            B_Rex.MousePointer = 14 'pfeil und fragezeichen
                            If Abbruch = False Then Mother.H = Lang_Res(153) 'dieses Element kann die Bandverbindung aufnehmen
                        End If
                        Line1.X2 = x
                        Line1.Y2 = y
                        L = Maxelementindex
                    End If
                End If
                If L = AktEl Then
                    Mother.H = ""
                    Line1.X2 = x
                    Line1.Y2 = y
                    Exit Sub
                End If
            Else
                If B_Rex.MousePointer <> 1 Then B_Rex.MousePointer = 1
            End If
        Loop Until L = Maxelementindex
    End If
    'L = 0
    If AktEl = 0 And ModusCalc = "bandauflegen" Then
        If left(Mother.H, 2) <> "kl" Then Mother.H = Lang_Res(154) 'klicken Sie nacheinander die zu verbindenden Elemente an
    End If
    
    'immer noch das ende der dicken grünen verbindungslinie neu ausrichten
    If ModusCalc = "bandauflegen" And AktEl > 0 Then
        Line1.X2 = x
        Line1.Y2 = y
    End If
End Sub
Private Sub Konstruktion_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim L As Integer
    Dim K As Integer
    Dim i As Integer
    Dim P As Integer
    L = 9
    Do 'rausfinden, auf welchem element fallengelassen wurde
        L = L + 1
        If x > Sys(L).left And x < Sys(L).left + Sys(L).Width And y > Sys(L).Top And y < Sys(L).Top + Sys(L).Height Then
            'auf einem element, also nur zum bandauflegen
            If AktEl = 0 Then Exit Sub
            If left(Sys(L).Tag, 1) <> "2" And L <> AktEl Then 'wurde auf einem freien steckplatz an einem zugelassenen element losgelassen
                If Sys(L).Verb(1, 1) = 0 Or Sys(L).Verb(2, 1) = 0 Then 'noch mind. eine Verbindung frei
                    
                    'schauen, ob bei geschloss. anlage denn auch eine antriebsscheibe drin ist
                    If (Sys(AktEl).Verb(1, 1) > 0 Or Sys(AktEl).Verb(2, 1) > 0) And (Sys(L).Verb(1, 1) > 0 Or Sys(L).Verb(2, 1) > 0) Then 'kreislauf wird ev. geschlossen
                        Abbruch = True
                        If Sys(L).Tag = "001" Then Abbruch = False 'würde als erstes element sonst nicht erfaßt
                        K = L 'kann auch irgendeine sein, aber hier ist bestimmt n element drin
                        If Sys(K).Verb(1, 1) > 0 Then
                            P = 1
                        Else
                            P = 2
                        End If
                        Do
                            i = K ' altes element merken
                            K = Sys(K).Verb(P, 1) 'und nächstes bestimmen
                            If Sys(K).Verb(1, 1) = i Then 'voreinstellungen für neuen durchlauf
                                P = 2
                            Else
                                P = 1
                            End If
                            If Sys(K).Tag = "001" Then Abbruch = False 'es ist eine prim. Antriebsscheibe im Kreislauf
                        Loop Until K = AktEl Or Sys(K).Verb(1, 1) = 0 Or Sys(K).Verb(2, 1) = 0
                        
                        If K = AktEl And Abbruch = False Then Endlos = True
                        If K = AktEl And Abbruch = True Then 'ist einmal rum (wäre tats. geschl. Kreislauf), hat aber keine Antriebsscheibe gefunden
                            Mother.H = Lang_Res(425)  'Der Kreislauf muß eine prim. Antriebsscheibe enthalten.
                            Exit Sub
                        End If
                    End If

                    'na schön, ist eine drin, kann weiter gehen
                    If Sys(AktEl).Verb(1, 1) = 0 Then
                        Sys(AktEl).Verb(1, 1) = L
                    Else
                        Sys(AktEl).Verb(2, 1) = L
                    End If
                    If Sys(L).Verb(1, 1) = 0 Then
                        Sys(L).Verb(1, 1) = AktEl
                    Else
                        Sys(L).Verb(2, 1) = AktEl
                    End If
                    Call CodeDraw.Alleelementeverbinden
                    L = Maxelementindex
                    'AktEl = 0
                    Line1.Visible = False
                    
                    If Sys(AktEl).E(41) = 0 Then Sys(AktEl).E(41) = 1
                    If Sys(AktEl).E(42) = 0 Then Sys(AktEl).E(42) = 2
                    If Sys(L).E(41) = 0 Then Sys(L).E(41) = 1
                    If Sys(L).E(42) = 0 Then Sys(L).E(42) = 2
                    
                    If Button = 2 Then
                        ModusCalc = ""
                        MousePointer = 1
                    End If
                    
                    'so, als wäre breite eingegeben worden
                    Call Eigschaftsverr.Verrechnung(1, 34, Sys(1).E(34), Sys(1).E(34))
                    
                    Call Eigschaftsverr.Bandmindestlängenberechnung(24) 'direkt ohne int_ext_abgleich
                    Call TabelleEig_ausfuellen
                    
                End If
            End If
            Exit Sub
        End If
    Loop Until L = Maxelementindex
    
    'und nachschauen, ob ein Band getroffen wurde, dann gleich markieren
    Line1.Visible = False 'alten moduscalc beenden
    'AktEl = 0
    
    If E3 > 0 Then 'wurde schon ein Band markiert
        Trumlänge.Visible = False
        Trumlängeneinheit.Visible = False
        Beispielanlagen.Visible = True
        E3 = 0
        E4 = 0
        EA3 = 0
        EA4 = 0
        ModusCalc = ""
        
        Call CodeDraw.Alleelementeverbinden 'also erst mal ursprungszustand wieder herstellen
    End If
    
    If ModusCalc = "bandlöschen" Then
        ModusCalc = "liniensuchen"
        Call CodeDraw.Alleelementeverbinden(x, y)
        ModusCalc = "bandlöschen" 'moduscalc würde sonst verloren gehen, aber vielleicht soll noch ein zweites band gelöscht werden
    Else
        ModusCalc = "liniensuchen" 'beim bandlöschen kann die gefundene verbindung gleich gelöscht werden
        Call CodeDraw.Alleelementeverbinden(x, y)
    End If
    
    If E3 > 0 Then
        AktEl = 1 'bandelement wurde markiert
        Call Markieren 'gegebenenfalls wird auch die tabelle ausgefüllt
    End If
    
    
    If E3 = 0 Then 'hat keine linie erwischt
        ModusCalc = ""
        B_Rex.MousePointer = 1
        
        Line1.Visible = False
        LastIndex = 0
        'AktEl = 0
    Else 'band erwischt
        'AktEl = 0
        If Zweischeiben = False Then
            Trumlänge.Visible = True
            Trumlänge = ""
            Trumlängeneinheit.Visible = True
            Beispielanlagen.Visible = False
            If Sys(E3).Verb(1, 1) = E4 And Sys(E3).Verb(1, 2) = EA3 Then
                If Sys(E3).Verb(1, 3) > 0 Then Trumlänge = Sys(E3).Verb(1, 3)
            Else
                'If Sys(E3).Verb(2, 1) = E4 And Sys(E3).Verb(2, 2) = EA3 Then
                    If Sys(E3).Verb(2, 3) > 0 Then Trumlänge = Sys(E3).Verb(2, 3)
                'End If
            End If
            Trumlänge.SetFocus
            
            Trumlänge.SelStart = 0 'damit man's gleich neu eintippen kann ohne zu löschen
            Trumlänge.SelLength = Len(Trumlänge)
        End If
        If Zweischeiben = True Then Mother.H = Lang_Res(155)  'Bei 2 Scheiben können geben Sie statt der Trumlänge bitte den Achsabstand ein.
        If ModusCalc = "bandlöschen" Then
            Call Steuerung_Click(0)
        End If
    End If
    

End Sub
Private Sub Konstruktion_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    'wenn ein neues oder altes element über der oberfläche bewegt wird
    If DragDrop = True Then
        DragDrop = False
        Auflegedehnung.Drag vbEndDrag
        Auflegedehnunganzeiger.Visible = False
        Exit Sub 'weil nämlich dann die auflegedehung bewegt wird
    End If
     
    If Trägergrkl = True Then 'element vergrößern oder verkleinern
        If Mark = True Then 'links wird bewegt
            x = CInt(x / 32) * 32
            If x > Sys(AktEl).left + Sys(AktEl).Width - 64 Then x = Sys(AktEl).left + Sys(AktEl).Width - 64
            GrKl(0).left = x
            Rahmen.Width = Sys(AktEl).Width - (x - Sys(AktEl).left)
            Rahmen.left = x
        Else 'rechts wird bewegt
            x = CInt(x / 32) * 32
            If x < Sys(AktEl).left + 64 Then x = Sys(AktEl).left + 64
            GrKl(1).left = x - GrKl(1).Width
            Rahmen.Width = x - Sys(AktEl).left
        End If
    Else 'das element(0) verschieben
        If AktEl > 0 Then
            'damit nicht beim blossen draufdrücken schon was passiert
            If Abs(Sys(AktEl).left + Sys(AktEl).Width / 2 - x) < 15 And Abs(Sys(AktEl).Top + Sys(AktEl).Height / 2 - y) < 30 Then Exit Sub 'hakelt leider n bisschen                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              'nicht jedesmal anzeigen, sonst zittert das bild, sollte unsinnige operationen verhindern
            Shape1.Width = Sys(AktEl).Width
            Shape1.Height = Sys(AktEl).Height
            If left(Sys(AktEl).Tag, 1) = "1" Then Shape1.Height = 16
            Shape1.Tag = Sys(AktEl).Tag
        Else
            Shape1.Width = Element(NeuEl).Width 'wenn' n' neues ist
            Shape1.Height = Element(NeuEl).Height
            Shape1.Tag = Element(NeuEl).Tag
        End If
        X2 = x
        Y2 = y
        Call Positionierung(Shape1.Width, Shape1.Tag)
        'hier wird x2 und y2 auch gerastert
        
        'IPR = X2
        'JPR = Y2
        If Abbruch = False Then Mother.H = ""
        If Int(Shape1.Top) <> Y2 Or Int(Shape1.left) <> X2 Then
            Shape1.Top = Y2
            If Sys(AktEl).Tag = "201" Then Shape1.Top = Y2 + 14
            Shape1.left = X2
            If Abbruch = True Then
                Shape1.BackColor = QBColor(12)
            Else
                Shape1.BackColor = QBColor(10)
            End If
            Shape1.Visible = True
        End If
    End If
End Sub
Public Sub Konstruktion_dragdrop(Source As Control, x As Single, y As Single)
    Dim i As Integer, j As Integer, Merk As Integer
    Dim Masse, Masse1 As Double
    
    Shape1.Visible = False
    Shape1.Tag = ""
    
    'nur im moduscalc vergrößern/verkleinern
    If Trägergrkl = True Then
        Call GrößerKleiner 'förderer vergrössern/verkleinern
        Exit Sub
    End If
    DoEvents
    X2 = x
    Y2 = y
    
    Abbruch = False
    If NeuEl > 0 Then   'element wird neu erstellt
        AktEl = 0
        Call Positionierung(Element(NeuEl).Width, Element(NeuEl).Tag)
        'Abbruch = False'nur zum Ausprobieren Grafik
        If Abbruch = True Then
            NeuEl = 0
            Exit Sub
        End If
        If Element(NeuEl).Tag = "001" Then 'primäre Antriebsscheibe wird nur einmal zugelassen
            For i = 1 To Maxelementindex
                If Sys(i).Tag = "001" Then
                    NeuEl = 0
                    Mother.H = Lang_Res(431)  'Es kann nur eine primäre Antriebsscheibe geben.
                    Exit Sub
                End If
            Next i
        End If
        
        i = 9
        Do
            i = i + 1
        Loop Until Sys(i).Element = ""
        If i > Maxelementindex Then Maxelementindex = i 'maximalen Index für spätere Schleifen merken
        If left(Element(NeuEl).Tag, 1) = "1" Then
            Sys(i).Width = 128
            Sys(i).Height = 32
        Else
            Sys(i).Width = 32
            Sys(i).Height = 32
        End If
        Sys(i).Element = Element(NeuEl).ToolTipText
        Sys(i).Tag = Element(NeuEl).Tag
        Sys(i).Top = Y2
        Sys(i).left = X2
        
        'grundeinstellungen in listen vornehmen einschieben, damit z.b. sofort durchbiegung gerechnet wird
            j = Elementnummer(Sys(i).Tag)
            If El(48).Eig(j) <> "" Then Sys(i).E(48) = 18 'messerkantenmaterial, gibt nur stahl
            If El(15).Eig(j) <> "" Then Sys(i).E(15) = 2 'Tischoberfläche, voreingestellt melaminharz
            If El(14).Eig(j) <> "" Then Sys(i).E(14) = 1 'scheibenoberflächen, voreingestellt stahl
            If El(60).Eig(j) <> "" Then Sys(i).E(60) = 22 'aut. überlast, voreingestellt gleichmäßiger betrieb, 0%
            If El(47).Eig(j) <> "" Then
                Sys(i).E(47) = 104 'material der scheibe voreingestellt stahl
                Sys(i).E(103) = 109 'gehört immer ins schlepp, wird selbst nicht angesprochen, e-modul
            End If
            If El(36).Eig(j) <> "" Then
                Sys(i).E(36) = 15 'transportgutart, voreingestellt pappkarton
                Sys(i).E(61) = 85 'gehört immer ins schlepp, wird selbst nicht angesprochen
            End If
            
        If Träger > 0 Then Sys(i).Zugehoerigkeit = Träger
        AktEl = i 'zum Zeichnen
        NeuEl = 0
        Lastaktel = 0 'darstellung erzwingen
        Call Eigschaftsverr.Zwei_Scheiben 'variable zweischeiben richtig einstellen

    Else 'element wird nur in seiner Lage verändert
        
        'wurde nur rumgespielt, aber keine Lage verändert?
        If Abs(Sys(AktEl).Top + Y3 - y) < Rasterungy / 2 And Abs(Sys(AktEl).left + X3 - x) < Rasterungx / 2 Then Exit Sub
        
        Call Positionierung(Sys(AktEl).Width, Sys(AktEl).Tag)
        If Abbruch = True Then Exit Sub
    
        'wenns n huckepack ist
        'mal an die position setzen, damit folgendes funktioniert
        Mother.H = ""
        Merk = Sys(AktEl).Zugehoerigkeit 'alte Zugehoerigkeit merken
        Sys(AktEl).Zugehoerigkeit = Träger 'übergabe vorbereiten
        
        'ist die stelle auf dem träger noch frei?
        If Abbruch = False And left(Sys(AktEl).Tag, 1) = "2" Then
            Call Eigschaftsverr.Trägeraufteilung(AktEl) 'den neuen träger testen
        
            'hier quell- und zielträger auf massewahrheit prüfen
            Masse = 0
            Masse1 = 0
            For j = 1 To Maxelementindex
                If Sys(j).Tag = "201" And Sys(j).Zugehoerigkeit = Merk Then Masse = Masse + Sys(j).E(28)
                If Sys(j).Tag = "204" And Sys(j).Zugehoerigkeit = Merk Then Masse = Masse - Sys(j).E(23) 'andere staumassen abziehen
                If Sys(j).Tag = "201" And Sys(j).Zugehoerigkeit = Träger Then Masse1 = Masse1 + Sys(j).E(28)
                If Sys(j).Tag = "204" And Sys(j).Zugehoerigkeit = Träger Then Masse1 = Masse1 - Sys(j).E(23) 'andere staumassen abziehen
            Next j

            If Masse < 0 Or Masse1 < 0 Then
                Mother.H = Lang_Res(433) & j & Lang_Res(434)   'förderer (nr), Stau- und Transportmassen würden nicht harmonieren
                Beep
                Abbruch = True
            End If
        
        End If

        If Abbruch = False Then
            Sys(AktEl).Top = Y2
            Sys(AktEl).left = X2
            
            If left(Sys(AktEl).Tag, 1) = "1" Then 'wenn's n träger ist, muß er noch was mitnehmen
                For i = 1 To Maxelementindex 'verschieben genehmigt, Huckepackelemente mitnehmen
                    If Sys(i).Zugehoerigkeit = AktEl And AktEl > 0 Then
                        Sys(i).Top = Sys(i).Top + (Y2 - Y1)
                        Sys(i).left = Sys(i).left + (X2 - X1)
                    End If
                Next i
            End If
            
            'transsportgut wurde versetzt, neue streckenlast vergeben
            If Sys(AktEl).Tag = "201" Then
                Call Eigschaftsverr.Verrechnung(AktEl, 28, Sys(AktEl).E(28), Sys(AktEl).E(28))
                Call TabelleEig_ausfuellen
            End If
            
            'träger zuordnen (auch wenn's wieder derselbe ist)
            If Träger > 0 Then Sys(AktEl).Zugehoerigkeit = Träger
        Else
            Sys(AktEl).Zugehoerigkeit = Merk 'wieder zurückstellen
        End If
    End If
    'Call Vollstaendigkeitskontrolle
    Call CodeCalc.Rechnungssteuerung("EVC")
    Call CodeDraw.Alleelementeverbinden
    Call Markieren
    Call Dateiverwaltung.Undo(0) 'irgendwas hat sich sowieso verändert
    Aktuel = False
    Abbruch = False
    Konstruktion.SetFocus
End Sub
Private Sub Kopfleiste_DragDrop(Source As Control, x As Single, y As Single)
    Shape1.Visible = False
    NeuEl = 0
End Sub
Private Sub Kopfleiste_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Trägergrkl = True Then Call Konstruktion_dragdrop(Element(0), 1, 1) 'unsinnige übergabe, will er aber so, es wird ohnehin nur die länge verändert
End Sub
Private Sub Steuerung_Click(Index As Integer)
    Dim i As Integer
    Mother.H = ""
    GrKl(0).Visible = False
    GrKl(1).Visible = False
    Select Case Index
        Case 0 'bandlöschen
            If ModusCalc = "bandlöschen" Then
                If E3 = 0 Then
                    ModusCalc = ""
                    
                End If
            Else
                If E3 > 0 Then
                    'es soll einfach nur weitergehen
                Else
                    
                    E3 = 0
                    E4 = 0
                    
                    Trumlänge.Visible = False
                    Trumlängeneinheit.Visible = False
                    Beispielanlagen.Visible = True
                    ModusCalc = "bandlöschen"
                    Mother.H = Lang_Res(156)  '"Klicken Sie auf den zu löschenden Bandabschnitt"
                    Exit Sub
                End If
            End If
            
            'verbindung löschen
            If Sys(E3).Verb(1, 1) = E4 And Sys(E3).Verb(1, 2) = EA3 Then 'es muß auch die richtige verbindungsstelle sein
                Sys(E3).Verb(1, 1) = 0
                Sys(E3).Verb(1, 2) = 0
                Sys(E3).Verb(1, 3) = 0
            Else
                Sys(E3).Verb(2, 1) = 0
                Sys(E3).Verb(2, 2) = 0
                Sys(E3).Verb(2, 3) = 0
            End If
            If Sys(E4).Verb(1, 1) = E3 And Sys(E4).Verb(1, 2) = EA4 Then
                Sys(E4).Verb(1, 1) = 0
                Sys(E4).Verb(1, 2) = 0
                Sys(E4).Verb(1, 3) = 0
            Else
                Sys(E4).Verb(2, 1) = 0
                Sys(E4).Verb(2, 2) = 0 'muß sein, sonst unsauberkeiten
                Sys(E4).Verb(2, 3) = 0
            End If
            
            If Sys(E3).Verb(1, 1) = 0 And Sys(E3).Verb(2, 1) = 0 Then
                Sys(E3).E(41) = 0 'kein band dran
                Sys(E3).E(42) = 0
            End If
            If Sys(E4).Verb(1, 1) = 0 And Sys(E4).Verb(2, 1) = 0 Then
                Sys(E4).E(41) = 0
                Sys(E4).E(42) = 0
            End If
            Call CodeCalc.Rechnungssteuerung("E")

            Trumlänge.Visible = False
            Trumlängeneinheit.Visible = False
            Beispielanlagen.Visible = True
            E3 = 0
            E4 = 0
            
            If ModusCalc <> "bandlöschen" Then
                ModusCalc = "" 'bei "liniensuchen würde der bildschirm nicht gelöscht
            Else
                Mother.H = Lang_Res(156)  '"Klicken Sie auf den zu löschenden Bandabschnitt"
            End If
            Call CodeDraw.Alleelementeverbinden
            Call Dateiverwaltung.Undo(0)
            Aktuel = False

        Case 1 'bandauflegen
            Trumlänge.Visible = False
            Trumlängeneinheit.Visible = False
            Beispielanlagen.Visible = True
            E3 = 0
            E4 = 0
            
            If ModusCalc = "bandauflegen" Then
                ModusCalc = ""
                
            Else
                ModusCalc = "bandauflegen"
            End If
            Mother.H = Lang_Res(157)  'ziehen Sie bei gedrückter linker Maustaste eine Verbindung zw. den Elementen
        Case 3 'element löschen
            If AktEl = 1 Then Exit Sub
            If ModusCalc = "elementlöschen" And AktEl = 0 Then
                ModusCalc = ""
                
            Else
                ModusCalc = "elementlöschen"
            End If
            Mother.H = Lang_Res(154)  'klicken Sie nacheinander die zu verbindenden Elemente an
            If AktEl = 0 Then Exit Sub
            For i = 9 To Maxelementindex 'besetzte Träger dürfen nicht gelöscht werden
                If Sys(i).Zugehoerigkeit = AktEl Then 'ist besetzt, löschen verweigern
                    Mother.H = Lang_Res(418)  'Bitte entfernen Sie zunächst die aufgesetzten Elemente
                    Exit Sub
                End If
            Next i
            Call Markieren
            Dummy$ = MsgBox(Lang_Res(419), vbYesNo + vbQuestion + vbDefaultButton1)  'Wollen Sie das Element wirklich löschen?
            If Dummy$ = vbNo Then Exit Sub
            Sys(AktEl) = Del 'setzt alles zurück, solange niemand in die 0 was reinschreibt
            Eiged.Clear
            Lastaktel = 0 'sonst wird die eigenschaftstabelle leer aufgerufen
            Call Eigschaftsverr.Zwei_Scheiben 'variable zweischeiben in ordnung bringen
            
            'verbindungen mit dem zu löschenden element löschen
            For i = 1 To Maxelementindex
                If Sys(i).Verb(1, 1) = AktEl Then
                    Sys(i).Verb(1, 1) = 0
                    Sys(i).Verb(1, 2) = 0
                    Sys(i).Verb(1, 3) = 0
                End If
                If Sys(i).Verb(2, 1) = AktEl Then
                    Sys(i).Verb(2, 1) = 0
                    Sys(i).Verb(2, 2) = 0
                    Sys(i).Verb(2, 3) = 0
                End If
                If Sys(i).Verb(1, 1) = 0 And Sys(i).Verb(2, 1) = 0 Then
                    Sys(i).E(41) = 0
                    Sys(i).E(42) = 0
                    'Sys(I).Vollstaendig = False
                End If
            Next i
            
            'maxelementezahl zurücksetzen
            i = Maxelementezahl
            Do
                i = i - 1
            Loop Until i < 10 Or Sys(i).Element <> ""
            
            Aktuel = False
            If i < 10 Then 'ist kein element mehr da
                Gespeichert = True
                Maxelementindex = 20
            Else
                Call Dateiverwaltung.Undo(0)
                Maxelementindex = i + 1 'damit er für ein neues element einen neuen platz finden kann
            End If
            
            Call CodeCalc.Rechnungssteuerung("E")

            AktEl = 0
            Call Markieren 'markierungen rechts und links löschen
            Call CodeDraw.Alleelementeverbinden
        End Select
        
        Call Eigschaftsverr.Bandmindestlängenberechnung(24)
        Call CodeCalc.Rechnungssteuerung("EVC")

End Sub
Private Sub Steuerung_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Shape1.Visible = False
    NeuEl = 0
    If AktEl > 0 Then
        Call Steuerung_Click(3)
        ModusCalc = ""
        
    End If
End Sub
Private Sub GrößerKleiner()
    If left(Sys(AktEl).Tag, 1) <> "1" Then Exit Sub
    Mother.H = ""
    Abbruch = False
    oGrenze = Sys(AktEl).Top - 32
    uGrenze = Sys(AktEl).Top + 32
    lGrenze = GrKl(0).left - 32
    rGrenze = GrKl(1).left + GrKl(1).Width
    Trägergrkl = False 'dragmodus zurückstellen
  
    For i = 9 To Maxelementindex
        Call Kollisionspruefung(i)
        If Abbruch = False And Sys(i).Zugehoerigkeit = AktEl Then 'für den Fall, daß es kleiner werden soll, darf kein Huckepack in der Luft stehen
            GrKl(0).left = Sys(AktEl).left
            GrKl(1).left = Sys(AktEl).left + Sys(AktEl).Width - GrKl(1).Width
            Rahmen.Visible = False
            Mother.H = Lang_Res(439)  'entfernen Sie erst das rechts aufgesetzte Element
            Exit Sub
        End If
        Abbruch = False
    Next i
    rGrenze = GrKl(1).left + GrKl(1).Width + 32
    For i = 9 To Maxelementindex
        Call Kollisionspruefung(i)
        If Abbruch = True And Sys(i).Zugehoerigkeit <> AktEl Then 'für den Fall, daß es größer werden soll und was im Weg ist
            GrKl(0).left = Sys(AktEl).left
            GrKl(1).left = Sys(AktEl).left + Sys(AktEl).Width - GrKl(1).Width
            Rahmen.Visible = False
            Mother.H = Lang_Res(441) 'unzulässige Berührung/ Kollision
            Exit Sub
        End If
        Abbruch = False
    Next i
    
    Konstruktion.Line (Sys(AktEl).left, Sys(AktEl).Top)-(Sys(AktEl).left + Sys(AktEl).Width, Sys(AktEl).Top + 32), vbWhite, BF
    Sys(AktEl).left = CInt(GrKl(0).left / 32) * 32
    Sys(AktEl).Width = CInt((GrKl(1).left - GrKl(0).left) / 32) * 32
    Rahmen.Visible = False
    Call CodeDraw.Alleelementeverbinden
End Sub
Private Sub Positionierung(ByVal Ewidth As Long, ByVal Etag As String)
    'ETAG ist der .tag des zu zu verändernden / neuen Elements
    'aktel enthält Nummer bei zu veränderndem element, ist global
    'neuel die daten neuer element, neuel und aktel schliessen einander aus
    
    Dim i As Integer
    Dim j As Integer
   
    Träger = 0 'zurücksetzen
    Rasterungx = 32
    Rasterungy = 16
    
    'x3,y3 ist die Stellung des Zeigers innerhalb des Elements
    X2 = X2 - X3 'x2, y2 werden als top/left-Position zurückgegeben
    Y2 = Y2 - Y3 'die Ablage entspricht so dem DragDrop-Symbol
    X2 = CInt(X2 / Rasterungx) * Rasterungx 'horizontale Rasterung
    Y2 = CInt(Y2 / Rasterungy) * Rasterungy 'vertikale Rasterung
    
    'Kantenüberwachung
    Do While X2 <= 0 'element von der linken Kante fernhalten
        X2 = X2 + Rasterungx
    Loop
    Do While X2 + Ewidth + 10 >= Konstruktion.Width 'element von der rechten Kante fernhalten
        X2 = X2 - Rasterungx
    Loop
    Do While Y2 <= 0 'element von der oberen Kante fernhalten
        Y2 = Y2 + Rasterungy
    Loop
    Do While Y2 + Rasterungy >= Konstruktion.Height 'element von der unteren Kante fernhalten
        Y2 = Y2 - Rasterungy
    Loop
    
    'kollisionsprüfung
    Abbruch = False
    Select Case left(Etag, 1)
        Case "0" 'einzelstehend
            oGrenze = Y2 - Rasterungy + 1
            uGrenze = Y2 + 32 + Rasterungy
            lGrenze = X2 - Rasterungx
            rGrenze = X2 + Ewidth + Rasterungx
            For i = 9 To Maxelementindex
                Call Kollisionspruefung(i)
                If Abbruch = True Then
                    Mother.H = Lang_Res(442) 'element darf andere nicht berühren
                    Exit Sub
                End If
            Next i
        Case "1" 'förderer
            oGrenze = Y2 - 1 * Rasterungy
            For j = 9 To Maxelementindex
                If Sys(j).Zugehoerigkeit = AktEl And NeuEl = 0 Then oGrenze = Y2 - 3 * Rasterungy 'ist ein huckepack drauf
            Next j
            uGrenze = Y2 + 2 * Rasterungy
            lGrenze = X2 - Rasterungx
            rGrenze = X2 + Ewidth + Rasterungx
            For i = 9 To Maxelementindex
                If Sys(i).Zugehoerigkeit <> AktEl Or NeuEl > 0 Then Call Kollisionspruefung(i)
                If Abbruch = True Then
                    Mother.H = Lang_Res(441) 'unzulässige Berührung/Kollision
                    Exit Sub
                End If
            Next i
        Case "2" 'huckepack
            i = 9
            Do
                i = i + 1
                Abbruch = False
                If Sys(i).Element <> "" Then 'element muß es noch geben
                    If left(Sys(i).Tag, 1) = "2" Then 'andere Huckepacks dürfen angrenzen
                        oGrenze = Y2
                        uGrenze = Y2 + 32
                        lGrenze = X2
                        rGrenze = X2 + Ewidth
                        Call Kollisionspruefung(i)
                        If Abbruch = True Then
                            Mother.H = Lang_Res(441)  'unzulässige Berührung/Kollision
                            Exit Sub
                        End If
                    End If
                    If left(Sys(i).Tag, 1) = "1" Then 'Träger dürfen drunter sitzen
                        oGrenze = Y2 'da muß ein förderer drunter sitzen
                        uGrenze = Y2 + 64
                        lGrenze = X2 ' - Rasterungx
                        rGrenze = X2 + Ewidth ' + Rasterungx
                        Call Kollisionspruefung(i)
                        If Abbruch = True Then
                            If Sys(i).Tag = "104" And Etag <> 206 Then
                                Mother.H = Lang_Res(443)  'Dieser Förderer kann das Element rechnerisch nicht erfassen
                                Exit Sub
                            Else
                                Träger = i 'betreffenden Träger merken
                                Y2 = Sys(i).Top - 32 'über den Träger zwingen
                            End If
                        End If
                    End If
                    If left(Sys(i).Tag, 1) = "0" Then 'einzelstehende
                        oGrenze = Y2 - 32
                        uGrenze = Y2 + 32 + Rasterungy
                        lGrenze = X2 - Rasterungx
                        rGrenze = X2 + Ewidth + Rasterungx
                        Call Kollisionspruefung(i) 'E's sind ohne if inbegriffen
                        If Abbruch = True Then
                            Mother.H = Lang_Res(441)  'unzulässige Berührung/Kollision
                            Exit Sub
                        End If
                    End If
                End If
            Loop Until i = Maxelementindex
            If Träger = 0 Then  'wenn kein Förderer drunter ist
                Mother.H = Lang_Res(444)  'dieses Element benötigt einen Förderer, z.B. Rollenbahn
                Abbruch = True
                Exit Sub
            End If
            Abbruch = False
    End Select
    Abbruch = False
End Sub
Private Sub Kollisionspruefung(ByVal i As Integer) 'element(zeiger) ist zu prüfen, eine I-Schleife für jedes element ist von außen zu steuern
    Dim M As Integer
    Dim K As Integer
    If Sys(i).Tag <> "" And AktEl <> i Then 'zu überpr. Element muß besetzt sein, zu pos. Element darf nicht mit sich selbst verglichen werden
        For M = oGrenze To uGrenze Step 8
            If M <= Sys(i).Top + Sys(i).Height And M > Sys(i).Top Then 'überhaupt nur prüfen, wenn oben oder unten eine Berührung möglich ist, erst Höhe begutachten
                K = 16
                For j = 1 To CInt(Sys(i).Width / 32) 'über gesamte elementlänge prüfen
                    If rGrenze > Sys(i).left + K And Sys(i).left + K > lGrenze Then
                        Abbruch = True 'Abstand verletzt
                        Exit Sub
                    End If
                    K = K + 32
                Next j
            End If
        Next M
    End If
End Sub
Private Sub Elementdarstellen(Element, x, y) 'nur so zur Info behalten
End Sub
Private Sub Trumlänge_Click()
    Eingabe.Visible = False
    Auswahl.Visible = False 'sonst wird dort die default-aktion und damit die auswertung ausgelöst
End Sub
Private Sub Trumlänge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then Call Steuerung_Click(0) 'bandlöschen, entfernen wird nur bei down abgetastet
    'auswertung von return unter defaultbutton
End Sub
Private Sub Trumlänge_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
    'keypress kann den buchstaben noch abfangen
End Sub
Private Sub Horizontal_Change()
    Konstruktion.Scale (Horizontal.Value, Vertikal.Value)-(Konstruktion.Width / Screen.TwipsPerPixelX + Horizontal.Value, Konstruktion.Height / Screen.TwipsPerPixelY + Vertikal.Value)
    ModusCalc = ""
    Call Alleelementeverbinden
    Call Markieren
End Sub
Private Sub Horizontal_DragDrop(Source As Control, x As Single, y As Single)
    Shape1.Visible = False
    NeuEl = 0
End Sub

Private Sub Vertikal_Change()
    Konstruktion.Scale (Horizontal.Value, Vertikal.Value)-(Konstruktion.Width / Screen.TwipsPerPixelX + Horizontal.Value, Konstruktion.Height / Screen.TwipsPerPixelY + Vertikal.Value)
    ModusCalc = ""
    Call Alleelementeverbinden
    Call Markieren
End Sub
Private Sub Vertikal_DragDrop(Source As Control, x As Single, y As Single)
    Shape1.Visible = False 'schönheitskorrektur
    NeuEl = 0
End Sub
Private Sub Horizontal_GotFocus()
    If Konstruktion.Visible = True Then Konstruktion.SetFocus 'sonst blinken diese leisten
End Sub
Private Sub Vertikal_GotFocus()
    If Konstruktion.Visible = True Then Konstruktion.SetFocus
End Sub


'programmteil zum organisieren des bildschirms und zum füllen der tabelle
Public Static Sub Eigenschaftsleiste_Resize()
Dim i As Integer, I1 As Double, I2 As Double, I3 As Double, J1 As Double, J2 As Double, J3 As Double, J4 As Double, Merk As Double
    On Error Resume Next
    If B_Rex.Visible = False Then Exit Sub
    If Noresize = True Then Exit Sub
    Noresize = True
    Anzeigemodus = 0
    If left(Datei(5).Tag, 1) = "E" Then Anzeigemodus = Anzeigemodus + 1 'konstruktion
    If left(Datei(6).Tag, 1) = "E" Then Anzeigemodus = Anzeigemodus + 2 'eigenschaften
    If left(Datei(7).Tag, 1) = "E" Then Anzeigemodus = Anzeigemodus + 4 'fukurve
    If left(Datei(8).Tag, 1) = "E" Then Anzeigemodus = Anzeigemodus + 8 'fehlerliste
    
    Merk = Anzeigemodus
    
    Typenliste.Visible = False
    Konstruktion.Visible = False
    Fehlerliste.Visible = False
    FuKurve.Visible = False
    Eigenschaftsleiste.Visible = False
    Horizontal.Visible = False
    Vertikal.Visible = False
    Eigenschaftsleiste.Align = 4 'damit stellt sich die richtige höhe immer von selbst ein
    DoEvents
    
    I1 = 0 'x-unterteilung
    I3 = Kopfleiste.Width
    I2 = I3 - Screen.TwipsPerPixelX * 405
    J1 = Kopfleiste.Height 'y-unterteilung
    J2 = Eigenschaftsleiste.Height / 2 + J1
    J3 = Eigenschaftsleiste.Height * 2 / 3 + J1
    J4 = Eigenschaftsleiste.Height + J1
    
    FuKurve.left = 0
    Eigenschaftsleiste.Align = 0
    Eigenschaftsleiste.Top = J1
    Konstruktion.left = 0
    Konstruktion.Top = J1

    DoEvents

    Select Case Anzeigemodus
        Case 0
        Case 1
            Konstruktion.Height = J4 - J1
            Konstruktion.Width = I3
        Case 2
            Eigenschaftsleiste.Height = J4 - J1
        Case 3
            Konstruktion.Height = J4 - J1
            Konstruktion.Width = I2
            Eigenschaftsleiste.Height = J4 - J1
        Case 4
            FuKurve.left = 0
            FuKurve.Top = J1
            FuKurve.Width = I3
            FuKurve.Height = J4 - J1
        Case 5
            Konstruktion.Height = J2 - J1
            Konstruktion.Width = I3
            FuKurve.Top = J2
            FuKurve.Width = I3
            FuKurve.Height = J4 - J2
        Case 6
            FuKurve.Top = J1
            FuKurve.Height = J4 - J1
            FuKurve.Width = I2
            Eigenschaftsleiste.Height = J4 - J1
        Case 7
            Konstruktion.Height = J2 - J1
            Konstruktion.Width = I2
            FuKurve.Top = J2
            FuKurve.Width = I2
            FuKurve.Height = J4 - J2
            Eigenschaftsleiste.Height = J4 - J1
        Case 8
            Fehlerliste.left = 0
            Fehlerliste.Top = J1
            Fehlerliste.Width = I3
            Fehlerliste.Height = J4 - J1
        Case 9
            Konstruktion.Height = J2 - J1
            Konstruktion.Width = I3
            Fehlerliste.left = 0
            Fehlerliste.Top = J2
            Fehlerliste.Width = I3
            Fehlerliste.Height = J4 - J2
        Case 10
            Fehlerliste.Top = J1
            Fehlerliste.Height = J4 - J1
            Fehlerliste.Width = I2
            Eigenschaftsleiste.Height = J4 - J1
        Case 11
            Konstruktion.Height = J4 - J1
            Konstruktion.Width = I2
            Eigenschaftsleiste.Height = J3 - J1
            Fehlerliste.left = I2
            Fehlerliste.Width = I3 - I2
            Fehlerliste.Top = J3
            Fehlerliste.Height = J4 - J3
        Case 12
            FuKurve.left = 0
            FuKurve.Top = J1
            FuKurve.Width = I3
            FuKurve.Height = J3 - J1
            Fehlerliste.left = 0
            Fehlerliste.Width = I3
            Fehlerliste.Top = J3
            Fehlerliste.Height = J4 - J3
        Case 13
            Konstruktion.Height = J2 - J1
            Konstruktion.Width = I3
            FuKurve.Top = J2
            FuKurve.Width = I2
            FuKurve.Height = J4 - J2
            Fehlerliste.left = I2
            Fehlerliste.Width = I3 - I2
            Fehlerliste.Top = J2
            Fehlerliste.Height = J4 - J2
        Case 14
            FuKurve.Top = J1
            FuKurve.Width = I2
            FuKurve.Height = J3 - J1
            Eigenschaftsleiste.Height = J4 - J1
            Fehlerliste.left = 0
            Fehlerliste.Width = I2
            Fehlerliste.Top = J3
            Fehlerliste.Height = J4 - J3
       Case 15
            Konstruktion.Height = J2 - J1
            Konstruktion.Width = I2
            FuKurve.Top = J2
            FuKurve.Width = I2
            FuKurve.Height = J4 - J2
            Eigenschaftsleiste.Height = J3 - J1
            Fehlerliste.left = I2
            Fehlerliste.Width = I3 - I2
            Fehlerliste.Top = J3
            Fehlerliste.Height = J4 - J3
    End Select
    
    
    
    If left(Datei(5).Tag, 1) = "E" Then 'konstruktion
        Konstruktion.Width = Konstruktion.Width - Vertikal.Width
        Konstruktion.Height = Konstruktion.Height - Horizontal.Height
        Vertikal.Top = Konstruktion.Top
        Horizontal.left = 0
        Vertikal.left = Konstruktion.left + Konstruktion.Width
        Horizontal.Top = Konstruktion.Top + Konstruktion.Height
        Horizontal.Width = Konstruktion.Width
        Vertikal.Height = Konstruktion.Height
        Elementleiste.Visible = False
        'Line2.X2 = Elementleiste.Width
        Konstruktion.Visible = True
        Horizontal.Visible = True
        Vertikal.Visible = True
        'wenn konstruktion verdeckt würde, dann richtig hinschieben
        If Sys(AktEl).left + Sys(AktEl).Width > Konstruktion.Width / Screen.TwipsPerPixelX Then
            i = (Konstruktion.Width / Screen.TwipsPerPixelX - Sys(AktEl).Width) / 2
            i = Sys(AktEl).left - i
            Konstruktion.Scale (i, 0)-(Konstruktion.Width / Screen.TwipsPerPixelX + i, Konstruktion.Height / Screen.TwipsPerPixelY)
            Call CodeDraw.Alleelementeverbinden
            Horizontal.Value = i
        Else
            Call CodeDraw.Alleelementeverbinden
        End If
        If left(Button(0).Tag, 1) = "A" Then Elementleiste.Visible = True

    End If
    
    If left(Datei(7).Tag, 1) = "E" Then
        FuKurve.Visible = True
        Set Destination = FuKurve
        If FuKurve.Visible = True And Endlos = True And Vollstaendig = True Then
            Set Destination = FuKurve
            FuKurve.Cls
            Call CodeCalc.Grafik(False)
        End If
    End If
    If left(Datei(8).Tag, 1) = "E" Then Fehlerliste.Visible = True
    If left(Datei(6).Tag, 1) = "E" Then 'eigenschaftsleiste
        Eigenschaftsleiste.left = I2
        Eigenschaftsleiste.Width = I3 - I2
        Call Tabelle_gestalten
    End If
    
    Noresize = False

End Sub
Private Sub Tabelle_gestalten()
        Eiged.Width = Eigenschaftsleiste.Width
        Eiged.left = 0
        Eiged.Height = Eigenschaftsleiste.Height - Screen.TwipsPerPixelX * 30
        Eiged.Top = EigButton(0).Height + Screen.TwipsPerPixelY * 1
           
        'von rechts
        EigButton(0).left = Eiged.Width - Screen.TwipsPerPixelX * 232
        EigButton(1).left = Eiged.Width - Screen.TwipsPerPixelX * 210
        EigButton(2).left = Eiged.Width - Screen.TwipsPerPixelX * 188
        EigButton(9).left = Eiged.Width - Screen.TwipsPerPixelX * 166
        EigButton(14).left = Eiged.Width - Screen.TwipsPerPixelX * 144
        EigButton(10).left = Eiged.Width - Screen.TwipsPerPixelX * 122
        
        EigButton(3).left = Eiged.Width - Screen.TwipsPerPixelX * 88
        EigButton(4).left = Eiged.Width - Screen.TwipsPerPixelX * 66
        EigButton(5).left = Eiged.Width - Screen.TwipsPerPixelX * 44
        
            
        'von links
        EigButton(6).left = Screen.TwipsPerPixelX * 0
        EigButton(7).left = Screen.TwipsPerPixelX * 22
        EigButton(8).left = Screen.TwipsPerPixelX * 44
        
        Eiged.Cols = 10
        Eiged.FixedCols = 7
        Eiged.FixedRows = 1
        Eiged.ColWidth(0) = 0 'ordnungskriterium
        Eiged.ColWidth(1) = 0 'ordnungskriterium
        Eiged.ColWidth(2) = 0 'ordnungskriterium
        Eiged.ColWidth(3) = 0
        Eiged.ColWidth(4) = 0
        'End If
        Eiged.ColWidth(5) = 0 'sekundäres Ordnungskriterium, wir dynamisch geändert
        Eiged.ColWidth(6) = Screen.TwipsPerPixelX * 243 'eigenschaft
        Eiged.ColWidth(7) = Screen.TwipsPerPixelX * 55 'einheit
        Eiged.ColWidth(8) = Screen.TwipsPerPixelX * 85 'einstellung
        Eiged.ColWidth(9) = 0 'eigenschaftsart

        Eigenschaftsleiste.Visible = True

End Sub
Public Sub Tabelle_ausfuellen(ByVal L As Integer)
'l enthält das element , das abgebildet werden soll. Überflüssig bei Anzeige aller elemente
Eingabe.Visible = False
Auswahl.Visible = False
If Eigenschaftsleiste.Visible = False Then Exit Sub 'And NeuEl = 0 Then Exit Sub
If Lastaktel = 1000 Then Exit Sub 'alle elemente befinden sich noch in der tabelle, neuausfüllen verhindern
Dim Ausgabe As Boolean
Dim i, M, N, K, P, Knopf, Anzelemente As Integer
Dim Merk$
    Lastaktel = L
    Eiged.Clear
    Eiged.Visible = False 'darstellung vorm sortieren unterdrücken
    
    If left(EigButton(5).Tag, 1) = "E" Then 'eigenschaften aller elemente
        For L = 1 To Maxelementindex
            If L = 2 Then L = L + 1 'die sicherung des bandes ist hier abgelegt
            If Sys(L).Element <> "" Then GoSub Element_eintragen
        Next L
        Lastaktel = 1000
    Else 'nur eigenschaften eines elements
        GoSub Element_eintragen
    End If
    On Error Resume Next
    Eiged.Row = 0
    Eiged.Col = 4
    Eiged.CellFontBold = True
    Eiged.Col = 6
    Eiged.CellFontBold = True
    Eiged.Col = 7
    Eiged.CellFontBold = True
    Eiged.Col = 8
    Eiged.CellFontBold = True
    Eiged.TextMatrix(0, 4) = Lang_Res(403)  'Element
    Eiged.TextMatrix(0, 6) = Lang_Res(404)  'Eigenschaft
    Eiged.TextMatrix(0, 7) = Lang_Res(406)  'Einheit
    Eiged.TextMatrix(0, 8) = Lang_Res(405) 'Einstellung
    If i = 0 Then Exit Sub 'keine eigenschaft anzeigbar
    Eiged.Rows = i + 1
    Eiged.Row = 1
    Eiged.RowSel = 1
    
    'ordnungsspalte nach der Elementspalte vorbereiten
    For L = 1 To Eiged.Rows - 1
        
        'zunächst mal leeren
        Eiged.TextMatrix(0, 5) = ""
        
        'nach eigenschaft alphabetisch
        If left(EigButton(3).Tag, 1) = "E" Then Eiged.TextMatrix(L, 5) = Eiged.TextMatrix(L, 6)
        
        'nach einheit
        If left(EigButton(4).Tag, 1) = "E" Then Eiged.TextMatrix(L, 5) = Eiged.TextMatrix(L, 7)
        
    Next L
    
    Eiged.Col = 5 'col muß links, colsel rechts sein
    If left(EigButton(5).Tag, 1) = "E" Then Eiged.Col = 4 'mehrere Elemente, also sortieren nach 2 kriterien
    
    Eiged.ColSel = 5
    Eiged.Sort = 5 'alphabetisch aufsteigend
    
    Eiged.Col = 8
    Eiged.ColSel = Eiged.Col 'selektierung aufheben
    Eiged.MergeCells = 2 'rows, 3 ist cols
    
    'zusätzliche überschriften einfügen, um elemente abzugrenzen
    If left(EigButton(5).Tag, 1) = "E" Then 'nur, wenn mehrere elemente dargestellt werden sollen
        On Local Error Resume Next
        For i = 1 To Eiged.Rows - 1 + Anzelemente  'er bekommt hier mehr zeilen
            'If I > Eiged.Rows - 1 Then GoTo Markierung 'zur sicherheit
            If Eiged.TextMatrix(i, 4) <> Merk$ Then
                If i < Eiged.Rows Then
                    Merk$ = Eiged.TextMatrix(i, 4)
                    Eiged.AddItem "", i
                    Eiged.TextMatrix(i, 4) = "Überschrift"
                    Eiged.Col = 6
                    Eiged.Row = i
                    Eiged = Eiged.TextMatrix(i + 1, 4)
                    Eiged.CellAlignment = 1
                    Eiged.CellFontBold = True
                    Eiged.CellBackColor = &H80FFFF   'vbyellow
                    Eiged.Col = 7
                    Eiged.CellBackColor = &H80FFFF   'vbyellow
                    Eiged.Col = 8
                    Eiged.CellBackColor = &H80FFFF   'vbyellow
                End If
            End If
        Next i
    End If
    
'Markierung:
    'und die eigenschaften zuletzt eintragen
    Call TabelleEig_ausfuellen

Exit Sub
        
Element_eintragen:
    If L > 9 Then Anzelemente = Anzelemente + 1
    M = 0
    Do
        M = M + 1
    Loop Until El(0).Eig(M) = Sys(L).Tag
    
    'wenns als transportgut auf einem tisch liegt,
    If Sys(L).Tag = "201" Then
        El(32).Eig(M) = 1
        El(36).Eig(M) = 1
        If Sys(Sys(L).Zugehoerigkeit).Tag = "101" Then
            'sind nicht soviele infos erforderlich
            'und auch nur, wenn kein abweiser da ist
            El(32).Eig(M) = 0
            'El(36).Eig(M) = 0'erstmal lieber doch, weil es sonst nicht veränderbar ist, wohl aber ausgedruckt wird
            For P = 1 To Maxelementindex
                If Sys(L).Zugehoerigkeit = Sys(P).Zugehoerigkeit Then
                    If Sys(P).Tag = "205" Then 'bei einem abweiser aber doch
                        El(32).Eig(M) = 1
                        El(36).Eig(M) = 1
                    End If
                    If Sys(P).Tag = "204" Then 'bei einem stau wenigstens die beladungsart
                        El(36).Eig(M) = 1
                    End If
                End If
            Next P
        End If
        'programmteil auch in der Vollstaendigkeitskontrolle vorhanden
    End If
    
    For N = -10 To Eigenschaftszahl
        If N = 0 Then N = 1 'den überspringen, weil's ihn nicht gibt
        
        If El(N).Eig(M) <> "" Then
            
            If left(EigButton(6).Tag, 1) = "E" And Sys(L).B(N) = False Then
                'garnichts, soll markierte darstellen, dieser eintrag ist aber nicht markiert
            Else
                'also gehts seinen normalen gang
                If left(EigButton(0).Tag, 1) = "E" And InStr(El(N).Eig(M), "1") > 0 Then Ausgabe = True 'muss eingaben
                If left(EigButton(0).Tag, 1) = "E" And InStr(El(N).Eig(M), "4") > 0 Then Ausgabe = True 'muss, aber nicht pflicht
                If left(EigButton(1).Tag, 1) = "E" And InStr(El(N).Eig(M), "2") > 0 Then Ausgabe = True 'kann
                If left(EigButton(2).Tag, 1) = "E" And InStr(El(N).Eig(M), "3") > 0 Then Ausgabe = True 'ergebnisse
                If left(EigButton(9).Tag, 1) = "E" And InStr(El(N).Eig(M), "5") > 0 Then Ausgabe = True 'beschleunigung
                If left(EigButton(14).Tag, 1) = "E" And InStr(El(N).Eig(M), "6") > 0 Then Ausgabe = True 'durchbiegung
                If left(EigButton(10).Tag, 1) = "E" And InStr(El(N).Eig(M), "8") > 0 Then Ausgabe = True 'frequenzen
                If Ausgabe = True Then
                    Ausgabe = False 'fürs nächste mal
                    i = i + 1 'protokolliert die zeilenzahl mit
                    If Eiged.Rows <= i Then Eiged.Rows = Eiged.Rows + 1 'wird unten verringert
                    Eiged.Row = i
                    
                    'alles ausser muss-eingaben wird kursiv gemacht
                    If InStr(El(N).Eig(M), "1") = 0 And InStr(El(N).Eig(M), "4") = 0 Then
                        Eiged.Col = 6
                        Eiged.CellFontItalic = True
                    End If
                    
                    'ergebnisse grau und unabänderbar
                    If InStr(El(N).Eig(M), "3") > 0 Then
                        Eiged.Col = 8
                        Eiged.CellBackColor = &H80000000
                    End If
                    
                   
                    'wenn man hier alignment macht, steht noch nichts drin und die Geschwindigkeit ist akzeptabel
                    Eiged.Col = 8
                    Eiged.CellAlignment = 2 'nach rechts zu den zahlen
                    Eiged.Col = 4
                    Eiged.CellAlignment = 1 'formatierung links
                    Eiged.Col = 7
                    Eiged.CellAlignment = 1 'formatierung links
                    If left(El(N).Feldart, 4) = "zahl" Then Eiged.CellBackColor = &H80000000                'grau
                    
                    Eiged.TextMatrix(i, 0) = K 'muß, kann oder i.E.
                    Eiged.TextMatrix(i, 1) = El(N).Feldart
                    Eiged.TextMatrix(i, 2) = L 'nummer des elements
                    Eiged.TextMatrix(i, 3) = N 'nummer der eigenschaft, nur bei komplettanzeige erforderlich
                    
                    Eiged.TextMatrix(i, 4) = "(" & L & ") " & Sys(L).Element 'bezeichnung des elements, nur bei komplettanzeige erforderlich
                    Eiged.TextMatrix(i, 9) = El(N).Eig(M) 'eigenschaftsart
                    Merk$ = El(N).Eigenschaft
                        Merk$ = El(N).Eigenschaft
                    If Val(El(N).Eig(M)) = 4 Then
                        Eiged.TextMatrix(i, 6) = Merk$ & " (0)" 'die null ist zulässig
                    Else
                        Eiged.TextMatrix(i, 6) = Merk$ 'bezeichnung des eigenschaft, nur bei komplettanzeige erforderlich
                    End If
                    
                    Eiged.Col = 6
                    If Sys(L).B(N) = True Then
                        Eiged.CellForeColor = QBColor(12)
                    Else
                        Eiged.CellForeColor = QBColor(0)
                    End If
                    Eiged.TextMatrix(i, 7) = El(N).Einheit
                    Eiged.TextMatrix(i, 8) = "" 'letzten eintrag löschen
                    'wenn englisch, dann characteristic
                End If
            End If
        End If
    Next N
    Return
End Sub
Private Sub EigButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim j, i, K As Integer
Dim AuflDehn As Double
    Call Mother.Neue_Knopfverwaltung(EigButton(Index))
    
    Select Case Index
        Case 3 To 6
            If Index = 5 Then 'alle elemente eintragen
                Call Tabelle_gestalten
            End If
            If left(EigButton(3).Tag, 1) = "E" And left(EigButton(4).Tag, 1) = "E" Then 'würde sich widersprechen
                If Index = 3 Then Call Mother.Knopfverwaltung(4, "KleinerKnopf", "Eigbutton", "B_Rex")
                If Index = 4 Then Call Mother.Knopfverwaltung(3, "KleinerKnopf", "Eigbutton", "B_Rex")
            End If
            If left(EigButton(3).Tag, 1) = "A" And left(EigButton(4).Tag, 1) = "A" Then 'einer muß drinnen sein
                If Index = 3 Then Call Mother.Knopfverwaltung(4, "KleinerKnopf", "Eigbutton", "B_Rex")
                If Index = 4 Then Call Mother.Knopfverwaltung(3, "KleinerKnopf", "Eigbutton", "B_Rex")
            End If
            Lastaktel = 0 'die tabelle muß angefaßt werden
            If Eigenschaftsleiste.Visible = True Then Call Tabelle_ausfuellen(AktEl)
        
        Case 0 To 2, 9, 10, 14 'mke, a,ytr,f
            'bei linker maustaste darf nur einer draussen sein
            If Button = 1 Then
                If Index <> 0 And left(EigButton(0).Tag, 1) = "E" Then Call Mother.Neue_Knopfverwaltung(EigButton(0))
                If Index <> 1 And left(EigButton(1).Tag, 1) = "E" Then Call Mother.Neue_Knopfverwaltung(EigButton(1))
                If Index <> 2 And left(EigButton(2).Tag, 1) = "E" Then Call Mother.Neue_Knopfverwaltung(EigButton(2))
                If Index <> 9 And left(EigButton(9).Tag, 1) = "E" Then Call Mother.Neue_Knopfverwaltung(EigButton(9))
                If Index <> 10 And left(EigButton(10).Tag, 1) = "E" Then Call Mother.Neue_Knopfverwaltung(EigButton(10))
                If Index <> 14 And left(EigButton(14).Tag, 1) = "E" Then Call Mother.Neue_Knopfverwaltung(EigButton(14))
            End If
            
            Lastaktel = 0 'die tabelle muß angefaßt werden
            If Eigenschaftsleiste.Visible = True Then Call Tabelle_ausfuellen(AktEl)
        
        Case 7 To 8
            Eiged.Visible = False
            Eiged.Col = 6
            For j = 1 To Eiged.Rows - 1
                Eiged.Row = j
                If Eiged.TextMatrix(Eiged.Row, 2) <> "" Then
                    If Index = 8 Then
                        Eiged.CellForeColor = QBColor(12)
                        Sys(Eiged.TextMatrix(Eiged.Row, 2)).B(Eiged.TextMatrix(Eiged.Row, 3)) = True
                    Else
                        Eiged.CellForeColor = QBColor(0)
                        Sys(Eiged.TextMatrix(Eiged.Row, 2)).B(Eiged.TextMatrix(Eiged.Row, 3)) = False
                    End If
                End If
            Next j
            Eiged.Visible = True
        

        Case 13 'alle typen durchrechnen
            
    End Select
End Sub
Public Sub TabelleEig_ausfuellen()
'erst zum schluß, nachdem alle anderen spalten gefüllt wurden
If Eigenschaftsleiste.Visible = False Then Exit Sub
On Local Error Resume Next
Dim i, x, y, M As Integer
Dim a$
    x = Eiged.Col
    y = Eiged.Row
    Eiged.Visible = False
    For i = 1 To Eiged.Rows - 1
        If Eiged.TextMatrix(i, 4) = "Überschrift" Then i = i + 1
        Select Case left(Eiged.TextMatrix(i, 1), 4) 'feldart liste oder text
            Case "zahl"
                Select Case Right(Eiged.TextMatrix(i, 1), 1)
                    Case 0
                        a$ = "#####0"
                    Case 1
                        a$ = "#####0.0"
                    Case 2
                        a$ = "#####0.00"
                    Case 3
                        a$ = "#####0.000"
                End Select
                Eiged.TextMatrix(i, 8) = Format(Sys(Eiged.TextMatrix(i, 2)).E(Eiged.TextMatrix(i, 3)), a$)
            Case "text"
                If Sys(Eiged.TextMatrix(i, 2)).S(Abs(Eiged.TextMatrix(i, 3))) = "" Then
                    Select Case Eiged.TextMatrix(i, 3)
                        Case -1, -2, -5
                            a$ = Lang_Res(169)  'kein Typ gewählt
                        Case -4
                            a$ = Lang_Res(170)  'Laufseite / Antriebsseite
                        Case -3
                            a$ = Lang_Res(171)  'Tragseite / Funktionsseite
                    End Select
                    Sys(Eiged.TextMatrix(i, 2)).S(Abs(Eiged.TextMatrix(i, 3))) = a$
                End If
                Eiged.TextMatrix(i, 8) = Sys(Eiged.TextMatrix(i, 2)).S(Abs(Eiged.TextMatrix(i, 3)))
            Case "list"
                Select Case Eiged.TextMatrix(i, 3)
                    Case 14 'oberfläche zum band, stahl
                        Eiged.TextMatrix(i, 8) = Kst(Sys(Eiged.TextMatrix(i, 2)).E(14)).Bezeichnung
                    Case 15 'oberfläche zum band/transportgut
                        Eiged.TextMatrix(i, 8) = Kst(Sys(Eiged.TextMatrix(i, 2)).E(15)).Bezeichnung
                    Case 36 'beladungsart
                        Eiged.TextMatrix(i, 8) = Kst(Sys(Eiged.TextMatrix(i, 2)).E(36)).Bezeichnung
                    Case 41 'bandkontaktfläche zum element
                        If Sys(Eiged.TextMatrix(i, 2)).E(41) = 0 Then
                            Sys(Eiged.TextMatrix(i, 2)).E(41) = 1 'ls
                        End If
                        If Sys(Eiged.TextMatrix(i, 2)).E(41) = 1 Then
                            Eiged.TextMatrix(i, 8) = Sys(1).S(4)
                            If Eiged.TextMatrix(i, 8) = "" Then Eiged.TextMatrix(i, 8) = El(-4).Eigenschaft
                        Else
                            Eiged.TextMatrix(i, 8) = Sys(1).S(3)
                            If Eiged.TextMatrix(i, 8) = "" Then Eiged.TextMatrix(i, 8) = El(-3).Eigenschaft
                        End If
                    Case 42 'bandkontaktfläche zu den tragrollen
                        If Sys(Eiged.TextMatrix(i, 2)).E(42) = 0 Then
                            Sys(Eiged.TextMatrix(i, 2)).E(42) = 2 'ts
                        End If
                        If Sys(Eiged.TextMatrix(i, 2)).E(42) = 1 Then
                            Eiged.TextMatrix(i, 8) = Sys(1).S(4)
                            If Eiged.TextMatrix(i, 8) = "" Then Eiged.TextMatrix(i, 8) = El(-4).Eigenschaft
                        Else
                            Eiged.TextMatrix(i, 8) = Sys(1).S(3)
                            If Eiged.TextMatrix(i, 8) = "" Then Eiged.TextMatrix(i, 8) = El(-3).Eigenschaft
                        End If
                    Case 47 'material der Scheibe
                        Eiged.TextMatrix(i, 8) = Kst(Sys(Eiged.TextMatrix(i, 2)).E(47)).Bezeichnung
                    
                    Case 48 'oberflächenmaterial zum Band, nur messerkante
                        Eiged.TextMatrix(i, 8) = Kst(Sys(Eiged.TextMatrix(i, 2)).E(48)).Bezeichnung
                    Case 60 'aut. überlast
                        Eiged.TextMatrix(i, 8) = Kst(Sys(Eiged.TextMatrix(i, 2)).E(60)).Bezeichnung
                End Select
        End Select
        
        Eiged.Row = i
        Eiged.Col = 8
        If Eiged = "" Then
            Eiged = "-"
        End If
        
        'ergebnisse bei unfertigen anlagen geheimhalten
        If InStr(Eiged.TextMatrix(i, 9), "3") > 0 And (Vollstaendig = False Or Endlos = False) Then
            Eiged.TextMatrix(i, 8) = "-"
        End If
        
        'immer bei auswahlen oder textfeldern, weil sie keine einheit besitzen
        If left(Eiged.TextMatrix(i, 1), 4) <> "zahl" Then
            Eiged.TextMatrix(i, 7) = Eiged.TextMatrix(i, 8)
        End If
        
        'damit bei auswahlen die spalten zusammengefaßt werden
        Eiged.MergeRow(i) = True
        
    Next i
    Eiged.Visible = True
    Eiged.Col = x
    Eiged.Row = y
    
   
        
    If left(Button(0).Tag, 1) = "E" Then
        Set Destination = Konstruktion 'sonst wird die grafik in die fukurve gemalt, weil die vorher destination war
        Call CodeDraw.Zweischeibentrieb_zeichnen
    End If
    
End Sub


'auflegedehnungsverwaltung
Private Sub FuKurve_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragDrop = False Then Auflegedehnunganzeiger.Visible = False
Dim i As Double
    If Abs(y - AuflTrumKraft) < (FuScaleY1 - FuScaleY2) / 50 Then
        If Zweischeiben = True Then
            i = Sys(1).E(74)
        Else
            i = Sys(1).E(33)
        End If
        If x > 0 And x < i Then
            FuKurve.MousePointer = 7 'nordsüd
        Else
            FuKurve.MousePointer = 0
        End If
    Else
        If FuKurve.MousePointer <> 0 Then FuKurve.MousePointer = 0

    End If

End Sub
Private Sub FuKurve_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Double
    If Zweischeiben = True Then
        i = Sys(1).E(74)
    Else
        i = Sys(1).E(33)
    End If
    Auflegedehnung.X1 = 0
    Auflegedehnung.X2 = i
    Auflegedehnung.Y1 = AuflTrumKraft
    Auflegedehnung.Y2 = Auflegedehnung.Y1
    
    Auflegedehnunganzeiger.X1 = 0
    Auflegedehnunganzeiger.X2 = i
    Auflegedehnunganzeiger.Y1 = AuflTrumKraft
    Auflegedehnunganzeiger.Y2 = Auflegedehnung.Y1
    
    Auflegedehnunganzeiger.Visible = True
    Auflegedehnung.Drag vbBeginDrag
    DragDrop = True
End Sub
Private Sub FuKurve_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Dim i As Double
Dim DragEnd As Boolean
    
    If Zweischeiben = True Then
        i = Sys(1).E(74)
    Else
        i = Sys(1).E(33)
    End If
    Auflegedehnunganzeiger.X1 = 0
    Auflegedehnunganzeiger.X2 = i
    Auflegedehnunganzeiger.Y1 = y
    Auflegedehnunganzeiger.Y2 = y
    
    'ränder setzen, bei denen dragdrop beendet wird
    If x < 0 Then DragEnd = True 'links
    If x > i Then DragEnd = True 'rechts
    i = Abs(FuScaleY1 - FuScaleY2) / 40 'bereich
    If FuScaleY2 < y And FuScaleY2 + i > y Then DragEnd = True 'unten
    'oben wird genutzt, um in 0.5 schritten hochzustellen (unter dragdrop)
        
    If DragEnd = True Then
        DragDrop = False
        Auflegedehnung.Drag vbEndDrag
        Auflegedehnunganzeiger.Visible = False
    End If
End Sub
Private Sub FuKurve_DragDrop(Source As Control, x As Single, y As Single)
    If DragDrop = False Then Exit Sub
    
    If y < 0 Then 'wenn unten, dann vorgabe ausschalten
        Sys(1).E(1) = 0
    Else
        'wenn oben, dann in 0.5er Schritten weiter
        Sys(1).E(1) = y * 2 / (SystemTyp.Kraftdehnung * Sys(1).E(34))
        
        i = Abs(FuScaleY1 - FuScaleY2) / 15 'bereich
        If FuScaleY1 > y And FuScaleY1 - i < y Then 'dann noch auf den nächsten 0.5er aufrunden
            Sys(1).E(1) = (Round((Sys(1).E(1) + 0.25) * 2)) / 2
        End If
    End If
    Mother.H = ""
    Auflegedehnunganzeiger.Visible = False
    Call CodeCalc.Rechnungssteuerung("C")
    Call TabelleEig_ausfuellen
    Call Dateiverwaltung.Undo(0)
End Sub





















