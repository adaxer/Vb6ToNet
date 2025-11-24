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
