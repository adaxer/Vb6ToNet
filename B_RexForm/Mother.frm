VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.MDIForm Mother 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   Caption         =   "B_Rex"
   ClientHeight    =   11700
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   18615
   Icon            =   "Mother.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.PictureBox Fussleiste 
      Align           =   2  'Unten ausrichten
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   396
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   18615
      TabIndex        =   0
      Top             =   11310
      Width           =   18615
      Begin VB.TextBox H 
         BackColor       =   &H00C0C0C0&
         Height          =   312
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   60
         Width           =   6492
      End
      Begin VB.Label SeitenZahl 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   9120
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin PicClip.PictureClip GrosseKnoepfe 
      Left            =   6000
      Top             =   1440
      _ExtentX        =   1667
      _ExtentY        =   15081
      _Version        =   393216
      Picture         =   "Mother.frx":030A
   End
End
Attribute VB_Name = "Mother"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

