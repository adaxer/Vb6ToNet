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

Private ArtListeMode As Integer
Private RPS_gesperrt As Boolean
Private Erst_speichern As Boolean
Private Arbeitet As Boolean
Private NavigatorZoom As Boolean


Public Sub MDIForm_Load()
Dim i As Integer
Dim a$
    On Local Error Resume Next
    Abfragesprache = "de"
    Sp = 0
    SystemTyp.Netto = left(SystemTyp.Artnr, 6)
    
    If UserStatus = 0 Then
        LetztesFenster = 9
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Artikeldaten.State = 1 Then Artikeldaten.Close
End Sub


Public Sub Knopfverwaltung(ByVal i As Integer, Knopfgröße$, Knopfname$, Programmteil$)
Dim j As Integer
On Local Error Resume Next
    Select Case Knopfgröße$

        Case "GrosserKnopf"

            GrosseKnoepfe.ClipWidth = 31
            GrosseKnoepfe.ClipHeight = 30

            Select Case Programmteil$
                Case "B_Rex"
                    If Knopfname$ = "Button" Then

                        If left(B_Rex.Datei(i).Tag, 1) = "E" Then
                            B_Rex.Datei(i).Tag = "A" & Right(B_Rex.Datei(i).Tag, Len(B_Rex.Datei(i).Tag) - 1)
                            GrosseKnoepfe.ClipX = 0

                        Else
                            If left(B_Rex.Datei(i).Tag, 1) = "A" Then
                                B_Rex.Datei(i).Tag = "E" & Right(B_Rex.Datei(i).Tag, Len(B_Rex.Datei(i).Tag) - 1)
                                GrosseKnoepfe.ClipX = 31
                            End If
                        End If

                        If B_Rex.Datei(i).Tag = "" Then Exit Sub 'kein bild zum tauschen

                        If i = 4 Then GrosseKnoepfe.ClipY = 13 * 30
                        If i = 5 Then GrosseKnoepfe.ClipY = 7 * 30
                        If i = 6 Then GrosseKnoepfe.ClipY = 10 * 30
                        If i = 7 Then GrosseKnoepfe.ClipY = 11 * 30
                        If i = 8 Then GrosseKnoepfe.ClipY = 12 * 30
                        If i = 9 Then GrosseKnoepfe.ClipY = 9 * 30
                        If i = 10 Then GrosseKnoepfe.ClipY = 14 * 30
                        If i = 11 Then GrosseKnoepfe.ClipY = 15 * 30


                        B_Rex.Datei(i).Picture = GrosseKnoepfe.Clip
                    End If

            End Select
    End Select

End Sub
Public Sub Neue_Knopfverwaltung(O As Object)
On Error GoTo Errorhandler
'A/E stehen vorne, dann koennen noch weitere anweisungen folgen, muessen aber nicht
    If left(O.Tag, 1) <> "A" And left(O.Tag, 1) <> "E" Then Exit Sub

    If left(O.Tag, 1) = "E" Then
        O.Tag = "A" & Right(O.Tag, Len(O.Tag) - 1)
    Else
        If left(O.Tag, 1) = "A" Then O.Tag = "E" & Right(O.Tag, Len(O.Tag) - 1)
    End If


Errorhandler:

End Sub

