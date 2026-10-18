VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6375
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "MString>>"
      Height          =   375
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test for Finding a Needle in a Haystack Using an Algorithm by Boyer, Moore and Horspool"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   600
      Width           =   10095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Set MApp.CurForm = Me
    MApp.AppInit
End Sub

Private Sub BtnInfo_Click()
    MsgBox MApp.AppInfo, vbInformation
End Sub

Private Sub Command1_Click()
    MApp.Debug_Print MApp.Haystack
    MApp.Debug_Print MApp.Needle
    If MApp.Start = 0 Then
        MApp.Start = MBoyMorHorspW.FindStr(MApp.Haystack, MApp.Needle)
    Else
        MApp.Start = MBoyMorHorspW.FindNext
    End If
    MApp.Debug_Print "pos: " & MApp.Start
    MApp.Start = MApp.Start + Len(MApp.Needle)
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Form_Resize()
    Dim l As Single: l = 0
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move l, t, W, H
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MBoyMorHorspW.Clear
End Sub
