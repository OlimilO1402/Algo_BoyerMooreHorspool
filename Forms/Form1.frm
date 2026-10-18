VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   420
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows-Standard
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
      TabIndex        =   3
      Top             =   600
      Width           =   10095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(W)>>"
      Height          =   375
      Left            =   8880
      TabIndex        =   2
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
End
Attribute VB_Name = "Form1"
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
    Dim Haystack() As Byte: Haystack = StrConv(MApp.Haystack, vbFromUnicode)
    Dim Needle()   As Byte: Needle = StrConv(MApp.Needle, vbFromUnicode)
    'Dim haystack As String: haystack = StrConv(m_Haystack, vbFromUnicode)
    'Dim needle   As String:   needle = StrConv(m_Needle, vbFromUnicode)
    
    MApp.Debug_Print MApp.Haystack
    MApp.Debug_Print MApp.Needle
    If MApp.Start = 0 Then
        MApp.Start = MBoyMorHorsp.Find(Haystack, Needle)
    Else
        MApp.Start = MBoyMorHorsp.Find(Haystack, Needle, MApp.Start)
    End If
    MBoyMorHorsp.Debug_Print "pos: " & MApp.Start
    MApp.Start = MApp.Start + Len(MApp.Needle)
End Sub

Private Sub Command2_Click()
    Form2.Show
End Sub

Private Sub Form_Resize()
    Dim l As Single: l = 0
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move l, t, W, H
End Sub
