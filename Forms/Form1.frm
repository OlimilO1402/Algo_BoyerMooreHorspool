VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test of searching for the needle in the haystack with an algorithm by Boyer, Moore and Horspool"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
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
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   600
      Width           =   9615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription, vbInformation
End Sub

Private Sub Command1_Click()
    Dim H As String: H = "ui ui ui ui ui the needle is in the haystack, go and find the needle"
    Dim n As String: n = "needle"
    Dim s As String
    Dim haystack() As Byte: haystack = StrConv(H, vbFromUnicode)
    Dim needle()   As Byte: needle = StrConv(n, vbFromUnicode)
    
    Dim pos1 As Long: pos1 = 8 '4 * 2
    Dim pos2 As Long: pos2 = UBound(haystack) - UBound(needle)
    Debug_Print H
    Debug_Print n
    Dim f As Long
    f = find(haystack, needle)
    If f = pos1 Then
        Debug_Print pos1 & " " & f
    ElseIf f = pos2 Then
        Debug_Print pos2 & " " & f
    End If
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub
