VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmParse 
   BackColor       =   &H8000000D&
   Caption         =   "Parsing Program"
   ClientHeight    =   4245
   ClientLeft      =   4965
   ClientTop       =   3645
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   6135
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5040
      TabIndex        =   4
      Top             =   3450
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Count"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2160
      TabIndex        =   1
      Top             =   990
      Width           =   1395
   End
   Begin MSComDlg.CommonDialog C1 
      Left            =   90
      Top             =   780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".MDB"
      FileName        =   "\\102.15.94.199\c$\Documents and Settings\All Users\Application Data\CR2\BWAC\Archives\*.*"
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "Consunants, Vowels And Numbers Counter Program    - By Habib Abdulkarim"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   30
      TabIndex        =   5
      Top             =   3990
      Width           =   6105
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   60
      TabIndex        =   3
      Top             =   1470
      Width           =   915
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Save Your Message in a Text File (Note Pad), Then Click ""Count"" Below And See The Result. "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   6045
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2055
      Left            =   30
      TabIndex        =   0
      Top             =   1860
      Width           =   5985
   End
End
Attribute VB_Name = "frmParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
C1.FileName = "*.txt"
C1.ShowOpen
Dim TextLine As String
If C1.FileName = "*.txt" Then Exit Sub
Open C1.FileName For Input As #1
Do Until EOF(1)
Line Input #1, TextLine
Dim xx As Long
Dim Wordlength As Long, Word As String
Dim Wajen As String, VowelCount As Integer, NumberCount As Integer, ConsCount As Integer
Dim TerminalPunctuation As Integer
On Error Resume Next
    Word = TextLine
    Wordlength = Len(Word)
    If Wordlength = 0 Then Wordlength = 1
    For xx = 1 To Wordlength
    Wajen = UCase$(Mid$(Word, xx, 1))
    If InStr("BCDFGHJKLMNPQRSTVWXYZ", Wajen) Then
    ConsCount = ConsCount + 1
    ElseIf InStr("AEIOU", Wajen) Then VowelCount = VowelCount + 1
    ElseIf InStr("1234567890", Wajen) Then NumberCount = NumberCount + 1
    ElseIf InStr(".,""?!", Wajen) Then TerminalPunctuation = TerminalPunctuation + 1
    End If
    Next xx
 Loop
Label1.Caption = "The Message Contains: " & vbCrLf & FormatNumber(VowelCount) & " Vowels, " & vbCrLf & FormatNumber(ConsCount) & " Consunants, " _
  & vbCrLf & FormatNumber(NumberCount) & " Numbers And " & vbCrLf & FormatNumber(TerminalPunctuation) & " Punctuation Marks."
Close #1
End Sub
