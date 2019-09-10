VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   " Right:"
      Height          =   6855
      Left            =   5100
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   6360
         Width           =   2835
      End
      Begin VB.ListBox lstContent 
         Height          =   4935
         Index           =   1
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   900
         Width           =   4695
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   4500
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtPath 
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   4515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Left:"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   6360
         Width           =   2835
      End
      Begin VB.CommandButton cmdDeleteWhatisOnRight 
         Caption         =   "Delete what is in Right"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   5940
         Width           =   4695
      End
      Begin VB.ListBox lstContent 
         Height          =   4935
         Index           =   0
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   4695
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   4500
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4395
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   7020
      Width           =   10035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click(Index As Integer)
  Dim sFile As String
  sFile = ShowOpenDlg(Me, "*.csv (CSV)|*.csv", "open Comma seperated file", App.Path)
  If (sFile <> "") Then
    txtPath(Index).Text = sFile
  End If
  
End Sub

Private Sub cmdDeleteWhatisOnRight_Click()
  Dim i As Integer
  Dim j As Integer
  
  Dim iTo As Integer
  Dim jTo As Integer
  
  
  lstContent(0).Visible = False
  lstContent(1).Visible = False
  
  iTo = lstContent(0).ListCount - 1
  
  
  For i = iTo To 0 Step -1
    jTo = lstContent(1).ListCount - 1
    For j = jTo To 0 Step -1
      If lstContent(1).List(j) = lstContent(0).List(i) Then
        lstContent(1).RemoveItem j
        lstContent(0).RemoveItem i
        Exit For
      End If
    Next 'j
    lblStatus.Caption = i & "  " & "Left: " & lstContent(0).ListCount & "  " & "Right: " & lstContent(1).ListCount
    DoEvents
  Next 'i
  lstContent(0).Visible = True
  lstContent(1).Visible = True
  
  lblStatus.Caption = "Finished.  " & "Left: " & lstContent(0).ListCount & "  " & "Right: " & lstContent(1).ListCount
End Sub

Private Sub cmdSave_Click(Index As Integer)
  Dim sFile As String
  sFile = ShowSaveDlg(Me, "*.csv|*.csv", "Save", App.Path)
  If sFile <> "" Then
    If Right$(LCase(sFile), 4) <> ".csv" Then
      sFile = sFile & ".csv"
    End If
    
    Dim i As Integer
    Dim iTo As Integer
    
    iTo = lstContent(Index).ListCount - 1
    Dim ff As Integer
    ff = FreeFile
    Open sFile For Output As #ff
    For i = 0 To iTo
      Print #ff, lstContent(Index).List(i)
    Next 'i
    Close #ff
    lblStatus.Caption = "Saved: " & sFile
  End If
End Sub

Private Sub txtPath_Change(Index As Integer)
  If FileExists(txtPath(Index).Text) Then
      Dim sContent As String
      sContent = ReadFile(txtPath(Index).Text)
      Dim i As Integer
      Dim sSplit() As String
      sSplit = Split(sContent, vbCrLf)
      
      Dim iTo As Integer
      iTo = UBound(sSplit())
      lstContent(Index).Clear
      lstContent(Index).Visible = False
      For i = 0 To iTo
        lstContent(Index).AddItem sSplit(i)
      Next 'i
      lstContent(Index).Visible = True
      lblStatus.Caption = "Loaded file: " & txtPath(Index).Text
  Else
    lblStatus.Caption = "File not found: " & txtPath(Index).Text
  End If
  
End Sub
