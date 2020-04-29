VERSION 5.00
Begin VB.Form frmPush 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   360
      Value           =   1
      Width           =   3975
   End
   Begin VB.Label lblPage 
      Caption         =   "Level: 1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmPush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim x As Long
Dim Signs(0 To 5) As Boolean
Dim Quavers(0 To 7) As Integer

Private Sub Form_Load()
    'Set levels where Quavers are found
    Quavers(0) = 12
    Quavers(1) = 21
    Quavers(2) = 30
    Quavers(3) = 44
    Quavers(4) = 56
    Quavers(5) = 63
    Quavers(6) = 78
    Quavers(7) = 89
    Call Calculate
End Sub

Private Sub HScroll1_Change()
    lblPage.Caption = "Level: " & HScroll1.Value
    Call Calculate
End Sub

Private Sub Calculate()
    Dim j As Integer

    For i = 0 To 5
        Signs(i) = True
    Next
    x = 0
    For i = 1 To HScroll1.Value
        'Level 64: just add 1
        If i = 64 Then
            x = x + 1
        Else
            If i > 64 Then
                'Level >64: start algorithm from beginning, i.e. at level 1
                j = i - 64
            Else
                'Level <64: proceed normally
                j = i
            End If
            If (j + 32) Mod 64 = 0 Then
                Increment (5)
            ElseIf (j + 16) Mod 32 = 0 Then Increment (4)
            ElseIf (j + 8) Mod 16 = 0 Then Increment (3)
            ElseIf (j + 4) Mod 8 = 0 Then Increment (2)
            ElseIf (j + 2) Mod 4 = 0 Then Increment (1)
            Else
                Increment (0)
            End If
        End If
    Next
    'Level 100: add another 32768. Don't know why, the normal algorithm generates a perfectly fine and unique code.
    'Level 100: 11775, closest would have been Level 26: 11782
    If HScroll1.Value = 100 Then x = x + 2 ^ 15
    Me.Caption = "Your level code is: " & x
End Sub

Private Sub Increment(IncPow As Integer)
    Dim NumberToAdd, FoundQuavers As Integer
    
    'Set normal increment
    NumberToAdd = 2 ^ (9 + IncPow)
    'Increment
    If Signs(IncPow) Then
        x = x + NumberToAdd
    Else
        x = x - NumberToAdd
    End If
    'Check if Quavers found at this level
    'I still can't figure out why the number of packs found is included in the code
    For FoundQuavers = 0 To 7
        If i = Quavers(FoundQuavers) Then
            x = x + 2 ^ (FoundQuavers + 1)
            Exit For
        End If
    Next
    'Change sign
    Signs(IncPow) = Not Signs(IncPow)
End Sub
