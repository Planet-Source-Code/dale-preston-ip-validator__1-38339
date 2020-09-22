VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Validate"
      Height          =   372
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1812
   End
   Begin VB.Label Label1 
      Caption         =   "Enter IP Address:"
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2052
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************
'IP Address Validator for dotted decimal IP addresses
'Copyright Dale Preston, 2002
'***************************

Private Sub command1_click()
MsgBox CheckIP(Text1.Text)
End Sub

Public Function CheckIP(strIPaddress As String) As Boolean
    Dim strOctet As String
    Dim x As Integer
    For x = 0 To 2
        strOctet = Left(strIPaddress, InStr(1, strIPaddress, ".") - 1)
        strIPaddress = Mid(strIPaddress, InStr(1, strIPaddress, ".") + 1)
        If Not CheckOctet(strOctet) Then
            CheckIP = False
            Exit Function
        End If
    Next
    If Not CheckOctet(strIPaddress) Then
        CheckIP = False
        Exit Function
    End If
    
    CheckIP = True
End Function

Public Function CheckOctet(strOctet As String) As Boolean
    Select Case Len(strOctet)
        Case 1
            If Val(strOctet) < 0 Or Val(strOctet) > 9 Then
                CheckOctet = False
                Exit Function
            End If
        Case 2
            If Val(strOctet) < 10 Or Val(strOctet) > 99 Then
                CheckOctet = False
                Exit Function
            End If
        Case 3
            If Val(strOctet) < 100 Or Val(strOctet) > 255 Then
                CheckOctet = False
                Exit Function
            End If
        Case Else
            CheckOctet = False
            Exit Function
    End Select
    If IsNumeric(strOctet) Then CheckOctet = True
End Function
