VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimator 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3000
      Top             =   960
   End
   Begin VB.Timer tmrRender 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   555
      Left            =   1680
      TabIndex        =   2
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CDL 
      Left            =   960
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   60
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hNewCursor As Long
Dim hOldCursor As Long
Dim AndBuffer As String
Dim XorBuffer As String
Dim aAndBits(0 To 127) As Byte
Dim aXorBits(0 To 127) As Byte
Dim I As Integer
Dim lRet As Long



Public Sub AnimateCursor()
    
    m_bActive = True
    
    Do Until bReset

    Loop
    
End Sub

Private Function ReadAndBuffer() As String
    
'Rules
    
'AND mask XOR mask Display
'0   0    Black
'0   1    White
'1   0    Screen
'1   1    Reverse screen

    Dim X As Integer
    Dim Y As Integer
    Dim lCol As Long
    Dim sRet As String
    Dim sHex As String
    Dim I As Integer
    
    For Y = 0 To 31
        For X = 0 To 31
            lCol = picSample.Point(X, Y)
            Select Case lCol
                Case vbBlack, vbWhite
                    sRet = sRet & "0"
                Case vbGreen, vbRed
                    sRet = sRet & "1"
                Case Else
                    Stop
            End Select
            
        Next
    Next
        
    ReadAndBuffer = sRet
    
End Function

Private Function ReadXorBuffer() As String
    
'Rules
    
'AND mask XOR mask Display
'0   0    Black
'0   1    White
'1   0    Screen
'1   1    Reverse screen

    Dim X As Integer
    Dim Y As Integer
    Dim lCol As Long
    Dim sRet As String
    Dim sHex As String
    Dim I As Integer
    
    For Y = 0 To 31
        For X = 0 To 31
            lCol = picSample.Point(X, Y)
            Select Case lCol
                Case vbBlack, vbGreen
                    sRet = sRet & "0"
                Case vbWhite, vbRed
                    sRet = sRet & "1"
            End Select
        Next
    Next
    
    ReadXorBuffer = sRet
    
End Function

Private Sub Command2_Click()

    bReset = True
    
End Sub

Private Sub Command3_Click()

    CDL.ShowOpen
    
        
End Sub

Private Function Bin2Dec(BinaryValue As Variant) As Variant
    
    Dim Bit As Integer
    Dim Value As Integer
    Dim Counter As Integer

    For Counter = Len(BinaryValue) To 1 Step -1
        Bit = Mid(BinaryValue, Counter, 1)
        If Bit = 1 Then
            Value = Value + 2 ^ (Len(BinaryValue) - Counter)
        End If
    Next
    
    Bin2Dec = Value
    
End Function

Sub ShowCursor(AndBuffer As String, XorBuffer As String)
    
    For I = 0 To 127
      aAndBits(I) = Bin2Dec(Mid(AndBuffer, 8 * I + 1, 8))
      aXorBits(I) = Bin2Dec(Mid(XorBuffer, 8 * I + 1, 8))
    Next
    
    hNewCursor = CreateCursor(g_hInstance, 0, 0, 32, 32, aAndBits(0), aXorBits(0))
    
    hOldCursor = SetCursor(hNewCursor)  ' change cursor
    If Not bOldCursorSet Then
        bOldCursorSet = True
        lOldCursor = hOldCursor
    End If
    
    lRet = DestroyCursor(hNewCursor)

End Sub

Private Sub tmrAnimator_Timer()
    
    AndBuffer = ReadAndBuffer
    XorBuffer = ReadXorBuffer
    ShowCursor AndBuffer, XorBuffer
    DoEvents

End Sub

Public Sub ClearBar()
    
    Dim X As Integer
    
    'Reset the bar
    For X = 8 To 31
        picSample.Line (X, 22)-(X, 26), vbGreen
    Next
    
    'Clear Digits
    For X = 15 To 25
        picSample.Line (X, 27)-(X, 32), vbGreen
    Next
    

End Sub
Private Sub tmrRender_Timer()

    Dim X As Single
    Dim Y As Single
    Dim sngRange As Single
    Dim sngPct As Single
    Dim pctVal As String
    
    'Reset percent
    For X = 15 To 25
        picSample.Line (X, 27)-(X, 31), vbGreen
    Next

    'Draw the progress bar
        
    sngRange = m_sngMax - m_sngMin
    If sngRange > 0 Then
        sngPct = sngRange / 24
        If m_sngValue >= 0 And m_sngValue <= m_sngMax Then
            For X = 0 To m_sngValue / sngPct + 2
                picSample.Line (X + 8, 22)-(X + 8, 26), vbRed
            Next
        End If
    End If
    
    'Draw Percent Value
    pctVal = CInt(m_sngValue / (sngRange / 100)) + 1
    
    Select Case Len(pctVal)
        Case 3
            DrawDigit Mid(pctVal, 1, 1), 15, 27
            DrawDigit Mid(pctVal, 2, 1), 19, 27
            DrawDigit Mid(pctVal, 3, 1), 23, 27
        Case 2
            DrawDigit Mid(pctVal, 1, 1), 19, 27
            DrawDigit Mid(pctVal, 2, 1), 23, 27
        Case 1
            DrawDigit Mid(pctVal, 1, 1), 23, 27
    End Select
    
End Sub

Private Sub DrawDigit(iDigit As Integer, X As Integer, Y As Integer)

    Select Case iDigit
        Case 0
            SetPixels X, Y, "111"
            SetPixels X, Y + 1, "1 1"
            SetPixels X, Y + 2, "1 1"
            SetPixels X, Y + 3, "1 1"
            SetPixels X, Y + 4, "111"
                        
        Case 1
            SetPixels X, Y, "11 "
            SetPixels X, Y + 1, " 1 "
            SetPixels X, Y + 2, " 1 "
            SetPixels X, Y + 3, " 1 "
            SetPixels X, Y + 4, "111"
        
        Case 2
            SetPixels X, Y, "111"
            SetPixels X, Y + 1, "  1"
            SetPixels X, Y + 2, "111"
            SetPixels X, Y + 3, "1  "
            SetPixels X, Y + 4, "111"
        
        Case 3
            SetPixels X, Y, "111"
            SetPixels X, Y + 1, "  1"
            SetPixels X, Y + 2, "111"
            SetPixels X, Y + 3, "  1"
            SetPixels X, Y + 4, "111"
        
        Case 4
            SetPixels X, Y, "1 1"
            SetPixels X, Y + 1, "1 1"
            SetPixels X, Y + 2, "111"
            SetPixels X, Y + 3, "  1"
            SetPixels X, Y + 4, "  1"
        
        Case 5
            SetPixels X, Y, "111"
            SetPixels X, Y + 1, "1  "
            SetPixels X, Y + 2, "111"
            SetPixels X, Y + 3, "  1"
            SetPixels X, Y + 4, "111"

        Case 6
            SetPixels X, Y, "111"
            SetPixels X, Y + 1, "1  "
            SetPixels X, Y + 2, "111"
            SetPixels X, Y + 3, "1 1"
            SetPixels X, Y + 4, "111"
        
        Case 7
            SetPixels X, Y, "111"
            SetPixels X, Y + 1, "  1"
            SetPixels X, Y + 2, "  1"
            SetPixels X, Y + 3, "  1"
            SetPixels X, Y + 4, "  1"
        
        Case 8
            SetPixels X, Y, "111"
            SetPixels X, Y + 1, "1 1"
            SetPixels X, Y + 2, "111"
            SetPixels X, Y + 3, "1 1"
            SetPixels X, Y + 4, "111"
        
        Case 9
            SetPixels X, Y, "111"
            SetPixels X, Y + 1, "1 1"
            SetPixels X, Y + 2, "111"
            SetPixels X, Y + 3, "  1"
            SetPixels X, Y + 4, "111"
        
        
    
    End Select

End Sub

Private Sub SetPixels(X As Integer, Y As Integer, BitMask As String)
    
    Dim I As Integer
    
    For I = 1 To Len(BitMask)
        If Mid(BitMask, I, 1) = "1" Then
            picSample.PSet (X + I - 1, Y), vbRed
        Else
            picSample.PSet (X + I - 1, Y), vbGreen
        End If
    Next
    
End Sub
