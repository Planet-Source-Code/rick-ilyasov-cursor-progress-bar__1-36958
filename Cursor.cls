VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get hInstance() As Long
    hInstance = g_hInstance
End Property

Public Property Let hInstance(ByVal vNewValue As Long)
    g_hInstance = vNewValue
End Property

Public Property Get Min() As Single
    Min = m_sngMin
End Property

Public Property Let Min(ByVal vNewValue As Single)
    m_sngMin = vNewValue
End Property

Public Property Get Max() As Single
    Max = m_sngMax
End Property

Public Property Let Max(ByVal vNewValue As Single)
    m_sngMax = vNewValue
End Property

Public Property Get Value() As Single
    Value = m_sngValue
    
End Property

Public Property Let Value(ByVal vNewValue As Single)

    DoEvents
    
    If vNewValue > m_sngMax Then
        m_sngValue = m_sngMax
    Else
        m_sngValue = vNewValue
        bReset = False
        Form1.tmrRender.Enabled = True
        If Not m_bActive Then
            m_bActive = True
            Form1.ClearBar
            Form1.tmrAnimator.Enabled = True
        End If
    End If
    If m_sngValue = m_sngMax Then
        bReset = True
        m_bActive = False
        Form1.tmrAnimator.Enabled = False
        Form1.tmrRender.Enabled = False
        bOldCursorSet = False
        If lOldCursor <> 0 Then
            SetCursor lOldCursor
        End If
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    m_bEnabled = vNewValue
    Form1.tmrRender.Enabled = True
End Property

Private Sub Class_Terminate()
    bReset = True
    m_bActive = False
    Form1.tmrAnimator.Enabled = False
End Sub
