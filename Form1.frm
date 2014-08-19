VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "X-Y"
   ClientHeight    =   960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public cMonitorClass As New clsMonitors

Sub GetCursor()
Dim LonCStat As Long
    LonCStat = GetCursorPos&(m_CursorPos)
    'to use this result, the data must be converted into Pixel

    m_CursorPos.x = m_CursorPos.x
    m_CursorPos.y = m_CursorPos.y
End Sub



Private Sub Timer1_Timer()

    Dim oMon As clsMonitor
    Dim prop As Double
    Dim oCurr As clsMonitor

    GetCursor
    Label1.Caption = m_CursorPos.x & " x " & m_CursorPos.y
    
    'Moniteur courant
    Dim i As Integer
    Dim lMonitor As Long
    
    lMonitor = cMonitorClass.GetMonitorFromXYPoint(m_CursorPos.x, m_CursorPos.y, 1)
    For Each oCurr In cMonitorClass.Monitors
        If oCurr.Handle = lMonitor Then Exit For
    Next

    'Souris en haut
    If m_CursorPos.y = oCurr.Top Then
        'Recherche si écran au-dessus
        For Each oMon In cMonitorClass.Monitors
            If oMon.Bottom = oCurr.Top Then
                prop = oCurr.Width / oMon.Width
                SetCursorPos oMon.Left + (m_CursorPos.x - oCurr.Left) / prop, oMon.Bottom - 5
            End If
        Next
    ElseIf m_CursorPos.y = oCurr.Bottom - 1 Then
        'Recherche si écran au-dessus
        For Each oMon In cMonitorClass.Monitors
            If oMon.Top = oCurr.Bottom Then
                prop = oCurr.Width / oMon.Width
                SetCursorPos oMon.Left + (m_CursorPos.x - oCurr.Left) / prop, oMon.Top + 5
            End If
        Next
    ElseIf m_CursorPos.x = oCurr.Left Then
        'Recherche si écran au-dessus
        For Each oMon In cMonitorClass.Monitors
            If oMon.Right = oCurr.Left Then
                prop = oCurr.Height / oMon.Height
                SetCursorPos oMon.Right - 5, (m_CursorPos.y - oCurr.Top) / prop
            End If
        Next
    ElseIf m_CursorPos.x = oCurr.Right - 1 Then
        'Recherche si écran au-dessus
        For Each oMon In cMonitorClass.Monitors
            If oMon.Left = oCurr.Right Then
                prop = oCurr.Height / oMon.Height
                SetCursorPos oMon.Left + 5, (m_CursorPos.y - oCurr.Top) / prop
                Debug.Print oMon.Left + 5 & ", " & (m_CursorPos.y - oCurr.Top) / prop
            End If
        Next
    End If

End Sub
