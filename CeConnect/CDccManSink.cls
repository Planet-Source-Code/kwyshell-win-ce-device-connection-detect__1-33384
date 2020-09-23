VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDccManSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objDccMan As CDccMan
Private m_lContext As Long

Implements IDccManSink

Private Sub Class_Initialize()
    Set m_objDccMan = New CDccMan
    m_objDccMan.Advise Me, m_lContext
End Sub

Private Sub Class_Terminate()
    m_objDccMan.Unadvise m_lContext
    Set m_objDccMan = Nothing
End Sub

Public Sub ShowCommSettings()
    m_objDccMan.ShowCommSettings
End Sub

Private Sub IDccManSink_OnLogActive()
    MsgBox "OnLogActive", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogAnswered()
    MsgBox "OnLogAnswered", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogDisconnection()
    MsgBox "OnLogDisconnection", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogError()
    MsgBox "OnLogError", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogInactive()
    MsgBox "OnLogInactive", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogIpAddr(ByVal dwIpAddr As Long)
    MsgBox "OnLogIpAddr IP: 0x" & FormateStr(dwIpAddr)
End Sub

Private Sub IDccManSink_OnLogListen()
    MsgBox "OnLogListen", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogTerminated()
    MsgBox "OnLogTerminated", vbOKOnly, "Device Connection Detector"
End Sub

Private Function FormateStr(lAddress As Long)

    Const clOffSet As Long = &H100
    Dim str As String
    Do
    str = Hex(lAddress Mod clOffSet) + str
    lAddress = lAddress \ clOffSet
    Loop While lAddress > 0

End Function
