Attribute VB_Name = "connected"
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255

Public Function Conectado() As Boolean
    Dim ret As Long
    ret = InternetGetConnectedStateEx(ret, sConnType, 254, 0)
    If ret = 1 Then
        Conectado = True
    Else
        Conectado = False
    End If
End Function

