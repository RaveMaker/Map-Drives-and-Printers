' Map Drives
'
' by RaveMaker - http://ravemaker.net

Option Explicit

Dim objNetwork, objFSO, wshNetwork

Set objNetwork = CreateObject("Wscript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshNetwork = CreateObject("WScript.Network")

'*****************************************************************************************
' Global Domain Users Mapping a network drive + Users Home folder

    If (MapDrive("m:", "\\10.0.0.1\Media") = False) Then
        MsgBox "Unable to Map M: to Media"
    End If

'*******************************************************************************************
Function MapDrive(ByVal strDrive, ByVal strShare)
    ' Map network share to a drive letter.

    Dim objDrive

    On Error Resume Next
    If (objFSO.DriveExists(strDrive) = True) Then
        Set objDrive = objFSO.GetDrive(strDrive)
        If (Err.Number <> 0) Then
            On Error GoTo 0
            MapDrive = False
            Exit Function
        End If
        If (objDrive.DriveType = 3) Then
            objNetwork.RemoveNetworkDrive strDrive, True, True
        Else
            MapDrive = False
            Exit Function
        End If
        Set objDrive = Nothing
    End If
    objNetwork.MapNetworkDrive strDrive, strShare
    If (Err.Number = 0) Then
        MapDrive = True
    Else
        Err.Clear
        MapDrive = False
    End If
    On Error GoTo 0
End Function
