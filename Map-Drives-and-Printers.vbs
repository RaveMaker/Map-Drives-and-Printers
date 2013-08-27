' Map Drives and Printers according to AD group memberships
'
' by RaveMaker - http://ravemaker.net

Option Explicit

Dim objNetwork, objSysInfo, strUserDN
Dim objGroupList, objUser, objFSO
Dim strComputerDN, objComputer
Dim wshNetwork

Set objNetwork = CreateObject("Wscript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")
Set wshNetwork = CreateObject("WScript.Network")

strUserDN = objSysInfo.userName
strComputerDN = objSysInfo.computerName

' Escape any forward slash characters, "/", with the backslash escape character.
strUserDN = Replace(strUserDN, "/", "\/")
strComputerDN = Replace(strComputerDN, "/", "\/")

' Bind to the user and computer objects with the LDAP provider.
Set objUser = GetObject("LDAP://" & strUserDN)
Set objComputer = GetObject("LDAP://" & strComputerDN)


'*****************************************************************************************
' Global Domain Users Mapping a network drive + Users Home folder

    If (MapDrive("m:", "\\server\Public") = False) Then
        MsgBox "Unable to Map M: to Public"
    End If

    If (MapDrive("p:", "\\server\Private\" & wshNetwork.UserName) = False) Then
        MsgBox "Unable to Map P: to Private"
    End If

'*****************************************************************************************
' Map a network drive if the user is a member of the group.

If (IsMember(objUser, "Administrators") = True) Then
    If (MapDrive("s:", "\\server\Admin") = False) Then
        MsgBox "Unable to Map S: to Admin"
    End If
End If

'*******************************************************************************************

' Add a network printer if the computer is a member of the group.
' Make this printer the default.

objNetwork.AddWindowsPrinterConnection "\\server\PublicPrinter"
objNetwork.SetDefaultPrinter "\\server\PublicPrinter"

If (IsMember(objComputer, "Admins Computers") = True) Then
    objNetwork.AddWindowsPrinterConnection "\\server\AdminPrinter"
    objNetwork.SetDefaultPrinter "\\server\AdminPrinter"
End If

'*******************************************************************************************

Function IsMember(ByVal objADObject, ByVal strGroup)
    ' Function to test for group membership.
    ' objGroupList is a dictionary object with global scope.

    If (IsEmpty(objGroupList) = True) Then
        Set objGroupList = CreateObject("Scripting.Dictionary")
    End If
    If (objGroupList.Exists(objADObject.sAMAccountName & "\") = False) Then
        Call LoadGroups(objADObject, objADObject)
        objGroupList.Add objADObject.sAMAccountName & "\", True
    End If
    IsMember = objGroupList.Exists(objADObject.sAMAccountName & "\" _
        & strGroup)
End Function


'*******************************************************************************************

Sub LoadGroups(ByVal objPriObject, ByVal objADSubObject)
    ' Recursive subroutine to populate dictionary object objGroupList.

    Dim colstrGroups, objGroup, j

    objGroupList.CompareMode = vbTextCompare
    colstrGroups = objADSubObject.memberOf
    If (IsEmpty(colstrGroups) = True) Then
        Exit Sub
    End If
    If (TypeName(colstrGroups) = "String") Then
        ' Escape any forward slash characters, "/", with the backslash
        ' escape character. All other characters that should be escaped are.
        colstrGroups = Replace(colstrGroups, "/", "\/")
        Set objGroup = GetObject("LDAP://" & colstrGroups)
        If (objGroupList.Exists(objPriObject.sAMAccountName & "\" _
                & objGroup.sAMAccountName) = False) Then
            objGroupList.Add objPriObject.sAMAccountName & "\" _
                & objGroup.sAMAccountName, True
            Call LoadGroups(objPriObject, objGroup)
        End If
        Set objGroup = Nothing
        Exit Sub
    End If
    For j = 0 To UBound(colstrGroups)
        ' Escape any forward slash characters, "/", with the backslash
        ' escape character. All other characters that should be escaped are.
        colstrGroups(j) = Replace(colstrGroups(j), "/", "\/")
        Set objGroup = GetObject("LDAP://" & colstrGroups(j))
        If (objGroupList.Exists(objPriObject.sAMAccountName & "\" _
                & objGroup.sAMAccountName) = False) Then
            objGroupList.Add objPriObject.sAMAccountName & "\" _
                & objGroup.sAMAccountName, True
            Call LoadGroups(objPriObject, objGroup)
        End If
    Next
    Set objGroup = Nothing
End Sub

'*******************************************************************************************

Function MapDrive(ByVal strDrive, ByVal strShare)
    ' Function to map network share to a drive letter.
    ' If the drive letter specified is already in use, the function
    ' attempts to remove the network connection.
    ' objFSO is the File System Object, with global scope.
    ' objNetwork is the Network object, with global scope.
    ' Returns True if drive mapped, False otherwise.

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
