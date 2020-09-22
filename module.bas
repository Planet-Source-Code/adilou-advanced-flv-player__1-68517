Attribute VB_Name = "Module2"
'Different functions

Public Sub Association(EXT As String, FileType As String, FileName As String)
 On Error Resume Next
 Dim b As Object
 Set b = CreateObject("wscript.shell")
 b.regwrite "HKCR\" & EXT & "\", FileType
 b.regwrite "HKCR\" & FileType & "\", "MY file"
 b.regwrite "HKCR\" & FileType & "\DefaultIcon\", FileName
 b.regwrite "HKCR\" & FileType & "\shell\open\command\", FileName & " %L"
 b.regdelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & EXT & "\Application"
 b.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & EXT & "\Application", FileName
 b.regdelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & EXT & "\OpenWithList\"
 b.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & EXT & "\OpenWithList\a", FileName
 
 End Sub

Function lignecommande(sNomParametre As String) As String

Dim iPos As Integer
Dim iposFin As Integer

' on recherche le nom de parametre
iPos = InStr(Command$, sNomParametre)
If iPos > 0 Then
iposFin = InStr(iPos, Command$, " ")
If iposFin > 0 Then
' on a un blanc derriere le parametre
'lignecommande = Mid$(command$, iPos + Len(sNomParametre), iposFin - iPos - Len(sNomParametre))
lignecommande = Mid$(Command$, iPos + Len(sNomParametre))
Else
' on est a la fin de la ligne de parametres et l'on n'a pas de blancs
lignecommande = Mid$(Command$, iPos + Len(sNomParametre))
End If
Else
lignecommande = ""
End If
End Function


