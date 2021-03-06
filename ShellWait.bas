Attribute VB_Name = "ShellWait"
Option Explicit

'Permet de faire une pause dans le code: Sleep 5000 (pause de 5 secondes)
'(pour laisser le temps � un process DOS de s'executer par exemple)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'API de gestion de l'heure.
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'API d'ouverture de Process.
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'API de fermeture de Process.
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400

Public Function ShellAndWaitForTermination( _
        sShell As String, _
        Optional ByVal eWindowStyle As VBA.VbAppWinStyle = vbNormalFocus, _
        Optional ByRef sError As String, _
        Optional ByVal lTimeOut As Long = 3600 _
    ) As Boolean
Dim hProcess As Long
Dim lR As Long
Dim bSuccess As Boolean
Dim Second As Long
    
On Error GoTo ShellAndWaitForTerminationError
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sShell, eWindowStyle))
    If (hProcess = 0) Then
        'Impossible de lancer la ligne de commande!
        sError = "Le programme n'a pu �tre lanc�, v�rifiez votre ligne de commande."
    Else
        bSuccess = True
        Second = 0
        Do
            'R�cup�ration du statut du process,
            'on v�rifie s'il est termin� (lR = 0).
            GetExitCodeProcess hProcess, lR
            'Pause en attendant la fin de notre commande sans
            'g�ner l'execution des autres process.
            If Second <= lTimeOut Then
                DoEvents: Sleep 1000
                Second = Second + 1
            Else
                'Trop long!
                Call TerminateProcess(hProcess, lR)
                Call CloseHandle(hProcess)
                sError = "Trop long: Le process a �t� stopp�...."
                lR = 0
                bSuccess = False
            End If
        Loop While lR = STILL_ACTIVE
    End If
    ShellAndWaitForTermination = bSuccess
        
    Exit Function

ShellAndWaitForTerminationError:
    sError = Err.DESCRIPTION
    Exit Function
End Function

