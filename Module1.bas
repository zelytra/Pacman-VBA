Attribute VB_Name = "Module1"
Global chemin As String
Global son As String
Global direction As String
Global last_direction As String
Global moove As Integer
Global ValueDeplacement As Double
Global boucle As Integer
Global score As Integer
Global niv As Integer
Global cellule As String

Global Pacman_Cell As Range
Global RGhost_Cell As Range


Public Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
ByValuReturnLength As Long, ByVal hwndCallback As Long) As Long _

Public Declare PtrSafe Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As _
String, ByVal cchBuffer As Long) As Long

Public Sub reset()
Range("DX33:HE126").Select
    Selection.Copy
    Range("AC33").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Public Sub Deplacement()

If direction = "H" Then
'Call Jeux.Haut_GhostR
Call Jeux.Haut
ElseIf direction = "B" Then
'Call Jeux.Bas_GhostR
Call Jeux.Bas
ElseIf direction = "G" Then
'Call Jeux.Gauche_GhostR
Call Jeux.Gauche
ElseIf direction = "D" Then
'Call Jeux.Droite_GhostR
Call Jeux.Droite
End If

last_direction = direction

Call Jeux.point
Call Jeux.level

Call tempo
If boucle = 1 Then
Exit Sub
End If


Debug.Print Now
Application.OnTime Now, "Deplacement", , True
End Sub
Sub tempo()
Dim start As Double
Dim pause As Double
pause = 0.065

start = timer

Do While timer < start + pause
    DoEvents
Loop
End Sub

Sub tempo_start()
Dim start As Double
Dim pause As Double
pause = 4

start = timer

Do While timer < start + pause
    DoEvents
Loop

End Sub


