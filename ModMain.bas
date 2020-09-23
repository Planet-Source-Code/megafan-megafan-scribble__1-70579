Attribute VB_Name = "ModMain"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public GrillePosibilite(14, 14) As String
Public Dico As ClsDictionnaire


Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Enum EPremierTour
    Aleatoire = 0
    Joueur = 1
    PC = 2
End Enum

Public Enum EMode
    Normal = 0
    Duplicate = 1
    Solo = 2
End Enum

Public Const KEYSCRIBBLE As String = "SOFTWARE\MegaFan\Scribble"
Public NIVEAU As Integer
Public PREMIERTOUR As EPremierTour
Public MAXMOTS As Long
Public POINTSDELETTRE() As Variant
Public STRPIOCHEDICO As String
Public MODEJEU As EMode
Public CHRONOMETRE As Long
' Pioche pour la partie en cours
Public STRPIOCHE As String
Public PARTIECHRONOMETREE As String

Public PAUSEAPRESCOUP As String
Public TEMPSPAUSE As Long

' Dictionnaire courant
Public STRDICO As String


Sub KeepOnTop(F As Form)

    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1

    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub
