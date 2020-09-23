Attribute VB_Name = "ModRegistry"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "ADVAPI32.DLL" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "ADVAPI32.DLL" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "ADVAPI32.DLL" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "ADVAPI32.DLL" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "ADVAPI32.DLL" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegDeleteKey Lib "ADVAPI32.DLL" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long



Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_DYN_DATA = &H80000004
Public Const KEY_ALL_ACCESS = &H3F
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type


'Const REG_EXPAND_SZ = 2
'Const REG_BINARY = 3

Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234
Const KEY_READ = &H20019

Public Function RegGetString(hKey As Long, StrPath As String, StrValue As String) As String

    Dim KeyHand As Long
    Dim lResult As Long
    Dim StrBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    
    If RegOpenKeyEx(hKey, StrPath, 0, KEY_ALL_ACCESS, KeyHand) = ERROR_SUCCESS Then
        lResult = RegQueryValueEx(KeyHand, StrValue, 0&, lValueType, ByVal 0&, lDataBufSize)
        If lValueType = REG_SZ Then
            StrBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(KeyHand, StrValue, 0&, 0&, ByVal StrBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                intZeroPos = InStr(StrBuf, Chr$(0))
                If intZeroPos > 0 Then
                    RegGetString = Left$(StrBuf, intZeroPos - 1)
                Else
                    RegGetString = StrBuf
                End If
            End If
        End If
        RegCloseKey (KeyHand)
    End If
    
End Function

Public Sub RegSaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)

    Dim KeyHand As Long
    Dim lResult As Long
    Dim lRetVal As Long
    Dim SA As SECURITY_ATTRIBUTES
    
    SA.nLength = 12&
    SA.lpSecurityDescriptor = 0&
    SA.bInheritHandle = False

    If RegOpenKeyEx(hKey, StrPath, 0, KEY_ALL_ACCESS, KeyHand) <> ERROR_SUCCESS Then
        lResult = RegCreateKeyEx(hKey, StrPath, 0&, vbNullString, 0&, KEY_ALL_ACCESS, SA, KeyHand, lRetVal)
        If lResult <> ERROR_SUCCESS Then Exit Sub
    End If
    
    lResult = RegSetValueEx(KeyHand, StrValue, 0, REG_SZ, ByVal StrData, CLng(Len(StrData) + 1))
    RegCloseKey (KeyHand)
    
End Sub

Public Function RegGetDword(ByVal hKey As Long, ByVal StrPath As String, ByVal strValueName As String) As Long

    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim KeyHand As Long

    If RegOpenKeyEx(hKey, StrPath, 0, KEY_ALL_ACCESS, KeyHand) = ERROR_SUCCESS Then
        lDataBufSize = 4
        lResult = RegQueryValueEx(KeyHand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            If lValueType = REG_DWORD Then
                RegGetDword = lBuf
            End If
        End If

        lResult = RegCloseKey(KeyHand)
    End If
    
End Function

Public Sub RegSaveDword(hKey As Long, StrPath As String, StrValue As String, LngData As Long)

    Dim KeyHand As Long
    Dim lResult As Long
    Dim lRetVal As Long
    Dim SA As SECURITY_ATTRIBUTES

    If RegOpenKeyEx(hKey, StrPath, 0, KEY_ALL_ACCESS, KeyHand) <> ERROR_SUCCESS Then
        SA.nLength = 12&
        SA.lpSecurityDescriptor = 0&
        SA.bInheritHandle = False
        
        lResult = RegCreateKeyEx(hKey, StrPath, 0&, vbNullString, 0&, KEY_ALL_ACCESS, SA, KeyHand, lRetVal)
        If lResult <> ERROR_SUCCESS Then Exit Sub
    End If
    
    lResult = RegSetValueEx(KeyHand, StrValue, 0, REG_DWORD, LngData, 4)
    RegCloseKey (KeyHand)
    
End Sub

