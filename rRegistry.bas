Attribute VB_Name = "rRegistry"
Option Explicit

Public Const HCR& = &H80000000
Private Const REG_SZ& = 1

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Public Function ReadReg$(hKey&, lpSubKey$, lpValueName$)

    Dim nBuffer&

    ReadReg = String(256, vbNullChar)
    Call RegOpenKey(hKey, lpSubKey, nBuffer)
    Call RegQueryValueEx(nBuffer, lpValueName, 0, REG_SZ, ByVal ReadReg, Len(ReadReg))
    Call RegCloseKey(nBuffer)
    ReadReg = Left$(ReadReg, InStr(1, ReadReg, Chr$(0)))

End Function

Public Sub WriteReg(hKey&, lpSubKey$, lpValueName$, lpValue$)

    Dim nBuffer&

    Call RegCreateKey(hKey, lpSubKey, nBuffer)
    Call RegSetValueEx(nBuffer, lpValueName, 0, REG_SZ, ByVal lpValue, Len(lpValue))
    Call RegCloseKey(nBuffer)

End Sub

