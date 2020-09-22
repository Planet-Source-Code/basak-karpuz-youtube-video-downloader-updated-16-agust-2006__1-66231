Attribute VB_Name = "rWebOperationsModule"
Option Explicit

'This module is completely coded by Ramci
'my email is ramci_geliyo@hotmail.com

Private Const IF_FROM_CACHE As Long = &H1000000
Private Const IF_MAKE_PERSISTENT As Long = &H2000000
Private Const IF_NO_CACHE_WRITE As Long = &H4000000
Private Const INTERNET_OPEN_TYPE_DIRECT As Long = 1
      
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'URL = http://www.x.x[/x][/x.x]

Public Function rDownloadFile(ByVal URL$, ByVal SaveFilePathName$) As Boolean

    On Local Error Resume Next

    If LCase(Left(URL, Len("http://"))) <> "http://" Then URL = "http://" + URL
    rDownloadFile = URLDownloadToFile(0, URL, SaveFilePathName, 0, 0) = 0

End Function

Public Function rDownloadUrlSource$(ByVal URL$)

    Const BUFFER_LEN As Long = 1024
    Dim hInternet&, hFile&, lReturn&, sBuffer As String * BUFFER_LEN

    On Local Error GoTo URLSourceError

    If LCase(Left(URL, Len("http://"))) <> "http://" Then URL = "http://" + URL
    hInternet = InternetOpen("VBURLSource", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hInternet = 0 Then GoTo URLSourceError
    hFile = InternetOpenUrl(hInternet, URL, vbNullString, ByVal 0&, IF_NO_CACHE_WRITE, ByVal 0&)
    If hFile = 0 Then GoTo URLSourceError
    Call InternetReadFile(hFile, sBuffer, BUFFER_LEN, lReturn)
    rDownloadUrlSource = Left(sBuffer, lReturn)
    While Not lReturn = 0
        Call InternetReadFile(hFile, sBuffer, BUFFER_LEN, lReturn)
        rDownloadUrlSource = rDownloadUrlSource + Left(sBuffer, lReturn)
    Wend
    Call InternetCloseHandle(hInternet)
    Exit Function

URLSourceError:
    rDownloadUrlSource = vbNullString
    If hInternet Then Call InternetCloseHandle(hInternet)

End Function
