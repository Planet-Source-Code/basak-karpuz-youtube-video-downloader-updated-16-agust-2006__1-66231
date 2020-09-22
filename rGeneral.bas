Attribute VB_Name = "rGeneral"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub OpenURL(URL$, ShowExplorer As Boolean)

    On Local Error GoTo OpenURLError

    Dim IExplorerPath$

    Screen.MousePointer = vbHourglass
    IExplorerPath = ReadReg(HCR, "Applications\iexplore.exe\shell\open\command", "")
    IExplorerPath = rGetString(IExplorerPath, Chr(34), Chr(34) + " ")
    Call ShellExecute(0&, "open", IExplorerPath, URL, vbNullString, IIf(ShowExplorer, 1, 0))
OpenURLError:
    Screen.MousePointer = vbDefault

End Sub

Public Function VideoDownloadURL$(YouTubeVideoURL$)

    On Local Error GoTo VideoDownloadURLError

    Dim VideoSource$, VideoParameters$

    Screen.MousePointer = vbHourglass
    VideoSource = rDownloadUrlSource(YouTubeVideoURL)
    VideoSource = rGetString(VideoSource, "SWFObject(" + Chr(34), Chr(34) + ",", 1, True) + "&s"
    VideoParameters = rGetString(VideoSource, "?", "&s", 1, True)
    VideoDownloadURL = "http://youtube.com/get_video.php?" + VideoParameters
VideoDownloadURLError:
    Screen.MousePointer = vbDefault

End Function
