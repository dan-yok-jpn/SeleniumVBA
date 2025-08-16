Option Explicit

Private root_directory As String
Private library_file   As String
Private driver_file    As String

' ****************************************************************
'   https://tonari-it.com/excel-vba-windowsapi-urldownloadtofile/
' ****************************************************************
Private Declare PtrSafe Function URLDownloadToFile _
  Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
) As Long

' ***************************************************************************
' https://github.com/GCuser99/SeleniumVBA/blob/main/src/WebDriverManager.cls
' ***************************************************************************
Private Function requestData(ByVal Url As String, Optional maxTries As Long = 20) As Variant

    Dim try As Long

    On Error GoTo ext
    With CreateObject("MSXML2.ServerXMLHTTP")
        For try = 1 To maxTries
            .Open "GET", Url, False
            .send
            While .readyState <> 4
                DoEvents
            Wend
            Select Case .status
              Case 200
                requestData = .responseText
                Exit Function
              Case Else
                Exit For
            End Select
        Next try
    End With
    Exit Function
ext:
    requestData = ""
End Function

Private Function get_chrome_version() As String

    Dim path_app  As String
    Dim c         As String
    Dim subfolder As Object

    path_app = Environ("ProgramFiles") & "\Google\Chrome\Application"
    With CreateObject("Scripting.FileSystemObject")
        If .FolderExists(path_app) = False Then Exit Function
        For Each subfolder In .GetFolder(path_app).SubFolders
            c = Left(subfolder.Name, 1)
            If IsNumeric(c) Then
                get_chrome_version = subfolder.Name
                Exit Function
            End If
        Next
    End With

End Function

Private Function download_installer() As Boolean

    Dim res As Long
    Dim msg As String

    On Error GoTo ext
    res = URLDownloadToFile(0, _
        "https://github.com/GCuser99/SeleniumVBA/raw/main/dist/SeleniumVBADLLSetup.exe", _
        ThisWorkbook.Path & "\SeleniumVBADLLSetup.exe", 0, 0)
    If res = 0 Then
        msg = "ライブラリのインストーラをダウンロードしました"
        msg = msg & vbCrLf & "インストールを実行して下さい"
        msg = msg & vbCrLf & "（インストール先はクリップボード在中の"
        msg = msg & vbCrLf & vbTab & root_directory
        msg = msg & vbCrLf & "を指定）"
        msg = msg & vbCrLf & "実行後、このワークブックを再起動してください"
        MsgBox msg
        download_installer = True ' MSForms.DataObject
        With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            .SetText root_directory
            .PutInClipboard
        End With
        Exit Function
    End If
ext:
    download_installer = False
End Function

Private Function get_installed_libraly_version() As String

    Dim version As Variant

    On Error GoTo ext
    With CreateObject("Scripting.FileSystemObject")
        version = Split(.GetFileVersion(library_file), ".")
        get_installed_libraly_version = version(0) & "." & version(1)
        Exit Function
    End With
ext:
    get_installed_libraly_version = ""
End Function

Private Function get_recent_libraly_version() As String

    Dim Url          As String
    Dim responseText As String
    Dim json         As New Scripting.Dictionary
    Dim tag_name     As Variant
Option Explicit

Private root_directory As String
Private library_file   As String
Private driver_file    As String

' ****************************************************************
'   https://tonari-it.com/excel-vba-windowsapi-urldownloadtofile/
' ****************************************************************
Private Declare PtrSafe Function URLDownloadToFile _
  Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
) As Long

' ***************************************************************************
' https://github.com/GCuser99/SeleniumVBA/blob/main/src/WebDriverManager.cls
' ***************************************************************************
Private Function requestData(ByVal Url As String, Optional maxTries As Long = 20) As Variant

    Dim try As Long

    On Error GoTo ext
    With CreateObject("MSXML2.ServerXMLHTTP")
        For try = 1 To maxTries
            .Open "GET", Url, False
            .send
            While .readyState <> 4
                DoEvents
            Wend
            Select Case .status
              Case 200
                requestData = .responseText
                Exit Function
              Case Else
                Exit For
            End Select
        Next try
    End With
    Exit Function
ext:
    requestData = ""
End Function

Private Function get_chrome_version() As String

    Dim path_app  As String
    Dim c         As String
    Dim subfolder As Object

    path_app = Environ("ProgramFiles") & "\Google\Chrome\Application"
    With CreateObject("Scripting.FileSystemObject")
        If .FolderExists(path_app) = False Then Exit Function
        For Each subfolder In .GetFolder(path_app).SubFolders
            c = Left(subfolder.Name, 1)
            If IsNumeric(c) Then
                get_chrome_version = subfolder.Name
                Exit Function
            End If
        Next
    End With

End Function

Private Function download_installer() As Boolean

    Dim res As Long
    Dim msg As String

    On Error GoTo ext
    res = URLDownloadToFile(0, _
        "https://github.com/GCuser99/SeleniumVBA/raw/main/dist/SeleniumVBADLLSetup.exe", _
        ThisWorkbook.Path & "\SeleniumVBADLLSetup.exe", 0, 0)
    If res = 0 Then
        msg = "ライブラリのインストーラをダウンロードしました"
        msg = msg & vbCrLf & "インストールを実行して下さい"
        msg = msg & vbCrLf & "（インストール先はクリップボード在中の"
        msg = msg & vbCrLf & vbTab & root_directory
        msg = msg & vbCrLf & "を指定）"
        msg = msg & vbCrLf & "実行後、このワークブックを再起動してください"
        MsgBox msg
        download_installer = True ' MSForms.DataObject
        With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            .SetText root_directory
            .PutInClipboard
        End With
        Exit Function
    End If
ext:
    download_installer = False
End Function

Private Function get_installed_libraly_version() As String

    Dim version As Variant

    On Error GoTo ext
    With CreateObject("Scripting.FileSystemObject")
        version = Split(.GetFileVersion(library_file), ".")
        get_installed_libraly_version = version(0) & "." & version(1)
        Exit Function
    End With
ext:
    get_installed_libraly_version = ""
End Function

Private Function get_recent_libraly_version() As String

    Dim Url          As String
    Dim responseText As String
    Dim json         As New Scripting.Dictionary
    Dim tag_name     As Variant

    Url = "https://api.github.com/repos/GCuser99/SeleniumVBA/releases/latest"
    responseText = requestData(Url)
    If Not IsEmpty(responseText) Then
        With CreateObject("SeleniumVBA.WebJsonConverter")
            Set json = .ParseJson(responseText)
            tag_name = Split(json("tag_name"), ".")
            get_recent_libraly_version = Replace(tag_name(0), "v", "") & "." & tag_name(1)
            Set json = Nothing
        End With
    End If

End Function

Private Sub Workbook_Open()

    Dim chrome_version    As String
    Dim MsgBoxResult      As VbMsgBoxResult
    Dim installed_version As String
    Dim recent_version    As String
    Dim driver_version    As String

    root_directory = Environ("LOCALAPPDATA") & "\SeleniumVBA\"
    library_file = root_directory & "\SeleniumVBA_win64.dll"
    driver_file = root_directory & "\chromedriver.exe"

    chrome_version = get_chrome_version()
    If chrome_version = "" Then
        MsgBox "Google Chrome がインストールされていません" & _
                    vbCrLf & "終了します"
        Exit Sub
    End If

    If Dir(library_file) = "" Then
        MsgBoxResult = MsgBox("SeleniumVBA がインストールされていません" & _
                        vbCrLf & "インストーラをダウンロードしますか？", vbYesNo)
        If MsgBoxResult = vbYes Then
            If download_installer() = False Then
                MsgBox "インストーラのダウンロードに失敗しました" & vbCrLf & "終了します"
                Exit Sub
            Else
                Application.Quit
            End If
        End If
    Else
        On Error Resume Next
        ThisWorkbook.VBProject.References.AddFromFile library_file
        If Not Err.Number = 0 Then
            MsgBox "参照設定に失敗しました"
            Exit Sub
        End If
        installed_version = get_installed_libraly_version()
        recent_version = get_recent_libraly_version()
        If (Not IsEmpty(recent_version)) And (Not installed_version = recent_version) Then
            MsgBoxResult = MsgBox("SeleniumVBA の更新版があります" & _
                            vbCrLf & "更新版のインストーラをダウンロードしますか？", vbYesNo)
            If MsgBoxResult = vbYes Then
                If download_installer() = False Then
                    MsgBox "インストーラのダウンロードに失敗しました" & vbCrLf & "終了します"
                    Exit Sub
                Else
                    Application.Quit
                End If
            End If
        End If
    End If

    With CreateObject("SeleniumVBA.WebDriverManager")
        If Dir(driver_file) = "" Then
            driver_version = "000.0.000.00"
        Else
            driver_version = .GetInstalledDriverVersion(1, driver_file)
        End If
        If .CheckCompatibilityLevel(chrome_version, driver_version) < 3 Then
            recent_version = .GetCompatibleDriverVersion(1, chrome_version)
            On Error Resume Next
            .DownloadAndInstallDriver 1, recent_version, driver_file
            If Not Err.Number = 0 Then
                MsgBox "ChromeDriver の更新に失敗しました"
            End If
        End If
    End With
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error Resume Next
    With ThisWorkbook.VBProject
        .References.Remove .References("SeleniumVBA")
    End With
    On Error GoTo 0

End Sub
