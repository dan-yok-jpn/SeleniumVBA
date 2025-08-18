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

Private Sub UpdateIniFileKey( _
        filePath As String, _
        targetSection As String, _
        targetKey As String, _
        newValue As String _
    )

    Dim fileContent   As String
    Dim newContent    As String

    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        MsgBox "file not found: " & filePath
        Exit Sub
    End If
    
    Dim tsIn    As Object
    Dim lines() As String
    Set tsIn = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    fileContent = tsIn.ReadAll
    tsIn.Close
    lines = Split(fileContent, vbCrLf)
    
    Dim sectionFound As Boolean
    Dim keyUpdated As Boolean
    Dim i As Long
    
    sectionFound = False
    keyUpdated = False
    
    For i = LBound(lines) To UBound(lines)
        Dim currentLine As String
        currentLine = Trim(lines(i))
        
        If Left(currentLine, 1) = "[" And Right(currentLine, 1) = "]" Then
            Dim sectionName As String
            sectionName = Mid(currentLine, 2, Len(currentLine) - 2)
            If sectionName = targetSection Then
                sectionFound = True
            Else
                sectionFound = False
            End If
        ElseIf sectionFound And InStr(currentLine, "=") > 0 Then
            Dim keyValuePair() As String
            Dim currentKey     As String
            Dim currentValue   As String
            keyValuePair = Split(currentLine, "=", 2)
            currentKey = Trim(keyValuePair(0))
            If currentValue = newValue Then Exit Sub
            If currentKey = targetKey Then
                lines(i) = currentKey & "=" & newValue ' update value
                keyUpdated = True
                Exit For
            End If
        End If
    Next i

    If Not keyUpdated Then
        MsgBox "not found section or key that you specified"
        Exit Sub
    End If

    newContent = Join(lines, vbCrLf)

    Dim tsOut As Object
    Set tsOut = fso.CreateTextFile(filePath, True) ' True = overwrite
    tsOut.Write newContent
    tsOut.Close
    
    Exit Sub
    
ErrorHandler:
    MsgBox "raised unsuspected error: " & Err.Description
End Sub

Private Sub Workbook_Open()

    Dim chrome_version    As String
    Dim MsgBoxResult      As VbMsgBoxResult
    Dim installed_version As String
    Dim recent_version    As String

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

    Call UpdateIniFileKey( _
        root_directory & "\SeleniumVBA.ini", _
        "GENERAL", _
        "driver_location_folder", _
        root_directory _
    )

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error Resume Next
    With ThisWorkbook.VBProject
        .References.Remove .References("SeleniumVBA")
    End With
    On Error GoTo 0

End Sub