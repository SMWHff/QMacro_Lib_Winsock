:'��������������������������������������Ҫ��������������޸ġ���������������������������������������������
:On Error Resume Next
:Sub bat
    echo off & cls
    echo '>nul&Title һ�����
    echo '>nul&set cDir=%~dp0
    echo '>nul&set cDir=%cDir:~,-1%
    echo '>nul&for /f "delims=" %%i in ("%cDir%") do set ProjectName=%%~ni
    echo '>nul&set ProjectName=������⡿�׽���
    echo '>nul&set SysDir=%SystemRoot%\System32
    echo '>nul&set SysBit=%PROCESSOR_ARCHITECTURE:~-2%
    echo '>nul&if %SysBit%==64 (set SysDir=%SystemRoot%\SysWOW64)
    echo '>nul&xcopy "%cDir%\lib\����_�׽���.html"   "%cDir%\source\%ProjectName%\��������2014\lib\"      /s /c /d /y
    echo '>nul&xcopy "%cDir%\lib\����_�׽���.qml"   "%cDir%\source\%ProjectName%\��������2014\lib\"      /s /c /d /y
    echo '>nul&xcopy "%cDir%\QMScript\!����_�׽���"    "%cDir%\source\%ProjectName%\��������2014\QMScript\!����_�׽���\" /s /c /d /y
    echo '>nul&%SysDir%\CScript.exe //nologo //E:vbscript "%~f0" "%ProjectName%" %*
    echo '>nul&explorer "%cDir%\source\"
    echo '>nul&echo �ű��Ѿ�ֹͣ����
    echo '>nul&pause>nul
    Exit Sub
End Sub

REM ������VBS����
'Set fso = CreateObject("Scripting.FileSystemObject")
'cd = fso.GetFile(wsh.ScriptFullName).ParentFolder.Path
'ProjectName = WScript.Arguments(0)
'sDir = cd & "\Releases\"& ProjectName &"\"
'ZipFile = cd & "\Releases\"& ProjectName &".zip"
'Call Zip(sDir, ZipFile)
'fso.DeleteFolder Left(sDir, Len(sDir)-1),True
'Set fso = Nothing
wsh.echo ""
wsh.echo "�����ɹ���"
wsh.echo ""


Sub Zip(ByVal mySourceDir, ByVal myZipFile)
    Dim fso,f,objShell,objTarget
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetExtensionName(myZipFile) <> "zip" Then
        Exit Sub
    ElseIf fso.FolderExists(mySourceDir) Then
        FType = "Folder"
    ElseIf fso.FileExists(mySourceDir) Then
        FType = "File"
        FileName = fso.GetFileName(mySourceDir)
        FolderPath = Left(mySourceDir, Len(mySourceDir) - Len(FileName))
    Else
        Exit Sub
    End If
    Set f = fso.CreateTextFile(myZipFile, True)
        f.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
        f.Close
    Set objShell = CreateObject("Shell.Application")
    Select Case Ftype
        Case "Folder"
            Set objSource = objShell.NameSpace(mySourceDir)
            Set objFolderItem = objSource.Items()
        Case "File"
            Set objSource = objShell.NameSpace(FolderPath)
            Set objFolderItem = objSource.ParseName(FileName)
    End Select
    Set objTarget = objShell.NameSpace(myZipFile)
    intOptions = 256
    objTarget.CopyHere objFolderItem, intOptions
    Do
        WScript.Sleep 1000
    Loop Until objTarget.Items.Count > 0
End Sub

Sub UnZip(ByVal myZipFile, ByVal myTargetDir)
    Dim fso,objShell,objSource,objTarget
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If NOT fso.FileExists(myZipFile) Then
        Exit Sub
    ElseIf fso.GetExtensionName(myZipFile) <> "zip" Then
        Exit Sub
    ElseIf NOT fso.FolderExists(myTargetDir) Then
        fso.CreateFolder(myTargetDir)
    End If
    Set objShell = CreateObject("Shell.Application")
    Set objSource = objShell.NameSpace(myZipFile)
    Set objFolderItem = objSource.Items()
    Set objTarget = objShell.NameSpace(myTargetDir)
    intOptions = 256
    objTarget.CopyHere objFolderItem, intOptions
End Sub