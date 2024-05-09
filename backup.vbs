Option Explicit

Dim objFSO, objShell, sourcePath, backupFolder, backupPath

' ファイルシステムオブジェクトの作成
Set objFSO = CreateObject("Scripting.FileSystemObject")

' シェルオブジェクトの作成
Set objShell = CreateObject("WScript.Shell")

' パラメータからソースパスを取得
If WScript.Arguments.Count = 0 Then
    WScript.Echo "Usage: CScript backup.vbs <sourcePath>"
    WScript.Quit
Else
    sourcePath = WScript.Arguments(0)
End If

' ソースパスが存在するかどうかを確認
If Not objFSO.FileExists(sourcePath) And Not objFSO.FolderExists(sourcePath) Then
    WScript.Echo "Source path does not exist."
    WScript.Quit
End If

' バックアップ先のフォルダ名を生成
Dim backupDate
backupDate = Year(Now) & Right("00" & Month(Now), 2) & Right("00" & Day(Now), 2) & Right("00" & Hour(Now), 2) & Right("00" & Minute(Now), 2) & Right("00" & Second(Now), 2)

backupFolder = objFSO.BuildPath(objFSO.GetParentFolderName(sourcePath), "_backup")
If objFSO.FolderExists(sourcePath) Then
    backupPath = objFSO.BuildPath(backupFolder, backupDate & "_" & objFSO.GetBaseName(sourcePath))
Else
    backupPath = objFSO.BuildPath(backupFolder, backupDate & "_" & objFSO.GetFileName(sourcePath))
End If

' バックアップフォルダが存在しない場合は作成
If Not objFSO.FolderExists(backupFolder) Then
    objFSO.CreateFolder backupFolder
End If

' フォルダをバックアップする場合
If objFSO.FolderExists(sourcePath) Then
    objFSO.CopyFolder sourcePath, backupPath, True
    WScript.Echo "Folder backed up to: " & backupPath
Else ' ファイルをバックアップする場合
    objFSO.CopyFile sourcePath, backupPath, True
    WScript.Echo "File backed up to: " & backupPath
End If

Set objFSO = Nothing
Set objShell = Nothing
