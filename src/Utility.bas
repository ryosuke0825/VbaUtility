Attribute VB_Name = "Utility"
Option Explicit

'Utiltyモジュール定数

'文字列のFalse。Stringを返す処理がキャンセルされた時に使用する。
Const STR_FALSE As String = "False"

'VarTypeにてデータ型がBoolean型
Const VAR_TYPE_BOOLEAN As String = "11"

'Dirコマンドの引数のフォルダ末尾に使用
Const cnsDIR As String = "\*.*"


'ダイアログからファイルを1つ選択する。
'選択したファイルのフルパスを戻り値とする。
'
'@parm(Optional) fileFilter ダイアログに表示するファイルの拡張子
'@parm(Optional) dialogTitle　ダイアログのタイトル
'@return openFileFullPaht 選択したファイルのフルパス。キャンセルを選択した時は「False」(String型)。
Function fetchFileFullPath(Optional fileFilter As String = "*", Optional dialogTitle As String = "ファイルを選択してください。") As String

    Dim currentFolderPath As String
    Dim thisWorkbookPath As String

    currentFolderPath = CurDir
    thisWorkbookPath = ThisWorkbook.Path & "\"
    ChDir thisWorkbookPath

    Dim argFileFilter As String
    argFileFilter = ",*." + fileFilter
    
    Dim openFileFullPaht As String
    openFileFullPaht = Application.GetOpenFilename(Title:=dialogTitle, fileFilter:=argFileFilter)
    
    ChDir currentFolderPath
    
    fetchFileFullPath = openFileFullPaht

End Function


'ダイアログからファイルを複数選択する。選択した全ファイルのフルパスを格納したコレクションをを戻り値とする。
'
'@parm(Optional) fileFilter ダイアログに表示するファイルの拡張子
'@parm(Optional) dialogTitle　ダイアログのタイトル
'@return openFileFullPaht 選択したファイルのフルパス。キャンセルを選択した時は「False」(String型)。
Function fetchFilesFullPath(Optional fileFilter As String = "*", Optional dialogTitle As String = "ファイルを選択してください。（複数選択可）") As Collection

    Dim i As Long
    Dim currentFolderPath As String
    Dim thisWorkbookPath As String

    'カレントディレクトリの設定
    currentFolderPath = CurDir
    thisWorkbookPath = ThisWorkbook.Path & "\"
    ChDir thisWorkbookPath

    'ファイルの選択
    Dim argFileFilter As String
    argFileFilter = ",*." + fileFilter

    Dim appGetOpenFileResult As Variant
    appGetOpenFilenameResult = Application.GetOpenFilename(Title:=dialogTitle, fileFilter:=argFileFilter, MultiSelect:=True)

    'カレントディレクトリの戻し
    ChDir currentFolderPath
    
    'String型配列への変換
    Dim openFileFullPaths As Collection
    Set openFileFullPaths = New Collection
    
    'キャンセルを選択している場合
    If (VarType(appGetOpenFilenameResult) = VAR_TYPE_BOOLEAN) Then
        openFileFullPaths.Add (STR_FALSE)
        Set fetchFilesFullPath = openFileFullPaths
    'ファイルを選択している場合
    Else
        For i = 1 To UBound(appGetOpenFilenameResult) Step 1
            openFileFullPaths.Add (appGetOpenFilenameResult(i))
        Next i
        Set fetchFilesFullPath = openFileFullPaths
    End If
    
End Function


'ダイアログからフォルダを選択する。
'選択したフォルダのフルパスを戻り値とする。
'
'@return 選択したファイルのフルパス。
Function fetchFolderFullPath() As String

    If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
        fetchFolderFullPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Else
        fetchFolderFullPath = STR_FALSE
    End If

End Function


'引数①「フルパス」からファイル名を取得して戻り値とする。
'
'@parm fullPath ファイルのフルパス
'@return ファイル名(\なし)
Function fetchFileName(ByVal fullPath As String) As String

    Dim delimiterPosition As Long
    delimiterPosition = InStrRev(fullPath, "\")

    If delimiterPosition <> "0" Then
        fetchFileName = Right(fullPath, Len(fullPath) - delimiterPosition)
    Else
        fetchFileName = STR_FALSE
    End If

End Function


'引数①「フルパス」から最後のフォルダ名を取得して戻り値とする。
'
'@parm fullPath ファイルのフルパス
'@return ファイル名(\なし)
Function fetchFolderName(ByVal fullPath As String) As String

    Dim delimiterPosition As Long
    delimiterPosition = InStrRev(fullPath, "\")

    If delimiterPosition <> 0 Then
        fetchFolderName = Right(fullPath, Len(fullPath) - delimiterPosition)
    Else
        fetchFolderName = STR_FALSE
    End If

End Function


'対象フォルダの全ファイルのフルパスを戻り値とする。
'引数が省略された場合は、ダイアログから選択する。
'
'@parm(Optional) folderPath ダイアログに表示するファイルの拡張子
'@return fileNames フォルダ内の全ファイル名が格納されたコレクション
Function fetchFileList(Optional folderPath As String) As Collection

    Dim fileNames As Collection
    
    '引数が省略された場合は、ダイアログからフォルダを選択する
    If Not IsMissing(folderPath) Then
        folderPath = fetchFolderFullPath()
        If folderPath = STR_FALSE Then
            Set fileNames = New Collection
            fileNames.Add (STR_FALSE)
            Set fetchFileList = fileNames
        End If
    End If


    Dim tmpFileName As String
    Set fileNames = New Collection

    tmpFileName = Dir(folderPath & cnsDIR, vbNormal)

    Do While tmpFileName <> ""
        fileNames.Add tmpFileName
        tmpFileName = Dir()
    Loop

    Set fetchFileList = fileNames

End Function


'引数①の拡張子付きファイル名の拡張子の手前に「_」と引数②を付与して戻り値とする
'
'@parm nowFilename 現在のファイル名
'@parm addString ファイル名に付け加える文字列
'@return newFilename 「_」と文字列が追加されたファイル名
Function AddStringFilename(ByVal nowFilename As String, ByVal addString As String) As String

    Dim filename As String
    Dim newFileNeme As String
    Dim fileExtension As String

    filename = Left(nowFilename, InStrRev(nowFilename, ".") - 1)
    fileExtension = Mid(nowFilename, InStrRev(nowFilename, ".") + 1, Len(nowFilename))
    newFileNeme = filename & "_" & addString & "." & fileExtension
    
    AddStringFilename = newFileNeme
    
End Function
