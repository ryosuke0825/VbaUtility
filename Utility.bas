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
'@return OpenFileName 選択したファイルのフルパス。キャンセルを選択した時は「False」(String型)。
Function fetchFileFullPath(Optional fileFilter As String = "*", Optional dialogTitle As String = "ファイルを選択してください。") As String

    Dim currentFolderPath As String
    Dim thisWorkbookPath As String

    currentFolderPath = CurDir
    thisWorkbookPath = ThisWorkbook.Path & "\"
    ChDir thisWorkbookPath

    Dim argFileFilter As String
    argFileFilter = ",*." + fileFilter
    
    Dim openFileName As String
    openFileName = Application.GetOpenFilename(Title:=dialogTitle, fileFilter:=argFileFilter)
    
    ChDir currentFolderPath
    
    fetchFileFullPath = openFileName

End Function

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

    Dim appGetOpenFilenameResult As Variant
    appGetOpenFilenameResult = Application.GetOpenFilename(Title:=dialogTitle, fileFilter:=argFileFilter, MultiSelect:=True)

    'カレントディレクトリの戻し
    ChDir currentFolderPath
    
    'String型配列への変換
    Dim openFileNames As Collection
    Set openFileNames = New Collection
    
    'キャンセルを選択している場合
    If (VarType(appGetOpenFilenameResult) = VAR_TYPE_BOOLEAN) Then
        openFileNames.Add (STR_FALSE)
        Set fetchFilesFullPath = openFileNames
    'ファイルを選択している場合
    Else
        For i = 1 To UBound(appGetOpenFilenameResult) Step 1
            openFileNames.Add (appGetOpenFilenameResult(i))
        Next i
        Set fetchFilesFullPath = openFileNames
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


'引数のブックを末尾に現在日時を付与して保存して閉じる
'
'@param wb ワークブック（名前を付けて保存する対象）
Sub closeAfterSaveAsBookNowTime(ByVal wb As Workbook)

    '現在日時を生成
    Dim nowTime As String
    nowTime = Format(Now, "yyyymmddHHMMSS")
    
    '新しいファイル名の生成
    Dim newFileNeme As String
    Dim oldFileName As String
    Dim fileExtension As String
    
    oldFileName = Left(wb.Name, InStrRev(wb.Name, ".") - 1)
    fileExtension = Mid(wb.Name, InStrRev(wb.Name, ".") + 1, Len(wb.Name))
    newFileNeme = oldFileName & "_" & nowTime & "." & fileExtension
    
    '名前を付けて保存した後に閉じる
    wb.SaveAs wb.Path & "\" & newFileNeme
    wb.Close

End Sub

