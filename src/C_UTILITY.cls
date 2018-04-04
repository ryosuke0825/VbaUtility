VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_UTILITY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'引数①が型の観点で数値であるか判定する。（文字列の数値の場合はFalseを返す）
'
'@param 型指定なし val
'@return Boolean True:数値 False:数値ではない
Function IsNumber(ByVal val) As Boolean

    If 2 <= VarType(val) And VarType(val) <= 5 Then
        IsNumber = True
        Exit Function
    End If
    
    IsNumber = False

End Function


'引数①のRangeが単一セルのみを選択してるか判定する。
'
'@param Range rng 判定するセル範囲
'@return Boolean True：単一セル False：複数セル
Function IsSingleCell(ByVal rng As Range) As Boolean

    If rng.count = 1 Then
        IsSingleCell = True
    End If
    
    IsSingleCell = False
    
End Function


'引数①のオブジェクトがRegExpオブジェクトであるか判定する。
'
'@param Object(RegExp) regexp
'@return Boolean True：RegExpオブジェクトである False：RegExpオブジェクトではない
Function IsRegExp(ByVal regexp As Object) As Boolean

    If TypeName(regexp) = "IRegExp2" Then
        IsRegExp = True
        Exit Function
    End If

    IsRegExp = False

End Function


'引数①のシートの引数②の列の最終入力行番号を返す
'
'@param Worksheet sh 対象のワークシート
'@param Long targetCol 対象行
'@return Range 引数①のセルから下方向へ入力ある範囲
Function CountLastRowNum(ByVal sh As Worksheet, ByVal targetCol As Long) As Long

    CountLastRowNum = sh.Cells(Rows.count, targetCol).End(xlUp).Row
        
End Function


'引数①のシートの引数②の行の最終入力列番号を返す
'
'@param sh 対象のワークシート
'@param targetRol 対象列
'@return range 引数①のセルから右方向へ入力ある範囲
Function CountLastCowNum(ByVal sh As Worksheet, ByVal targetRol As Long) As Long

    CountLastCowNum = sh.Cells(targetRol, Columns.count).End(xlToLeft).Column
        
End Function


'引数①のシートの引数②の列から引数③の値と一致するセルを1行目から空白のセルまで検索して、
'一致するセルが見つかった場合は､そのセルの行番号を返す｡見つからない場合は0を返す。
'
'@param Worksheet 入力範囲の取得元ワークシート
'@param Long 対象の列番号
'@param String 検索する文字列
'@param (Optional)Long 検索を開始する行番号（未入力時は1行目から）
'@return Long 文字列が入力されている行番号
Function GetFindWordEnteredCellRowNo(ByVal sh As Worksheet, ByVal colNum As Long, ByVal findWord As String, Optional ByVal firstRow As Long = 1) As Long

    Dim rowNum As Long
    rowNum = firstRow

    Do While sh.Cells(rowNum, colNum) <> ""
        If sh.Cells(rowNum, colNum) = findWord Then
            GetFindWordEnteredCellRowNo = rowNum
            Exit Function
        End If
        rowNum = rowNum + 1
    Loop
    
    GetFindWordEnteredCellRowNo = 0
    
End Function


'引数①のシートの引数②の行番号から引数③の値と一致するセルを1列目から右方向へ空白のセルまで探索する。
'一致するセルが見つかった場合は､そのセルの行番号を返す｡見つからない場合は0を返す。
'
'@param Worksheet 入力範囲の取得元ワークシート
'@param Long 対象の行番号
'@param String 検索する文字列
'@param (Optional)Long 検索を開始する列番号（未入力時は1行目から）
'@return Long 文字列が入力されている列番号
Function GetFindWordEnteredCellColNo(ByVal sh As Worksheet, ByVal rowNum As Long, ByVal findWord As String, Optional ByVal firstCol As Long = 1) As Long

    Dim colNum As Long
    colNum = firstCol

    Do While sh.Cells(rowNum, colNum) <> ""
        If sh.Cells(rowNum, colNum) = findWord Then
            GetFindWordEnteredCellColNo = colNum
            Exit Function
        End If
        colNum = colNum + 1
    Loop
    
    GetFindWordEnteredCellColNo = 0

End Function


'ダイアログからファイルを1つ選択する。
'選択したファイルのフルパスを戻り値とする。
'
'@param(Optional) fileFilter ダイアログに表示するファイルの拡張子
'@param(Optional) dialogTitle　ダイアログのタイトル
'@return openFileFullPaht 選択したファイルのフルパス。キャンセルを選択した時は「False」(String型)。
Function FetchFileFullPath(Optional fileFilter As String = "*", Optional dialogTitle As String = "ファイルを選択してください。") As String

    Dim currentFolderPath As String
    Dim thisWorkbookPath As String

    currentFolderPath = CurDir
    ChDrive Left(ThisWorkbook.path, 1)
    thisWorkbookPath = ThisWorkbook.path & "\"
    ChDir thisWorkbookPath

    Dim argFileFilter As String
    argFileFilter = ",*." + fileFilter
    
    Dim openFileFullPaht As String
    openFileFullPaht = Application.GetOpenFilename(title:=dialogTitle, fileFilter:=argFileFilter)
    
    ChDir currentFolderPath
    
    FetchFileFullPath = openFileFullPaht

End Function


'ファイル選択ダイアログからエクセルを選択して、そのエクセルを開いてworkbookオブジェクトを返す。
'@param string 選択ダイアログのタイトル
'@param string 選択ダイアログに表示する拡張子
'@param(Optional)  string 開いたエクセルが正しいかチェックする文字列
'@return workbook
Function SetSelectExcelOpenToWorkbook(Optional title As String = "対象ファイルを選択して下さい。", Optional extension As String = "xlsx", Optional targetFilename As String) As Workbook

    Dim wbFullPath As String
    wbFullPath = FetchFileFullPath(extension, title)
    
    If wbFullPath = "False" Then
        SetSelectExcelOpenToWorkbook = Nothing
    End If
    
    If Not IsMissing(targetFilename) Then
        If InStr(FetchFileName(wbFullPath), targetFilename) = 0 Then
            SetSelectExcelOpenToWorkbook = Nothing
        End If
    End If
    
    Workbooks.Open wbFullPath
    Set SetSelectExcelOpenToWorkbook = Workbooks(FetchFileName(wbFullPath))
    
End Function


'ファイル選択ダイアログからエクセルを選択して、そのエクセルを開いてworkbookオブジェクトを返す。
'@param string 選択ダイアログのタイトル
'@param string 選択ダイアログに表示する拡張子
'@param(Optional)  string 開いたエクセルが正しいかチェックする文字列
'@return workbook
Function SetOpenExcelToWorkbook(path As String) As Workbook
    
    Workbooks.Open path
    Set SetOpenExcelToWorkbook = Workbooks(FetchFileName(path))
    
End Function


'ダイアログからファイルを複数選択する。選択した全ファイルのフルパスを格納したコレクションを戻り値とする。
'
'@param(Optional) fileFilter ダイアログに表示するファイルの拡張子
'@param(Optional) dialogTitle　ダイアログのタイトル
'@return openFileFullPaht 選択したファイルのフルパス。キャンセルを選択した時は「False」(String型)。
Function FetchFilesFullPath(Optional fileFilter As String = "*", Optional dialogTitle As String = "ファイルを選択してください。（複数選択可）") As Collection

    Dim i As Long
    Dim currentFolderPath As String
    Dim thisWorkbookPath As String

    'カレントディレクトリの設定
    currentFolderPath = CurDir
    ChDrive Left(ThisWorkbook.path, 1)
    thisWorkbookPath = ThisWorkbook.path & "\"
    ChDir thisWorkbookPath

    'ファイルの選択
    Dim argFileFilter As String
    argFileFilter = ",*." + fileFilter

    Dim appGetOpenFilenameResult As Variant
    appGetOpenFilenameResult = Application.GetOpenFilename(title:=dialogTitle, fileFilter:=argFileFilter, MultiSelect:=True)

    'カレントディレクトリの戻し
    ChDir currentFolderPath
    
    'String型配列への変換
    Dim openFileFullPaths As Collection
    Set openFileFullPaths = New Collection
    
    'キャンセルを選択している場合
    If (VarType(appGetOpenFilenameResult) = "11") Then
        openFileFullPaths.Add ("False")
        Set FetchFilesFullPath = openFileFullPaths
    'ファイルを選択している場合
    Else
        For i = 1 To UBound(appGetOpenFilenameResult) Step 1
            openFileFullPaths.Add i, (appGetOpenFilenameResult(i))
        Next i
        Set FetchFilesFullPath = openFileFullPaths
    End If
    
End Function


'ダイアログからフォルダを選択する。
'選択したフォルダのフルパスを戻り値とする。
'
'@return 選択したファイルのフルパス。
Function FetchFolderFullPath() As String

    If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
        FetchFolderFullPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Else
        FetchFolderFullPath = "False"
    End If

End Function


'引数①「フルパス」からファイル名を取得して戻り値とする。
'
'@param fullPath ファイルのフルパス
'@return ファイル名(\なし) 「\」がない場合は、既にファイル名のみと判断してfullpathのまま返す
Function FetchFileName(ByVal fullpath As String) As String

    Dim delimiterPosition As Long
    delimiterPosition = InStrRev(fullpath, "\")

    If delimiterPosition <> 0 Then
        FetchFileName = Right(fullpath, Len(fullpath) - delimiterPosition)
    Else
        FetchFileName = fullpath
    End If

End Function


'引数①「フルパス」から拡張子なしのファイル名を取得して返す。
'
'@param fullPath ファイルのフルパス
'@return ファイル名(先頭に\なし、末尾に拡張子なし)
Function FetchFileNameNoExtension(ByVal fullpath As String) As String

    Dim delimiterPosition As Long
    delimiterPosition = InStrRev(fullpath, "\")
    
    If delimiterPosition <> 0 Then
        Dim fileName As String
        fileName = Right(fullpath, Len(fullpath) - delimiterPosition)
        FetchFileNameNoExtension = Left(fileName, InStrRev(fileName, ".") - 1)
    Else
        FetchFileNameNoExtension = "False"
    End If

End Function


'引数①「フルパス」から拡張子を取得して返す。（「.」なし）
'
'@param String fullpath 拡張子を取得するフルパス
'@return String フルパスから取得した拡張子
Function GetFileExtension(ByVal fullpath As String) As String

    GetFileExtension = Right(fullpath, Len(fullpath) - InStrRev(fullpath, "."))

End Function



'引数①「フルパス」から最後のフォルダ名を取得して戻り値とする。
'
'@param fullPath ファイルのフルパス
'@return ファイル名(\なし)
Function FetchFolderName(ByVal fullpath As String) As String

    Dim delimiterPosition As Long
    delimiterPosition = InStrRev(fullpath, "\")

    If delimiterPosition <> 0 Then
        FetchFolderName = Right(fullpath, Len(fullpath) - delimiterPosition)
    Else
        FetchFolderName = "False"
    End If

End Function


'引数①「フルパス」からパスを取得して戻り値とする。
'
'@param fullPath ファイルのフルパス
'@return パス(\あり)
Function FetchFilePath(ByVal fullpath As String) As String

    Dim delimiterPosition As Long
    delimiterPosition = InStrRev(fullpath, "\")

    If delimiterPosition <> "0" Then
        FetchFilePath = Left(fullpath, delimiterPosition)
    Else
        FetchFilePath = "False"
    End If

End Function


'対象フォルダの全ファイルのファイル名が格納されたcollectionを返す。
'引数が省略された場合は、ダイアログから選択する。
'
'@param(Optional) folderPath ダイアログに表示するファイルの拡張子
'@return filenames フォルダ内の全ファイル名が格納されたコレクション
Function GetFileFilenameWithinFolder(Optional folderPath As String) As Collection

    Dim filenames As Object

    '引数が省略された場合は、ダイアログからフォルダを選択する
    If Not IsMissing(folderPath) Then
        folderPath = FetchFolderFullPath()
        If folderPath = "False" Then
            Set filenames = New Collection
            filenames.Add 0, "False"
            Set GetFileFilenameWithinFolder = filenames
        End If
    End If

    Dim tmpFileName As String
    Set filenames = New Collection

    tmpFileName = Dir(folderPath & "\*.*", vbNormal)

    Dim i As Long
    i = 0
    Do While tmpFileName <> ""
        filenames.Add i, tmpFileName
        i = i + 1
        tmpFileName = Dir()
    Loop

    Set GetFileFilenameWithinFolder = filenames

End Function


'対象フォルダの全ファイルのフルパスを戻り値とする。
'引数が省略された場合は、ダイアログから選択する。
'
'@param(Optional) folderPath ダイアログに表示するファイルの拡張子
'@return filefullpaths フォルダ内の全ファイルのフルパスが格納されたコレクション
Function GetFileFullPathWithinFolder(Optional folderPath As String) As Collection

    Dim filefullpaths As Collection

    '引数が省略された場合は、ダイアログからフォルダを選択する
    If Not IsMissing(folderPath) Then
        folderPath = FetchFolderFullPath()
        If folderPath = "False" Then
            Set filefullpaths = New Collection
            filefullpaths.Add ("False")
            Set GetFileFullPathWithinFolder = filefullpaths
        End If
    End If

    Dim tmpFileName As String
    Set filefullpaths = New Collection

    tmpFileName = Dir(folderPath & "\*.*", vbNormal)

    Do While tmpFileName <> ""
        filefullpaths.Add folderPath & "\" & tmpFileName
        tmpFileName = Dir()
    Loop

    Set GetFileFullPathWithinFolder = filefullpaths

End Function


'引数①の拡張子付きファイル名の拡張子の手前に「_」と引数②を付与して戻り値とする
'
'@param nowFilename 現在のファイル名
'@param addString ファイル名に付け加える文字列
'@return newFilename 「_」と文字列が追加されたファイル名
Function AddStringFilename(ByVal nowFilename As String, ByVal addString As String) As String

    If nowFilename = "" Then
        AddStringFilename = ""
        Exit Function
    ElseIf addString = "" Then
        AddStringFilename = nowFilename
        Exit Function
    End If

    Dim fileName As String
    Dim newFileNeme As String
    Dim fileExtension As String

    fileName = Left(nowFilename, InStrRev(nowFilename, ".") - 1)
    fileExtension = Mid(nowFilename, InStrRev(nowFilename, ".") + 1, Len(nowFilename))
    newFileNeme = fileName & "_" & addString & "." & fileExtension

    AddStringFilename = newFileNeme
    
End Function


'引数①のシートのA1から入力のある最終セル（右下）までの範囲を取得する。
'
'@param Worksheet 入力範囲の取得元ワークシート
'@return range 引数①のワークシート内の入力ある範囲
Function FetchInputRangeFromSheetA1(ByVal sh As Worksheet)

    Dim lastCol As Long
    Dim lastRow As Long
    lastCol = sh.Range("A1").SpecialCells(xlLastCell).Column
    lastRow = sh.Range("A1").SpecialCells(xlLastCell).Row
    
    Set FetchInputRangeFromSheetA1 = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, lastCol))

End Function


'引数①のシートの引数②のセルから入力のある最終セル（右下）までの範囲を返す。
'
'@param Worksheet 入力範囲の取得元ワークシート
'@param String  入力範囲の取得元ワークシート
'@return range 引数①のワークシート内の入力ある範囲
Function FetchInputRangeFromSheet(ByVal sh As Worksheet, ByVal baseCell As Range)
    
    Dim lastCol As Long
    Dim lastRow As Long
    lastCol = baseCell.SpecialCells(xlLastCell).Column
    lastRow = baseCell.SpecialCells(xlLastCell).Row
    
    Set FetchInputRangeFromSheet = sh.Range(baseCell, sh.Cells(lastRow, lastCol))

End Function


'引数①のシートの引数②のセルから下方向へ最終入力行までセルの値をコレクションに格納して返す。
'
'@param Worksheet 入力範囲の取得元ワークシート
'@param String  値取得の開始位置となるセル（A1等）
'@param(Optional) Boolean 空文字をコレクションに加えるかのフラグ(デフォルトは加えない)
'@return collection 値を格納したコレクション
Function PutLastRowToValueCollection(ByVal sh As Worksheet, ByVal baseCell As Range, Optional blankFlag As Boolean = False) As Collection

    Dim col As Collection
    Set col = New Collection
    
    Dim lastRow As Long
    lastRow = sh.Cells(Rows.count, baseCell.Column).End(xlUp).Row
    
    Dim i As Long
    For i = baseCell.Row To lastRow Step 1
        If sh.Cells(i, baseCell.Column).Value <> "" Or blankFlag Then
            col.Add sh.Cells(i, baseCell.Column).Value
        End If
    Next i
    
    Set PutLastRowToValueCollection = col
  
End Function


'引数①のシートの引数②のセルから右方向へ最終入力行までセルの値をコレクションに格納して返す。
'
'@param Worksheet 入力範囲の取得元ワークシート
'@param String  値取得の開始位置となるセル（A1等）
'@param Boolean 空文字をコレクションに加えるかのフラグ(デフォルトは加えない)
'@return collection 値を格納したコレクション
Function PutLastColToValueCollection(ByVal sh As Worksheet, ByVal baseCell As Range, Optional blankFlag As Boolean = False) As Collection

    Dim col As Collection
    Set col = New Collection
    
    Dim lastCol As Long
    lastCol = sh.Cells(baseCell.Row, Columns.count).End(xlToLeft).Column
    
    Dim i As Long
    For i = baseCell.Column To lastCol Step 1
        If sh.Cells(i, baseCell.Row).Value <> "" Or blankFlag Then
            col.Add sh.Cells(i, baseCell.Row).Value
        End If
    Next i
    
    Set PutLastColToValueCollection = col
  
End Function


'引数①のシートの引数②のセルから下方向へ空白セルまでセルの値をコレクションに格納して返す。
'
'@param Worksheet セルを検索するワークシート
'@param string 検索を開始する基準セル
'@return collection 値を格納したコレクション
Function PutRowBlankCellToValueCollection(ByVal sh As Worksheet, ByVal baseCell As Range) As Collection

    Dim col As Collection
    Set col = New Collection

    Dim i As Long
    i = baseCell.Row
    Do While sh.Cells(i, baseCell.Column) <> ""
        col.Add sh.Cells(i, baseCell.Column).Value
        i = i + 1
    Loop
  
    Set PutRowBlankCellToValueCollection = col
  
End Function


'引数①のシートの引数②のセルから右方向へ空白セルまでセルの値をコレクションに格納して返す。
'
'@param Worksheet 入力範囲の取得元ワークシート
'@param String  値取得の開始位置となるセル（A1等）
'@return collection 値を格納したコレクション
Function PutColBlankCellToValueCollection(ByVal sh As Worksheet, ByVal baseCell As String) As Collection

    Dim col As Collection
    Set col = New Collection

    Dim i As Long
    i = Range(baseCell).Column
    Do While sh.Cells(Range(baseCell).Row, i) <> ""
        col.Add sh.Cells(Range(baseCell).Row, i).Value
        i = i + 1
    Loop
  
    Set PutColBlankCellToValueCollection = col
  
End Function


'引数①のコレクション内が全て数値であるかチェックする。
'
'@param collction チェックするコレクション
'@return boolean 判定結果
Function isCheckOnlyNumberInCollection(ByVal col As Collection) As Boolean

    '数値関係の型の「VarType」の出力値
    'vbInteger         2     整数型(Integer)
    'vbLong             3     長整数型(Long)
    'vbSingle           4     単精度浮動小数点数型(Single)
    'vbDouble         5     倍精度浮動小数点数型(Double)
    
    Dim val As Variant
    For Each val In col
        If Not (VarType(val) >= 2 And VarType(val) <= 5) Then
            isCheckOnlyNumberInCollection = False
            Exit Function
        End If
    Next
    
    isCheckOnlyNumberInCollection = True

End Function


'引数①のコレクション内の数値の最大値を返す。
'
'@param collction チェックするコレクション
'@return Long コレクション内の最大値
Function PopMaxValueInCollection(ByVal col As Collection)
    
    If col.count = 0 Then
        PopMaxValueInCollection = 0
        Exit Function
    End If
    
    Dim max As Long

    Dim val As Variant
    For Each val In col
        If max < val Then
            max = val
        End If
    Next
    
    PopMaxValueInCollection = max
    
End Function


'引数①のコレクション内の数値の最小値を返す。
'
'@param collction チェックするコレクション
'@return Long コレクション内の最小値
Function PopMinValueInCollection(ByVal col As Collection)
    
    If col.count = 0 Then
        PopMinValueInCollection = 0
        Exit Function
    End If

    Dim min As Long
    min = col.Item(1)

    Dim val As Variant
    For Each val In col
        If min > val Then
            min = val
        End If
    Next
    
    PopMinValueInCollection = min
    
End Function


'引数①のコレクション内の数値の合計値を返す。
'
'@param collction col 値が格納されているコレクション
'@return long sum 合計値
Function CalculateSumInCollection(ByVal col As Collection) As Long

  Dim sum As Long
  Dim Value As Variant
  For Each Value In col
    sum = sum + Value
  Next

  CalculateSumInCollection = sum
  
End Function


'引数①のコレクション内の数値の平均を返す。
'
'@param collction col チェックするコレクション
'@return Double ave コレクション内の値の平均値
Function CalculateAveInCollection(ByVal col As Collection) As Double

    If col.count = 0 Then
        CalculateAveInCollection = 0
        Exit Function
    End If

    Dim Value As Variant
    Dim sum As Double
    Dim ave As Double
    For Each Value In col
        sum = sum + Value
    Next
    ave = sum / col.count
    CalculateAveInCollection = ave

End Function


'引数①の文字列がコレクション内に存在するかチェックする。
'
'@param string ComparisonString 存在するか確認する文字列
'@param collection col 文字列が存在するかチェックするコレクション
'@return boolean true:存在する false:存在しない
Function IsCompareCellAndCollectionStr(ByVal comparisonString As String, ByVal col As Collection) As Boolean

  Dim val As Variant
  For Each val In col
    If comparisonString = val Then
      IsCompareCellAndCollectionStr = True
      Exit Function
    End If
  Next

  IsCompareCellAndCollectionStr = False
  
End Function


'引数①の文字列が引数②のコクション内の何れかの正規表現と一致するかチェックする。
'
'@param String comparisonString 正規表現とマッチするかチェックする文字列
'@param Collection regexpCol 正規表現が格納されたコレクション
'@param Object(RegExp) regexp 正規表現オブジェクト
'@return Boolean True:何れかの正規表現にマッチする False：全ての正規表現とマッチしない
Function IsCompareCellAndCollectionRegexp(ByVal comparisonString As String, ByVal col As Collection, regexp As Object) As Boolean

    Dim regexpPattern As Variant
    For Each regexpPattern In col
        regexp.pattern = regexpPattern
        If regexp.test(comparisonString) Then
            IsCompareCellAndCollectionRegexp = True
            Exit Function
        End If
    Next

    IsCompareCellAndCollectionRegexp = False
  
End Function


'引数①の文字列が、引数②の正規表現にマッチするかチェックする。
'
'@param String comparisonString チェックする文字列
'@param String pattern 正規表現パターン
'@param Object(RegExp) regexpObj 正規表現オブジェクト
'@return
Function IsCompareCellRegexp(ByVal comparisonString As String, ByVal pattern As String, ByVal regexp As Object) As Boolean

    regexp.pattern = pattern
    If regexp.test(comparisonString) Then
        IsCompareCellRegexp = True
        Exit Function
    End If

    IsCompareCellRegexp = False
    
End Function


'引数①の数値がコレクション内に存在するかチェックする。
'
'@param Long comparisonInt
'@param Collection col
'@return Boolean True:存在する  False:存在しない
Function IsCompareCellAndCollectionInt(ByVal comparisonInt As Long, ByVal col As Collection) As Boolean

    Dim val As Variant
    For Each val In col
        If Not IsNumber(val) Then
            GoTo NextFor
        End If

        If comparisonInt = val Then
            IsCompareCellAndCollectionInt = True
            Exit Function
        End If
NextFor:
    Next
    
    IsCompareCellAndCollectionInt = False

End Function


'引数①のシートの引数②のセルから下方向に入力のある最終セルまでの範囲を返す。
'
'@param Worksheet sh 対象シート
'@param Range baseCell 取得を開始するセル
'@return Range 下方向への入力あるのセル範囲
Function GetBaseCellFromUnderToInputHasRange(ByVal sh As Worksheet, ByVal baseCell As Range) As Range

    Dim Row As Long
    Row = baseCell.Row
    Do While sh.Cells(Row + 1, baseCell.Column) <> ""
       Row = Row + 1
    Loop

    Set GetBaseCellFromUnderToInputHasRange = sh.Range(sh.Cells(baseCell.Row, baseCell.Column), sh.Cells(Row, baseCell.Column))
    
End Function


'引数①のセルから右方向に入力のある最終セルまでの範囲を返す。
'
'@param Worksheet sh 対象シート
'@param Range baseCell 取得を開始するセル
'@return Range 右方向への入力あるのセル範囲
Function GetBaseCellFromRightToInputHasRange(ByVal sh As Worksheet, ByVal baseCell As Range) As Range

    Dim Column As Long
    Column = baseCell.Column
    Do While sh.Cells(baseCell.Row, Column) <> ""
       Column = Column + 1
    Loop

    Set GetBaseCellFromRightToInputHasRange = sh.Range(sh.Cells(baseCell.Row, baseCell.Column), sh.Cells(baseCell.Row, Column))

End Function


'引数①のブックの全シートのアクティブセルをA1にする。
'
'@param Workbook wb 対象ワークブック
Sub SetA1ActiveCellTotalSheet(ByVal book As Workbook)

    Dim sh As Variant
    For Each sh In book.Worksheets
        sh.Activate
        sh.Range("A1").Activate
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
    Next
    book.Worksheets(1).Activate

End Sub


'引数①の文字列を引数②の文字で区切ってコレクションに格納して返す。
'
'@param String targetStr 対象文字列
'@param String delimiter 区切り文字
'@return Collection 区切り文字で区切った文字を格納したコレクション
Function SetStringDelimitToCollection(ByVal targetStr As String, ByVal delimiter As String) As Collection

    Dim arr() As String
    arr() = Split(targetStr, delimiter)

    Dim col As Collection
    Set col = New Collection
    Dim val As Variant
    For Each val In arr()
        col.Add (val)
    Next

    Set SetStringDelimitToCollection = col
    
End Function


'引数①の文字を引数②の長さになるまで、先頭に引数③の文字を追加して返す。
'
'@param String targetStr
'@param Long digitNumber
'@param String fillCharacter
'@return String padStr 先頭を引数③の文字で埋めた後の文字列
Function PadStart(ByVal targetStr As String, ByVal digitNumber As Long, ByVal fillCharacter As String) As String

    If digitNumber <= Len(targetStr) Then
        PadStart = targetStr
        Exit Function
    End If

    Dim padStr As String
    padStr = targetStr

    Dim i As Long
    For i = digitNumber - Len(targetStr) To digitNumber Step 1
        padStr = fillCharacter & padStr
    Next i

    PadStart = padStr

End Function


'引数①の文字を引数②の長さになるまで、末尾に引数③の文字を追加して返す。
'
'@param String targetStr
'@param Long digitNumber
'@param String fillCharacter
'@return String padStr 末尾を引数③の文字で埋めた後の文字列
Function PadEnd(ByVal targetStr As String, ByVal digitNumber As Long, ByVal fillCharacter As String) As String

    If digitNumber <= Len(targetStr) Then
        PadEnd = targetStr
        Exit Function
    End If

    Dim padStr As String
    padStr = targetStr

    Dim i As Long
    For i = digitNumber - Len(targetStr) To digitNumber Step 1
        padStr = padStr & fillCharacter
    Next i

    PadEnd = padStr

End Function


'引数①の文字を引数②の回数繰り返した文字を返す。
'
'@param String targetStr 繰り返す対象文字列
'@param Long repeatTimes 繰り返す回数
'@return String 対象文字列を引数②の回数繰り返した文字列
Function RepeatString(ByVal targetStr As String, ByVal repeatTimes As Long) As String

    If repeatTimes < 1 Then
        RepeatString = targetStr
        Exit Function
    End If

    Dim afterString As String
    Dim i As Long
    For i = 1 To repeatTimes Step 1
        afterString = afterString & targetStr
    Next i

    RepeatString = afterString
    
End Function


'引数①のブック内から引数②の文字列を含むセルを検索して、最初にヒットしたセルを返す。
'
'@param Workbooks 検索対象のワークブック
'@param String 検索する文字列
'@return range 対象の値を含むセル
Function FindWordInBook(ByVal book As Workbook, ByVal word As String) As Range

    Dim sh As Worksheet
    For Each sh In book.Worksheets
        Dim hitRange As Range
        Set hitRange = sh.Cells.Find(word)
         
        If (hitRange Is Nothing) = False Then
            Exit For
        End If
    Next
     
    Set FindWordInBook = hitRange
    
End Function


'6桁又は8桁の日付を表す文字列にスラッシュを付与して返す。
'
'@param String dateString 日付を表す6桁か8桁の文字列
'@return String ／区切りの日付を表す文字列
Function GiveSlashToDate(ByVal dateString As String) As String

    If Not Len(dateString) = 6 Or Len(dateString) = 8 Then
        GiveSlashToDate = dateString
        Exit Function
    End If

    If Not IsNumeric(dateString) Then
        GiveSlashToDate = dateString
        Exit Function
    End If

    Dim slashDate As String
    If Len(dateString) = 6 Then
        slashDate = Left(dateString, 4) & "/" & Right(dateString, 2)
    ElseIf Len(dateString) = 8 Then
        slashDate = Left(dateString, 4) & "/" & Mid(dateString, 5, 2) & "/" & Right(dateString, 2)
    End If

    GiveSlashToDate = slashDate

End Function


'引数①のブックに引数②の名前のシートが存在するかチェックする。
'
'@param Workbook book
'@param String name
'@return Boolean True:存在する False：存在しない
Function IsSheet(ByVal book As Workbook, ByVal name As String) As Boolean

    name = "*" & name & "*"
    Dim val As Variant
    For Each val In book.Worksheets
        If val.name Like name Then
            IsSheet = True
            Exit Function
        End If
    Next

    IsSheet = False

End Function


'引数①のブックの初期シート名（Sheet1）を引数②の名前に変更する。（Sheet2,Sheet3があれば削除する）
'
'@param Workbook book 対象ブック
'@param String name 変更後のシート名
'@return Boolean True：変更成功 False：変更失敗（Sheet1が無い、同じ名前のシートがある）
Function ChangeInitialSheetName(ByVal book As Workbook, ByVal name As String) As Boolean

    Dim sh As Worksheet
    Dim sheet1 As Worksheet

    On Error Resume Next
    Set sh = book.Sheets(name)
    Set sheet1 = book.Worksheets("Sheet1")
    On Error GoTo 0

    Dim shExist As Boolean
    If Not sh Is Nothing Or sheet1 Is Nothing Then
        ChangeInitialSheetName = False
        Exit Function
    End If

    sheet1.name = name

    Dim Sheet2 As Worksheet
    Dim Sheet3 As Worksheet
    On Error Resume Next
    Set Sheet2 = book.Worksheets("Sheet2")
    Set Sheet3 = book.Worksheets("Sheet3")
    On Error GoTo 0

    Application.DisplayAlerts = False
    
    If Not Sheet2 Is Nothing Then
        Sheet2.Delete
    End If

    If Not Sheet3 Is Nothing Then
        Sheet3.Delete
    End If

    ChangeInitialSheetName = True
    
    Application.DisplayAlerts = True
    
End Function


'引数①の文字列の引数の②の位置から最後までを取得する。
'
'@param String target 文字列の取得元
'@param Long position 取得を開始する位置
'@return String 取得した文字列
Function MidLast(ByVal target As String, ByVal position As Long) As String

    If Len(target) < position Then
        MidLast = target
        Exit Function
    End If

    MidLast = Mid(target, position, Len(target))

End Function


'ステータスバーに表示する引数①の文字列を変更する。引数①が「False」の場合はステータスバーの非表示にする。
'
'@param String message ステータスバーに表示するメッセージ
Sub DisplayToStatsBar(ByVal message As String)

    If message = "False" Then
        Application.StatusBar = False
        Exit Sub
    End If

    Application.StatusBar = message

End Sub


'引数①の名前が付くエクセルが開かれているかチェックする。
'
'@param String name 開いているか確認するブック名
'@param Boolean perfectMatching 一致条件を完全一致(True)又は部分一致(False)のどちらにするかのフラグ
'@return Boolean True:開かれてる Flase：開かれていない
Function IsOpneExcel(ByVal name As String, Optional perfectMatching As Boolean = False) As Boolean

    If Not perfectMatching Then
        name = "*" & name & "*"
    End If

    Dim book As Variant
    For Each book In Workbooks
        If book.name Like name Then
            IsOpneExcel = True
            Exit Function
        End If
    Next

    IsOpneExcel = False

End Function


'引数①のブックを末尾に現在日時を付けて保存して、引数①のブックを閉じる。
'
'@param Workbook book 対象のワークブック
Sub CloseSaveBookInNowTime(ByVal book As Workbook)

    Dim nowTime As String
    nowTime = Format(Now, "yyyymmddHHMMSS")

    Dim newBookName As String
    newBookName = AddStringFilename(FetchFileName(book.name), Format(Now, "yyyymmddHHMMSS"))

    Application.DisplayAlerts = False
    book.SaveAs fileName:=ThisWorkbook.path & "\" & newBookName
    Application.DisplayAlerts = True
    
    book.Close

End Sub


'引数①のブックを引数②の名前で同じフォルダに保存する。
'
'@param Workbook book
'@param String name
'@return Boolean True:保存成功 False:同名ファイルがあり失敗
Function CloseSaveBookInAlias(ByVal book As Workbook, ByVal name As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(book.path & "\" & name) Or _
        fso.FileExists(book.path & "\" & name & "." & GetFileExtension(book.name)) Then
        CloseSaveBookInAlias = False
        Exit Function
    End If

    Application.DisplayAlerts = False
    If InStr(name, ".") <> 0 Then
        book.SaveAs (book.path & "\" & name)
    Else
        book.SaveAs (book.path & "\" & name & "." & GetFileExtension(book.name))
    End If
    Application.DisplayAlerts = True
    
    book.Close
    
    CloseSaveBookInAlias = True

End Function


'引数①のブックを末尾に現在日時を付けて保存する。
'
'@param Workbook book 名前をつけて保存するブック
Sub SaveBookInNowTime(ByVal book As Workbook)

    Dim nowTime As String
    nowTime = Format(Now, "yyyymmddHHMMSS")

    Dim newBookName As String
    newBookName = AddStringFilename(FetchFileName(book.name), Format(Now, "yyyymmddHHMMSS"))

    Application.DisplayAlerts = False
    book.SaveCopyAs fileName:=ThisWorkbook.path & "\" & newBookName
    Application.DisplayAlerts = True

End Sub


'引数①のブックに、引数②のシートのシート名の末尾に現在日時を付与してコピーする。
'
'@param Workbook book 保存先ブック
'@param Worksheet sh 保存元シート
Sub SaveSheetInNowTime(ByVal book As Workbook, ByVal sh As Worksheet)

    Dim nowTime As String
    nowTime = Format(Now, "yyyymmddHHMMSS")

    sh.Copy After:=book.Sheets(sh.name)

    Dim i As Long
    For i = 1 To book.Worksheets.count Step 1
        If book.Worksheets(i).name = sh.name Then
            book.Worksheets(i + 1).name = sh.name & "_" & nowTime
            Exit Sub
        End If
    Next i

End Sub


'引数①のブックに引数③番目に引数②の名前のシートを追加する。
'
'@param Workbook book 対象ワークブック
'@param String name 追加するシート名
'@param Long place 追加する位置（前から何番目のシートか）
'@return Boolean True：作成成功、False：作成失敗（同名シートが存在する）
Function AddSheet(ByVal book As Workbook, ByVal name As String, ByVal place As Long) As Boolean

    If IsSheet(book, name) Then
        AddSheet = False
        Exit Function
    End If

    If place < book.Worksheets.count() Then
        book.Worksheets.Add After:=Worksheets(place - 1)
        book.Worksheets(place).name = name
    Else
        book.Worksheets().Add After:=book.Worksheets(book.Worksheets.count)
        book.Worksheets(book.Worksheets.count).name = name
    End If
    
    AddSheet = True

End Function


'引数①のブックの全シート名をコレクションに格納して返す。
'
'@param Workbook book シート名を取得するブック
'@return Collection col 全シート名を格納したコレクション
Function GetAllSheetName(ByVal book As Workbook) As Collection

    Dim col As Collection
    Set col = New Collection

    Dim val As Variant
    For Each val In book.Worksheets
        col.Add val.name
    Next

    Set GetAllSheetName = col

End Function


'引数①のアルファベットをエクセルの列番号に変換した数値を返す。
'
'@param String alphabet
'@return Long number
Function ConvertAlphabetIntoNumber(ByVal alphabet As String) As Long

    On Error Resume Next
    Dim rng As Range
    Set rng = Range(alphabet & ":" & alphabet)

    If Err.Number = 1004 Then
        ConvertAlphabetIntoNumber = 0
        End
    End If
    
    ConvertAlphabetIntoNumber = rng.Column

End Function


'引数①の数値をエクセルの列を表すアルファベットに変換した文字列を返す。
'
'@param Long num アルファベットに変換する数字
'@return String 変換されたアルファベット
Function ConvertNumberIntoAlphabet(ByVal num As Long) As String

    If num <= 0 Then
        ConvertNumberIntoAlphabet = ""
        Exit Function
    End If

    Dim alphabet As String
    alphabet = ""

    alphabet = Cells(1, num).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    alphabet = Left(alphabet, Len(alphabet) - 1)

    ConvertNumberIntoAlphabet = alphabet

End Function


'引数①のシートの引数②の列番号から引数③の列番号までの列幅を自動調整する。引数に問題があるとFalseを返す。
'
'@param worksheet sh 対象シート
'@param Long startCol 列幅調整をする最初の列番号
'@param Long endCol 列幅調整をする最後の列番号
'@return Boolean True:処理成功 False:処理失敗（引数エラー）
Function AdjustColumnWidthAuto(ByVal sh As Worksheet, ByVal startCol As Long, ByVal endCol As Long) As Boolean

    If startCol < 0 Or endCol < startCol Then
        AdjustColumnWidthAuto = False
        Exit Function
    End If

    sh.Range(sh.Columns(startCol), sh.Columns(endCol)).EntireColumn.AutoFit

    AdjustColumnWidthAuto = True

End Function


'引数①のシートの引数②の行番号から引数③の行番号までの行幅を自動調整する。引数に問題があるとFalseを返す。
'
'@param worksheet sh 対象シート
'@param Long startRow 行幅調整をする最初の行番号
'@param Long endRow 行幅調整をする最後の行番号
'@return Boolean True:処理成功 False:処理失敗（引数エラー）
Function AdjustRowWidthAuto(ByVal sh As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Boolean

    If startRow < 0 Or endRow < startRow Then
        AdjustRowWidthAuto = False
        Exit Function
    End If

    sh.Range(sh.Rows(startRow), sh.Rows(endRow)).EntireRow.AutoFit

    AdjustRowWidthAuto = True

End Function


'色選択ウィンドウを表示して、選択した色コードを返す。
'
'@return Long 選択位した色のカラーコード(選択していない場合は、現在のから)
Function GetSelectedColor() As Long

    'カラーパレットで選択した色は、ブック全体で共通の設定なので、
    '処理前の色のバックアップを取得する。
    Dim thisbookColorPaletteNo1Backup As Long
    thisbookColorPaletteNo1Backup = ThisWorkbook.Colors(1)

    If ThisWorkbook.Application.Dialogs(xlDialogEditColor).Show(1) Then
        GetSelectedColor = ThisWorkbook.Colors(1)
        ThisWorkbook.Colors(1) = thisbookColorPaletteNo1Backup
        Exit Function
    End If

    ThisWorkbook.Colors(1) = thisbookColorPaletteNo1Backup
    GetSelectedColor = ThisWorkbook.Colors(1)

End Function



'引数①のベースセルから右方向へ空白セルまで、引数②の文字列と一致するセルを探索し、最初に一致した列番号を返す。
'
'@param Range checkCell 探索を開始するセル
'@param String findWord 探索する文字列
'@return Long 最初に一致したセルの列番号
Function GetSearchStringMatchCellColumnNumToRight(ByVal checkCell As Range, ByVal findWord As String) As Long

    Do While checkCell <> ""
        If checkCell = findWord Then
            GetSearchStringMatchCellColumnNumToRight = checkCell.Column
            Exit Function
        End If
        Set checkCell = checkCell.Offset(0, 1)
    Loop

    GetSearchStringMatchCellColumnNumToRight = 0

End Function


'引数①のベースセルから右方向へ空白セルまで、引数②の正規表現と一致するセルを探索し、最初に一致した列番号を返す。
'
'@param Range checkCell 探索を開始するセル
'@param String regexpWord 探索する正規表現
'@param Object(RegExp) regexp 正規表現オブジェクト
'@return Long 最初に一致したセルの列番号(見つからない場合は、0を返す。)
Function GetSearchRegexpMatchCellColumnNumToRight(ByVal checkCell As Range, ByVal regexpWord As String, ByVal regexp As Object) As Long

    If Not IsRegExp(regexp) Then
        GetSearchRegexpMatchCellColumnNumToRight = 0
        Exit Function
    End If

    regexp.pattern = regexpWord
    Do While checkCell <> ""
        If regexp.test(checkCell.Value) Then
            GetSearchRegexpMatchCellColumnNumToRight = checkCell.Column
            Exit Function
        End If
        Set checkCell = checkCell.Offset(0, 1)
    Loop

    GetSearchRegexpMatchCellColumnNumToRight = 0

End Function


'引数①のベースセルから右方向へ空白セルまで、引数②のコレクション内の文字列の何れかと一致するセルを探索し、最初に一致した列番号を返す。
'
'@param Range checkCell 探索の開始セル
'@param Collection findWords 検索ワードの格納されたコレクション
'@return Long 一致したセルの列番号
Function GetSearchStringInCollectionMatchCellColumnNumToRight(ByVal checkCell As Range, ByVal findWords As Collection) As Long

    Do While checkCell <> ""
        If IsCompareCellAndCollectionStr(checkCell.Value, findWords) Then
            GetSearchStringInCollectionMatchCellColumnNumToRight = checkCell.Column
            Exit Function
        End If
        Set checkCell = checkCell.Offset(0, 1)
    Loop

    GetSearchStringInCollectionMatchCellColumnNumToRight = 0

End Function


'引数①のベースセルから右方向へ空白セルまで、引数②のコレクション内の正規表現の何れかと一致するセルを探索し、最初に一致した列番号を返す。
'
'@param Range checkCell 探索の開始セル
'@param Collection findWords 正規表現が格納されたコレクション
'@param
'@return Long 一致したセルの列番号
Function GetSearchRegexpInCollectionMatchCellColumnNumToRight(ByVal checkCell As Range, _
                                                              ByVal findWords As Collection, _
                                                              ByVal regexp As Object) As Long

    If Not IsRegExp(regexp) Then
        GetSearchRegexpInCollectionMatchCellColumnNumToRight = 0
        Exit Function
    End If

    Do While checkCell <> ""
        If IsCompareCellAndCollectionRegexp(checkCell.Value, findWords, regexp) Then
            GetSearchRegexpInCollectionMatchCellColumnNumToRight = checkCell.Column
            Exit Function
        End If
        Set checkCell = checkCell.Offset(0, 1)
    Loop

    GetSearchRegexpInCollectionMatchCellColumnNumToRight = 0

End Function


'引数①のベースセルから下方向へ空白セルまで、引数②の文字列と一致するセルを探索し、最初に一致した行番号を返す。
'
'@param Range baseCell 探索を開始するセル
'@param String findWord 探索する文字列
'@return Long 最初に一致したセルの列番号
Function GetSearchStringMatchCellRowNumToLower(ByVal baseCell As Range, ByVal findWord As String) As Long

    Do While baseCell <> ""
        If baseCell = findWord Then
            GetSearchStringMatchCellRowNumToLower = baseCell.Row
            Exit Function
        End If
        Set baseCell = baseCell.Offset(1, 0)
    Loop

    GetSearchStringMatchCellRowNumToLower = 0

End Function


'引数①のベースセルから下方向へ空白セルまで、引数②の正規表現と一致するセルを探索し、最初に一致した行番号を返す。
'
'@param Range checkCell 探索を開始するセル
'@param String regexpWord 探索する正規表現
'@param Object(RegExp) regexp 正規表現オブジェクト
'@return Long 最初に一致したセルの行番号(見つからない場合は、0を返す。)
Function GetSearchRegexpMatchCellRowNumToLower(ByVal checkCell As Range, ByVal regexpWord As String, ByVal regexp As Object) As Long

    If Not IsRegExp(regexp) Then
        GetSearchRegexpMatchCellRowNumToLower = 0
        Exit Function
    End If

    regexp.pattern = regexpWord
    Do While checkCell <> ""
        If regexp.test(checkCell.Value) Then
            GetSearchRegexpMatchCellRowNumToLower = checkCell.Row
            Exit Function
        End If
        Set checkCell = checkCell.Offset(1, 0)
    Loop

    GetSearchRegexpMatchCellRowNumToLower = 0

End Function


'引数①のベースセルから下方向へ空白セルまで、引数②のコレクション内の文字列の何れかと一致するセルを探索し、最初に一致した行番号を返す。
'
'@param Range checkCell 探索の開始セル
'@param Collection findWords 検索ワードの格納されたコレクション
'@return Long 一致したセルの行番号
Function GetSearchStringInCollectionMatchCellRowNumToLower(ByVal checkCell As Range, ByVal findWords As Collection) As Long

    Do While checkCell <> ""
        If IsCompareCellAndCollectionStr(checkCell.Value, findWords) Then
            GetSearchStringInCollectionMatchCellRowNumToLower = checkCell.Row
            Exit Function
        End If
        Set checkCell = checkCell.Offset(1, 0)
    Loop

    GetSearchStringInCollectionMatchCellRowNumToLower = 0

End Function


'引数①のベースセルから下方向へ空白セルまで、引数②のコレクション内の正規表現の何れかと一致するセルを探索し、最初に一致した行番号を返す。
'
'@param Range checkCell 探索の開始セル
'@param Collection findWords 正規表現が格納されたコレクション
'@param
'@return Long 一致したセルの行番号
Function GetSearchRegexpInCollectionMatchCellRowNumToLower(ByVal checkCell As Range, _
                                                              ByVal findWords As Collection, _
                                                              ByVal regexp As Object) As Long

    If Not IsRegExp(regexp) Then
        GetSearchRegexpInCollectionMatchCellRowNumToLower = 0
        Exit Function
    End If

    Do While checkCell <> ""
        If IsCompareCellAndCollectionRegexp(checkCell.Value, findWords, regexp) Then
            GetSearchRegexpInCollectionMatchCellRowNumToLower = checkCell.Row
            Exit Function
        End If
        Set checkCell = checkCell.Offset(1, 0)
    Loop

    GetSearchRegexpInCollectionMatchCellRowNumToLower = 0

End Function


'引数①から引数②までのランダムな整数を返す。
'
'@param Long lowestValue
'@param Long highestValue
'@return Long 生成した乱数を返す。LowestValue>HighestValueの時は0を返す。
Function CreateRandomNumber(ByVal lowestValue As Long, ByVal highestValue As Long) As Long

    If lowestValue > highestValue Then
        CreateRandomNumber = 0
        Exit Function
    End If

    CreateRandomNumber = Int((highestValue - lowestValue + 1) * Rnd + lowestValue)

End Function


'引数①のShapeオブジェクトのある１つ下のセルの行番号を返す。
'
'@param Shape obj 対象のShapeオブジェクト
'@return Long 1つ下のセルの行番号
Function GetObjectOneUnderRowNo(ByVal obj As Shape) As Long

    GetObjectOneUnderRowNo = obj.BottomRightCell.Row + 1

End Function


'引数①のShapeオブジェクトのある１つ右のセルの列番号を返す。
'
'@param Shape obj 対象のShapeオブジェクト
'@return Long 1つ右のセルの列番号
Function GetObjectOneUnderColumnNo(ByVal obj As Shape) As Long

    GetObjectOneUnderColumnNo = obj.BottomRightCell.Column + 1

End Function


'引数①の文字列の全ての空白(半角、全角、タブ)を削除した文字列を返す。
'
'@param String str 空白を削除する対象の文字列
'@return String 空白を削除した文字列
Function RemoveAllSpace(ByVal str As String) As String

    str = Replace(str, " ", "")
    str = Replace(str, "　", "")
    str = Replace(str, "   ", "")
    
    RemoveAllSpace = str

End Function


'引数①の文字列を日付に変換して返す。
'
'@param String strDate 日付に変換する文字列
'@return Date 変換した日付。日付に変換できない場合は、「西暦9999年12月31日」を返す。
Function ConvertStringIntoDate(ByVal strDate As String) As Date

    If IsDate(strDate) = True Then
        ConvertStringIntoDate = CDate(strDate)
        Exit Function
    End If

    ConvertStringIntoDate = #12/31/9999#
    
End Function


'引数①のフォルダが存在するかチェックして、存在しない場合は作成する。
'
'@param String folderPath
Sub CreateFolder(ByVal folderPath As String)

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

End Sub


'引数①のファイル名（フルパス）を引数②に変更する。
'
'@param String filePath 変更対象ファイルのフルパス
'@param String newName 変更後のファイル名
'@return Boolean True:変更完了 False：同名ファイルが存在する、対象ファイルが存在しない
Function ChangeFilename(ByVal filePath As String, ByVal newName As String) As Boolean

    Dim newFileFullpath As String
    newFileFullpath = FetchFilePath(filePath) & newName
    
    If Dir(newFileFullpath) <> "" Or Dir(filePath) = "" Then
        ChangeFilename = False
        Exit Function
    End If

    Name filePath As newFileFullpath
    ChangeFilename = True
    
End Function


'引数①のファイル（フルパス）の拡張子を引数②に変更する。
'
'@param String filePath 変更対象ファイルのフルパス
'@param String newExtension 変更後の拡張子
'@return Boolean True:変更成功 False：同名ファイルが存在する、対象ファイルが存在しない
Function ChangeFileExtension(ByVal filePath As String, ByVal newExtension As String) As Boolean

    Dim newFileFullpath As String
    newFileFullpath = FetchFilePath(filePath) & FetchFileNameNoExtension(filePath) & "." & newExtension

    If Dir(newFileFullpath) <> "" Or Dir(filePath) = "" Then
        ChangeFileExtension = False
        Exit Function
    End If

    Name filePath As newFileFullpath
    ChangeFileExtension = True

End Function


'A1から引数①の行番号、引数②の列番号を印刷エリアを設定する。
'
'@param Worksheet sh 印刷範囲を設定するワークシート
'@param Long lastRow 印刷範囲の最終行番号
'@param Long lastCol 印刷範囲の最終列番号
'@return Boolena True: False:引数に0以下の値がある
Function SetPrintingAreas(ByVal sh As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long) As Boolean

    If lastRow < 1 Or lastCol < 1 Then
        SetPrintingAreas = False
        Exit Function
    End If

    sh.PageSetup.PrintArea = Range(Cells(1, 1), Cells(lastRow, lastCol)).Address
    SetPrintingAreas = True

End Function


'引数①のRange範囲から引数②の文字列と一致する全てのセルを格納したコレクションを返す。
'
'@param Range rng チェック範囲
'@param String findWord チェックする文字列
'@return Collection 対象の文字列と一致する全セルを格納したコレクション
Function GetStringMatchAllCell(ByVal rng As Range, ByVal findWord As String) As Collection

    Dim col As Collection
    Set col = New Collection

    Dim val As Variant
    For Each val In rng
        If val = findWord Then
            col.Add val
        End If
    Next

    Set GetStringMatchAllCell = col

End Function


'引数①のRange範囲から引数②の正規表現と一致する全てのセルを格納したコレクションを返す。
'
'@param Range rng 対象セル範囲
'@param String regexpWord 正規表現のワード
'@param Object(RegExp) regexp 正規表現オブジェクト
'@return Collection 正規表現と一致するセルを格納したコレクション
Function GetRegexpMatchAllCell(ByVal rng As Range, ByVal regexpWord As String, ByVal regexp As Object) As Collection

    regexp.pattern = regexpWord

    Dim col As Collection
    Set col = New Collection

    Dim val As Variant
    For Each val In rng
        If regexp.test(val.Value) Then
            col.Add val
        End If
    Next

    Set GetRegexpMatchAllCell = col
    
End Function


'引数①のコレクションの内容を引数②のセルから下方向へ書き出す。(For eachでインデックス順に取り出す)
'
'@param Collection col 書き出す値を格納したコレクション
'@param Range baseCell 書き出し始めるセル
Sub WriteCollectionToLower(ByVal col As Collection, ByVal baseCell As Range)

    Dim val As Variant
    For Each val In col
        baseCell.Value = val
        Set baseCell = baseCell.Offset(1, 0)
    Next

End Sub


'引数①のコレクションの内容を引数②のセルから右方向へ書き出す。(For eachでインデックス順に取り出す)
'
'@param Collection col 書き出す値を格納したコレクション
'@param Range baseCell 書き出し始めるセル
Sub WriteCollectionToRight(ByVal col As Collection, ByVal baseCell As Range)

    Dim val As Variant
    For Each val In col
        baseCell.Value = val
        Set baseCell = baseCell.Offset(0, 1)
    Next

End Sub


'引数①のブックの全シートのアクティブセルをA1にして、画面を一番右上へスクロールする。
'
'@param Workbook book 対象ワークブック
Sub SetAllSheetActiveCellToA1(ByVal book As Workbook)

    Dim sh As Variant
    For Each sh In book.Worksheets
        sh.Activate
        sh.Range("A1").Select
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
    Next
    
    book.Worksheets(1).Activate

End Sub


'引数①のRange範囲の文字と罫線と背景色を削除する。
'
'@param Range rng 対象のセル範囲
Sub ClearFormatAndContent(ByVal rng As Range)

    rng.ClearFormats
    rng.ClearContents

End Sub


'引数①の文字列の左側から引数②の数だけ文字を削除した文字列を返す。
'
'@param String str 対象文字列
'@param Long num 削除する文字数
'@return String 指定文字数削除した後の文字列
Function DeleteLeft(ByVal str As String, ByVal num As Long) As String

    If Len(str) <= num Then
        DeleteLeft = ""
    End If

    DeleteLeft = Mid(str, num + 1, Len(str))
    
End Function
