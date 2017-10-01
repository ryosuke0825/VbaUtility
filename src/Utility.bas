Attribute VB_Name = "Utility"
Option Explicit

'Utilty���W���[���萔

'�������False�BString��Ԃ��������L�����Z�����ꂽ���Ɏg�p����B
Const STR_FALSE As String = "False"

'VarType�ɂăf�[�^�^��Boolean�^
Const VAR_TYPE_BOOLEAN As String = "11"

'Dir�R�}���h�̈����̃t�H���_�����Ɏg�p
Const cnsDIR As String = "\*.*"


'�_�C�A���O����t�@�C����1�I������B
'�I�������t�@�C���̃t���p�X��߂�l�Ƃ���B
'
'@parm(Optional) fileFilter �_�C�A���O�ɕ\������t�@�C���̊g���q
'@parm(Optional) dialogTitle�@�_�C�A���O�̃^�C�g��
'@return openFileFullPaht �I�������t�@�C���̃t���p�X�B�L�����Z����I���������́uFalse�v(String�^)�B
Function fetchFileFullPath(Optional fileFilter As String = "*", Optional dialogTitle As String = "�t�@�C����I�����Ă��������B") As String

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


'�_�C�A���O����t�@�C���𕡐��I������B�I�������S�t�@�C���̃t���p�X���i�[�����R���N�V��������߂�l�Ƃ���B
'
'@parm(Optional) fileFilter �_�C�A���O�ɕ\������t�@�C���̊g���q
'@parm(Optional) dialogTitle�@�_�C�A���O�̃^�C�g��
'@return openFileFullPaht �I�������t�@�C���̃t���p�X�B�L�����Z����I���������́uFalse�v(String�^)�B
Function fetchFilesFullPath(Optional fileFilter As String = "*", Optional dialogTitle As String = "�t�@�C����I�����Ă��������B�i�����I���j") As Collection

    Dim i As Long
    Dim currentFolderPath As String
    Dim thisWorkbookPath As String

    '�J�����g�f�B���N�g���̐ݒ�
    currentFolderPath = CurDir
    thisWorkbookPath = ThisWorkbook.Path & "\"
    ChDir thisWorkbookPath

    '�t�@�C���̑I��
    Dim argFileFilter As String
    argFileFilter = ",*." + fileFilter

    Dim appGetOpenFileResult As Variant
    appGetOpenFilenameResult = Application.GetOpenFilename(Title:=dialogTitle, fileFilter:=argFileFilter, MultiSelect:=True)

    '�J�����g�f�B���N�g���̖߂�
    ChDir currentFolderPath
    
    'String�^�z��ւ̕ϊ�
    Dim openFileFullPaths As Collection
    Set openFileFullPaths = New Collection
    
    '�L�����Z����I�����Ă���ꍇ
    If (VarType(appGetOpenFilenameResult) = VAR_TYPE_BOOLEAN) Then
        openFileFullPaths.Add (STR_FALSE)
        Set fetchFilesFullPath = openFileFullPaths
    '�t�@�C����I�����Ă���ꍇ
    Else
        For i = 1 To UBound(appGetOpenFilenameResult) Step 1
            openFileFullPaths.Add (appGetOpenFilenameResult(i))
        Next i
        Set fetchFilesFullPath = openFileFullPaths
    End If
    
End Function


'�_�C�A���O����t�H���_��I������B
'�I�������t�H���_�̃t���p�X��߂�l�Ƃ���B
'
'@return �I�������t�@�C���̃t���p�X�B
Function fetchFolderFullPath() As String

    If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
        fetchFolderFullPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Else
        fetchFolderFullPath = STR_FALSE
    End If

End Function


'�����@�u�t���p�X�v����t�@�C�������擾���Ė߂�l�Ƃ���B
'
'@parm fullPath �t�@�C���̃t���p�X
'@return �t�@�C����(\�Ȃ�)
Function fetchFileName(ByVal fullPath As String) As String

    Dim delimiterPosition As Long
    delimiterPosition = InStrRev(fullPath, "\")

    If delimiterPosition <> "0" Then
        fetchFileName = Right(fullPath, Len(fullPath) - delimiterPosition)
    Else
        fetchFileName = STR_FALSE
    End If

End Function


'�����@�u�t���p�X�v����Ō�̃t�H���_�����擾���Ė߂�l�Ƃ���B
'
'@parm fullPath �t�@�C���̃t���p�X
'@return �t�@�C����(\�Ȃ�)
Function fetchFolderName(ByVal fullPath As String) As String

    Dim delimiterPosition As Long
    delimiterPosition = InStrRev(fullPath, "\")

    If delimiterPosition <> 0 Then
        fetchFolderName = Right(fullPath, Len(fullPath) - delimiterPosition)
    Else
        fetchFolderName = STR_FALSE
    End If

End Function


'�Ώۃt�H���_�̑S�t�@�C���̃t���p�X��߂�l�Ƃ���B
'�������ȗ����ꂽ�ꍇ�́A�_�C�A���O����I������B
'
'@parm(Optional) folderPath �_�C�A���O�ɕ\������t�@�C���̊g���q
'@return fileNames �t�H���_���̑S�t�@�C�������i�[���ꂽ�R���N�V����
Function fetchFileList(Optional folderPath As String) As Collection

    Dim fileNames As Collection
    
    '�������ȗ����ꂽ�ꍇ�́A�_�C�A���O����t�H���_��I������
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


'�����@�̊g���q�t���t�@�C�����̊g���q�̎�O�Ɂu_�v�ƈ����A��t�^���Ė߂�l�Ƃ���
'
'@parm nowFilename ���݂̃t�@�C����
'@parm addString �t�@�C�����ɕt�������镶����
'@return newFilename �u_�v�ƕ����񂪒ǉ����ꂽ�t�@�C����
Function AddStringFilename(ByVal nowFilename As String, ByVal addString As String) As String

    Dim filename As String
    Dim newFileNeme As String
    Dim fileExtension As String

    filename = Left(nowFilename, InStrRev(nowFilename, ".") - 1)
    fileExtension = Mid(nowFilename, InStrRev(nowFilename, ".") + 1, Len(nowFilename))
    newFileNeme = filename & "_" & addString & "." & fileExtension
    
    AddStringFilename = newFileNeme
    
End Function
