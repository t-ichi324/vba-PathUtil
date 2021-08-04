VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'-----------------------
' auth : t.ichi
'-----------------------
'���O���[�o���f�t�H���g�C���X�^���X��
'
'�uAttribute VB_PredeclaredId = true�v���ݒ肳��Ă��܂��B
'
' �ڐA���̓G�f�B�^����R�s�y�ł͂Ȃ��A
' �G�N�X�|�[�g���e�L�X�g�G�f�B�^��Attribute���m�F���C���|�[�g�̎菇���s���Ă��������B
'

'//////////////////////////////////////
'--- Access�̃t�@�C�������擾
Public Property Get THIS_FILE_NAME()
    THIS_FILE_NAME = Application.CurrentProject.name
End Property
'--- Access�̃t�@�C���̃p�X���擾
Public Property Get THIS_FILE()
    THIS_FILE = Application.CurrentProject.FullName
End Property
'--- Access�̃f�B���N�g���̃p�X���擾
Public Property Get THIS_DIRECTORY()
    THIS_DIRECTORY = Application.CurrentProject.path
End Property

'--- ���O�C�����[�U�[�̃f�X�N�g�b�v
Public Property Get USER_DESKTOP()
    USER_DESKTOP = GetSpecialFolders(4) '("Desktop")
End Property
'--- ���O�C�����[�U�[�̃A�v���p�f�[�^
Public Property Get USER_APPDATA()
    USER_APPDATA = GetSpecialFolders(5) '("AppData")
End Property
'--- ���O�C�����[�U�[�̃}�C�h�L�������g
Public Property Get USER_DOCUMENTS()
    USER_DOCUMENTS = GetSpecialFolders(16) '("MyDocuments")
End Property

'--- ���O�C�����[�U�[�̃v���O�������j���[
Public Property Get USER_PROGRAMS()
    USER_PROGRAMS = GetSpecialFolders(17) '("MyDocuments")
End Property

'//////////////////////////////////////
'--- �t�@�C�����݊m�F
Public Function IsExistsFile(ByVal FullName As String) As Boolean
On Error GoTo ex:
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    IsExistsFile = fs.FileExists(FullName)
ex:
    Set fs = Nothing
End Function

'--- �t�H���_���݊m�F
Public Function IsExistsDirectory(ByVal FullName As String) As Boolean
On Error GoTo ex:
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    IsExistsDirectory = fs.FolderExists(FullName)
ex:
    Set fs = Nothing
End Function

'//////////////////////////////////////
'--- Path�̐��K��
Public Function Normalize(ByVal path As String) As String
    path = Replace(path, "/", "\")

    Dim nDrive As Boolean
    nDrive = (Left(path, 2) = "\\")
    If nDrive Then: path = Mid(path, 3)
    
    Dim col As New Collection
    Dim p As Variant
    For Each p In Split(path, "\")
        p = Trim(CStr(p))
        If p = "" Or p = "." Then
            '�������Ȃ�
        ElseIf p = ".." Then
            If col.Count > 0 Then
                Call col.Remove(col.Count)
            End If
        Else
            Call col.Add(p)
        End If
    Next
    
    Dim ret As String: ret = ""
    For Each p In col
        If ret <> "" Then ret = ret & "\"
        ret = ret & CStr(p)
    Next
    
    Normalize = IIf(nDrive, "\\", "") & ret
End Function

'--- Path�̌���
Public Function Combine(ParamArray paths() As Variant) As String
    Dim ret As String: ret = ""
    Dim isTop As Boolean: isTop = True
    Dim i As Variant
    Dim d As String
    For Each i In paths
        d = CStr(i)
        If Not isTop Then
            If Left(d, 1) = "\" Then
                d = Mid(d, 2)
            End If
        End If
        
        If Right(d, 1) = "\" Then
            ret = ret & d
        Else
            ret = ret & d & "\"
        End If
        
        isTop = False
    Next
    
    If Right(ret, 1) = "\" Then
        ret = Mid(ret, 1, Len(ret) - 1)
    End If
    Combine = ret
End Function

'//////////////////////////////////////
Private Function GetSpecialFolders(ByVal id As Integer) As String
On Error GoTo ex
    GetSpecialFolders = ""
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    GetSpecialFolders = ws.SpecialFolders(id)
ex:
    Set ws = Nothing
End Function