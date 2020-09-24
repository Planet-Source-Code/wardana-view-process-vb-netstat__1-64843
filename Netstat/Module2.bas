Attribute VB_Name = "Module2"
'To get the properties of file
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

'to send file to recyclebin
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Public Sub ShowProps(FileName As String, OwnerhWnd As Long)
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        'Set the structure's size
        .cbSize = Len(SEI)
        'Seet the mask
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        'Set the owner window
        .hwnd = OwnerhWnd
        'Show the properties
        .lpVerb = "properties"
        'Set the filename
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
     ShellExecuteEx SEI
End Sub

Public Sub RecycleBin(Path As String)
    'to send file to RecycleBin
    Dim SHFileOp As SHFILEOPSTRUCT
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = Path
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation SHFileOp
End Sub

