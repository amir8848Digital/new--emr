VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDirectory2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------
' CDIRECTORY2 Class
'
' This is primarily an example of a class that
' can be showed in a treeview control by implementing
' the IExplorerItem interface
'---------------------------------------------------

Option Explicit

Const CLASS_NAME = "CDirectory2"

Implements IExplorerItem

Public Path As String

' the "name" of the directory is simply the portion of
' the path that follows the last backslash, or the whole
' path if this is a root directory

Property Get Name() As String
    Dim i As Integer
    ' get the last backslash
    i = InstrLast(1, Path, "\")
    If i = Len(Path) Then
        Name = Path
    Else
        Name = Mid$(Path, i + 1)
    End If
End Property

' return the collection of subdirectories

Function Subdirs(Optional justOneIsOK As Boolean) As Collection
    Dim newDir As CDirectory2, path2 As String, dirPath As String
    
    Set Subdirs = New Collection
    
    ' add a backslash to the path if necessary
    path2 = Path & IIF(Right$(Path, 1) <> "\", "\", "")
    
    On Error Resume Next
    
    dirPath = Dir$(path2 & "*.*", vbDirectory)
    Do While Len(dirPath)
        If (GetAttr(path2 & dirPath) And vbDirectory) = 0 Then
            ' it's a file or an error
            ' (we need this because GetAttr raises an error
            '  when applied to WinNT's PAGEFILE.SYS )
        Else
            ' it is a directory, now discard "." and ".."
            If Left$(dirPath, 1) <> "." Then
                Set newDir = New CDirectory2
                newDir.Path = path2 & dirPath
                Subdirs.Add newDir
                ' exit if we don't need more
                If justOneIsOK Then Exit Function
            End If
        End If
        ' get the next item, exit if null string
        dirPath = Dir$()
    Loop
End Function

' returns the collection of files in the directory

Function Files() As Collection
    Dim path2 As String, dirPath As String
    
    Set Files = New Collection
    
    ' add a backslash to the path if necessary
    path2 = Path & IIF(Right$(Path, 1) <> "\", "\", "")
    
    dirPath = Dir$(path2 & "*.*")
    Do While Len(dirPath)
        Files.Add dirPath
        ' get the next item, exit if null string
        dirPath = Dir$()
    Loop
End Function

' get the last occurrence of a string

Private Function InstrLast(ByVal Start As Long, ByVal Source As String, ByVal Search As String, Optional Compare As VbCompareMethod) As Long
    Start = Start - 1
    Do
        Start = InStr(Start + 1, Source, Search, Compare)
        If Start = 0 Then Exit Function
        InstrLast = Start
    Loop
End Function

'---------------------------------------------------
' The IExplorerItem interface
'---------------------------------------------------

Private Property Get IExplorerItem_Text() As String
    IExplorerItem_Text = Name
End Property

Private Property Let IExplorerItem_Text(ByVal RHS As String)
    ' prevent from renaming a directory
    Err.Raise 999
End Property

Private Function IExplorerItem_HasChildren() As Boolean
    IExplorerItem_HasChildren = Subdirs(True).Count
End Function

Private Function IExplorerItem_GetChildren() As Collection
    Set IExplorerItem_GetChildren = Subdirs(False)
End Function

Private Sub IExplorerItem_ShowItems(Explorer As CExplorer2)
    Dim file As Variant
    Explorer.ClearItems
    For Each file In Files
        Explorer.AddItem CStr(file)
    Next
End Sub

Private Sub IExplorerItem_Delete()
    ' refuse deletion by raising an error
    Err.Raise 999, CLASS_NAME, "Unable to delete this object"
End Sub


