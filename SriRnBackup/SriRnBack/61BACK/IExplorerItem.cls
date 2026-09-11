VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IExplorerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------
' IExplorerItem Abstract Class
'
' Classes that support this interface can be assigned
' to the Root property of the TreeView class
'------------------------------------------------------

Option Explicit

' the string that must be showed on the TreeView
' the object should refuse invalid assignments
Public text As String

' returns True if this node has any children

Function HasChildren() As Boolean
    '
End Function

' returns the collection of children

Function GetChildren() As Collection
    '
End Function

' show the properties

Sub ShowItems(Explorer As CExplorer2)
    '
End Sub

' delete the object

Sub Delete()
    '
End Sub

