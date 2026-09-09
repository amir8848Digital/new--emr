VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExplorer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------
'  CEXPLORER2 Class
'-------------------------------------------------

Option Explicit

Const CLASS_NAME = "CExplorer2"

Event GetChildren(Data As Variant)
Event GetItems(Data As Variant)
Event RenameNode(Data As Variant, NewName As String, Cancel As Integer)

' public boolean properties
Public AllowRename As Boolean
Public AllowDelete As Boolean

' allocation unit for m_ItemData()
Const ITEMDATA_CHUNK = 50
' text used for dummy children
Const DUMMY_CHILD = "dummy99999999"

Private Type TItemData
    Node As ComctlLib.Node
    Data As Variant
End Type

' the attached TreeView and ListView controls
Private WithEvents TV As ComctlLib.TreeView
Attribute TV.VB_VarHelpID = -1
Private WithEvents LV As ComctlLib.ListView
Attribute LV.VB_VarHelpID = -1

' the default icon for TreeView items (can be Missing)
Private m_Image As Variant
' the default icon for selected TreeView items (can be Missing)
Private m_SelectedImage As Variant
' the default icon for ListView items (can be Missing)
Private m_ItemImage As Variant

' this array stores the ItemData values
Dim m_ItemData() As TItemData
' the number of used items in the array
Private m_ItemDataCount As Long

Private m_Root As IExplorerItem

Property Get Root() As IExplorerItem
    Set Root = m_Root
End Property

Property Set Root(NewValue As IExplorerItem)
    Set m_Root = NewValue
    ClearNodes
    AddNode NewValue.text, NewValue, NewValue.HasChildren
End Property

' initialize the class with attached TreeView and ListView controls

Sub Init(TreeView As ComctlLib.TreeView, Optional ListView As ComctlLib.ListView, Optional Image As Variant, Optional SelectedImage As Variant, Optional ItemImage As Variant)
    ' storeinfo about the attached TreeView control
    Set TV = TreeView
    m_Image = Image
    m_SelectedImage = SelectedImage
    TV.LabelEdit = tvwManual
    
    ' storeinfo about the attached ListView control, if not omitted
    Set LV = ListView
    m_ItemImage = ItemImage
    
    ' clear the two controls
    ClearNodes
    If Not (LV Is Nothing) Then ClearItems
End Sub

'--------------------------------------------
'  TreeView management
'--------------------------------------------

' clear all nodes

Sub ClearNodes()
    TV.Nodes.Clear
    ReDim m_ItemData(0) As TItemData
    m_ItemDataCount = 0
End Sub

' add a new node

Function AddNode(text As String, Optional Data As Variant, Optional HasChildren As Boolean, Optional parent As Variant, Optional Image As Variant, Optional SelectedImage As Variant) As ComctlLib.Node
    Dim newNode As ComctlLib.Node, parentNode As ComctlLib.Node
    
    ' provide default for missing arguments
    If IsMissing(Image) Then Image = m_Image
    If IsMissing(SelectedImage) Then SelectedImage = Image
    
    If IsMissing(parent) Then
        ' add as root node
        Set newNode = TV.Nodes.Add(, , , text, Image, SelectedImage)
    Else
        ' add as a child node, so we need to retrieve the parent node
        Set parentNode = m_ItemData(GetItemDataIndex(parent)).Node
        ' following code will raise error if parentNode hasn't been found
        If parentNode.children = 1 Then
            ' if the parent has a dummy child, delete it now
            If parentNode.Child.text = DUMMY_CHILD Then
                TV.Nodes.Remove parentNode.Child.Index
            End If
        End If
        Set newNode = TV.Nodes.Add(parentNode.Index, tvwChild, , text, Image, SelectedImage)
    End If
    Set AddNode = newNode
    
    ' store the companion data
    If Not IsMissing(Data) Then
        m_ItemDataCount = m_ItemDataCount + 1
        If m_ItemDataCount > UBound(m_ItemData) Then
            ReDim Preserve m_ItemData(UBound(m_ItemData) + ITEMDATA_CHUNK) As TItemData
        End If
        Set m_ItemData(m_ItemDataCount).Node = newNode
        ' objects and non-objects must be dealt with separatedly
        If IsObject(Data) Then
            Set m_ItemData(m_ItemDataCount).Data = Data
        Else
            m_ItemData(m_ItemDataCount).Data = Data
        End If
    End If
    
    ' if this node has children, add a dummy node just to show the "+" symbol
    If HasChildren Then
        AddPlusSign newNode
    End If
End Function

' ensure that there is a "+" symbol besides the node

Private Sub AddPlusSign(Node As ComctlLib.Node)
    If Node.children = 0 Then
        TV.Nodes.Add Node.Index, tvwChild, , DUMMY_CHILD
    End If
End Sub

' remove a node
' the argument may be the node itself or its companion item data

Sub RemoveNode(Item As Variant)
    Dim dataIndex As Long, nodeIndex As Long
    
    If TypeOf Item Is ComctlLib.Node Then
        ' it is a TreeView node
        nodeIndex = Item.Index
        dataIndex = GetNodeIndex(TV.Nodes(nodeIndex))
    Else
        ' it is something else - search in the array
        dataIndex = GetItemDataIndex(Item)
        If dataIndex = 0 Then Exit Sub
        nodeIndex = m_ItemData(dataIndex).Node.Index
    End If
    
    ' if there is an associated element in ItemData, delete it
    If dataIndex Then
        ' check if the object supports the IExplorerItem interface
        On Error Resume Next
        Err.Clear
        Dim tvItem As IExplorerItem
        Set tvItem = m_ItemData(dataIndex).Data
        If Err = 0 Then
            ' the interface is supported, now inform the
            ' object that we are going to delete it
            tvItem.Delete
            ' if the object refuses to be deleted, exit
            If Err Then Exit Sub
        End If
    
        ' delete the item by moving the last one in this position
        m_ItemData(dataIndex) = m_ItemData(m_ItemDataCount)
        m_ItemDataCount = m_ItemDataCount - 1
    End If
    
    ' finally, remove the node itself
    TV.Nodes.Remove nodeIndex
End Sub

' get/let the ItemData associated to a node
' the argument may be the node's index or key, or the node itself

' Get returns empty/nothing if no data is associated to that note

Property Get ItemData(Item As Variant) As Variant
    Dim Node As ComctlLib.Node, dataIndex As Long
    
    If TypeOf Item Is ComctlLib.Node Then
        Set Node = Item
    Else
        ' the argument is the node's key or index
        Set Node = TV.Nodes(Item)
    End If
    ' find the corresponding data item
    dataIndex = GetNodeIndex(Node)
    If dataIndex Then
        ' objects and non-objects must be dealt with separatedly
        If IsObject(m_ItemData(dataIndex).Data) Then
            Set ItemData = m_ItemData(dataIndex).Data
        Else
            ItemData = m_ItemData(dataIndex).Data
        End If
    End If
End Property

Property Let ItemData(Item As Variant, NewValue As Variant)
    Dim Node As ComctlLib.Node, dataIndex As Long
    
    If TypeOf Item Is ComctlLib.Node Then
        Set Node = Item
    Else
        ' the argument is the node's key or index
        Set Node = TV.Nodes(Item)
    End If
    ' find the corresponding data item
    dataIndex = GetNodeIndex(Node)
    
    If dataIndex Then
        ' the item already exists
        If IsEmpty(NewValue) Then
            ' we want to delete an existing item
            m_ItemData(dataIndex) = m_ItemData(m_ItemDataCount)
            m_ItemDataCount = m_ItemDataCount - 1
        Else
            ' we want to replace an element
            m_ItemData(dataIndex).Data = NewValue
        End If
    Else
        ' the item doesn't exist - let's add a new element
        If Not IsEmpty(NewValue) Then
            m_ItemDataCount = m_ItemDataCount + 1
            If m_ItemDataCount > UBound(m_ItemData) Then
                ReDim Preserve m_ItemData(UBound(m_ItemData) + ITEMDATA_CHUNK) As TItemData
            End If
            Set m_ItemData(m_ItemDataCount).Node = Node
            m_ItemData(m_ItemDataCount).Data = NewValue
        End If
    End If
    
End Property

Property Set ItemData(Item As Variant, NewValue As Object)
    Dim Node As ComctlLib.Node, dataIndex As Long
    
    If TypeOf Item Is ComctlLib.Node Then
        Set Node = Item
    Else
        ' the argument is the node's key or index
        Set Node = TV.Nodes(Item)
    End If
    ' find the corresponding data item
    dataIndex = GetNodeIndex(Node)
    
    If dataIndex Then
        ' the item already exists
        If NewValue Is Nothing Then
            ' if we want to delete an existing item
            m_ItemData(dataIndex) = m_ItemData(m_ItemDataCount)
            m_ItemDataCount = m_ItemDataCount - 1
        Else
            ' we want to replace an element
            Set m_ItemData(dataIndex).Data = NewValue
        End If
    Else
        ' the item doesn't exist - let's add a new element
        If Not (NewValue Is Nothing) Then
            m_ItemDataCount = m_ItemDataCount + 1
            If m_ItemDataCount > UBound(m_ItemData) Then
                ReDim Preserve m_ItemData(UBound(m_ItemData) + ITEMDATA_CHUNK) As TItemData
            End If
            Set m_ItemData(m_ItemDataCount).Node = Node
            Set m_ItemData(m_ItemDataCount).Data = NewValue
        End If
    End If
    
End Property

' return the node corresponding to a given path

Function GetNodeFromPath(FullPath As String) As ComctlLib.Node
    Dim Node As ComctlLib.Node
    ' perform sequential search (might be optimized)
    For Each Node In TV.Nodes
        If Node.FullPath = FullPath Then
            Set GetNodeFromPath = Node
            Exit Function
        End If
    Next
End Function

' returns the index of the element in ItemData()
' that corresponds to a given item value

Private Function GetItemDataIndex(ItemData As Variant) As Long
    Dim i As Long
    If IsObject(ItemData) Then
        ' return zero if the argument is Nothing
        If ItemData Is Nothing Then Exit Function
        ' search the object in the array
        For i = 1 To m_ItemDataCount
            If m_ItemData(i).Data Is ItemData Then
                GetItemDataIndex = i
                Exit Function
            End If
        Next
    Else
        ' return zero if the argument is null
        If IsEmpty(ItemData) Or IsNull(ItemData) Then Exit Function
        ' search the value in the array
        For i = 1 To m_ItemDataCount
            If m_ItemData(i).Data = ItemData Then
                GetItemDataIndex = i
                Exit Function
            End If
        Next
    End If
End Function

' returns the index of the item in ItemData()
' that corresponds to a given node

Private Function GetNodeIndex(Node As ComctlLib.Node) As Long
    Dim i As Long
    For i = 1 To m_ItemDataCount
        If m_ItemData(i).Node Is Node Then
            GetNodeIndex = i
            Exit Function
        End If
    Next
End Function

'-------------------------------------------------
' TreeView event procedures
'-------------------------------------------------

' before expanding a node, replace any dummy child
' and ask the main program for the real children

Private Sub TV_Expand(ByVal Node As ComctlLib.Node)
    Dim itemIndex As Long, Item As Variant, ParentItem As Variant
    Dim tvItem As IExplorerItem, tvChild As IExplorerItem
    
    ' exit if no need to ask for children
    If Node.children <> 1 Then Exit Sub
    If Node.Child.text <> DUMMY_CHILD Then Exit Sub
    ' remove the dummy child
    TV.Nodes.Remove Node.Child.Index
    
    ' retrieve the corresponding object, exit if not found
    itemIndex = GetNodeIndex(Node)
    If itemIndex = 0 Then Exit Sub
    
    ' we need to create a reference to the parent item in a simple variant
    ' because if we pass the array item to the AddNode method the m_ItemData
    ' array would be locked and it would be prevented from being expanded
    If IsObject(m_ItemData(itemIndex).Data) Then
        Set ParentItem = m_ItemData(itemIndex).Data
    Else
        ParentItem = m_ItemData(itemIndex).Data
    End If
    
    On Error Resume Next
    Err.Clear
    
    ' does the item support the IExplorerItem interface?
    Set tvItem = m_ItemData(itemIndex).Data
    If Err = 0 Then
        ' yes, ask it about its children and add to the treeview control
        For Each Item In tvItem.GetChildren
            ' retrieve the child's secondary interface
            Set tvChild = Item
            ' add to the treeview control
            Dim parent As Variant
            
            AddNode tvChild.text, tvChild, tvChild.HasChildren, ParentItem
        Next
    Else
        ' no, we may only ask the main program to expand the node
        RaiseEvent GetChildren(ParentItem)
    End If
End Sub

' trap key to delete objects

Private Sub TV_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            ' delete the node ony if allowed
            If AllowDelete Then RemoveNode TV.SelectedItem
    End Select
End Sub

' refuse to rename an object, if AllowRename = False

Private Sub TV_BeforeLabelEdit(Cancel As Integer)
    If Not AllowRename Then Cancel = True
End Sub

' inform the program that an item has been renamed

Private Sub TV_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim itemIndex As Long, tvItem As IExplorerItem
    itemIndex = GetNodeIndex(TV.SelectedItem)
    If itemIndex = 0 Then Exit Sub

    On Error Resume Next
    Err.Clear
    
    ' retrieve the IExplorerItem interface
    Set tvItem = m_ItemData(itemIndex).Data
    If Err = 0 Then
        ' the interface is supported - try to modify its text property
        tvItem.text = NewString
        If tvItem.text <> NewString Then Cancel = True
    Else
        ' the interface is not supported, we may only raise an event
        RaiseEvent RenameNode(m_ItemData(itemIndex).Data, NewString, Cancel)
    End If

End Sub

' inform the program that a node has been clicked
' and that information should be showed in the ListView control

Private Sub TV_NodeClick(ByVal Node As ComctlLib.Node)
    Dim itemIndex As Long, tvItem As IExplorerItem
    itemIndex = GetNodeIndex(Node)
    If itemIndex = 0 Then Exit Sub
    
    On Error Resume Next
    Err.Clear
    ' try to retrieve the IExplorerItem interface
    Set tvItem = m_ItemData(itemIndex).Data
    If Err = 0 Then
        ' the interface is supported - ask the object to show its properties
        tvItem.ShowItems Me
    Else
        ' the interface is not supported
        ' ask the program to show the object's properties
        RaiseEvent GetItems(m_ItemData(itemIndex).Data)
    End If
End Sub

'--------------------------------------------
'  ListView management
'--------------------------------------------

' clear all listview items

Sub ClearItems()
    LV.ListItems.Clear
End Sub

' add a new item, and optionally up to 5 subitems

Function AddItem(text As String, ParamArray subitems() As Variant) As ComctlLib.ListItem
    Dim newItem As ComctlLib.ListItem, i As Integer
    
    Set newItem = LV.ListItems.Add(, , text, m_ItemImage, m_ItemImage)
    Set AddItem = newItem
    
    ' add any subitems, if specified
    For i = 0 To UBound(subitems)
        newItem.subitems(i + 1) = subitems(i)
    Next
End Function

' remove an item

Sub RemoveItem(Index As Variant)
    LV.ListItems.Remove Index
End Sub

' set the columns caption and width
' arguments are in pair (caption, width)
' if the width argument of the last pair is omitted, it extends to the right edge of the control

Sub SetColumns(ParamArray args() As Variant)
    Dim i As Long, colWidth As Long
    LV.ColumnHeaders.Clear
    
    For i = LBound(args) To UBound(args) Step 2
        If i < UBound(args) Then
            colWidth = args(i + 1)
        Else
            colWidth = LV.Width - LV.ColumnHeaders(LV.ColumnHeaders.Count).Left - LV.ColumnHeaders(LV.ColumnHeaders.Count).Width - 730
            If colWidth < 2000 Then colWidth = 2000
        End If
        LV.ColumnHeaders.Add , , args(i), colWidth
    Next
End Sub

