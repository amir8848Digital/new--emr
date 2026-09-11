VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDORowSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum EnMDOFldType
    mdoftString = 0
    mdoftInt = 1
    mdoftFloat = 2
    mdoftDateTime = 3
    mdoftNumeric = 4
End Enum
Public Enum EnMDOEditMode
    mdoemAddNew = 0
    mdoemEditInProgress = 1
    mdoemDelete = 2
    mdoemNone = 3
End Enum
Public Enum EnMDOResyncOpt
    mdoryAll = 0
    mdoryCurrent = 1
End Enum
Dim m_adorecordset As ADODB.Recordset

Dim me_EditMode As EnMDOEditMode
Public Type tFldDets
    fdFldName As String
    fdFldType As EnMDOFldType
    fdFldSize As Integer
End Type
Public Property Get Bookmark() As Double
    Bookmark = m_adorecordset.Bookmark
End Property
Public Property Let Bookmark(ByVal vnewvalue As Double)
    If me_EditMode <> mdoemNone Then err.Raise vbObjectError + 512, "", "RowSet not in mdoemNone mode.Cannot do this operation"
    m_adorecordset.Bookmark = CDbl(vnewvalue)
End Property
Public Property Set adorecordset(vnewvalue As ADODB.Recordset)
    Set m_adorecordset = vnewvalue
End Property
Public Property Get RecCount() As Long
    RecCount = m_adorecordset.RecordCount
End Property
Public Property Get EOF() As Boolean
    EOF = m_adorecordset.EOF
End Property
Public Property Get BOF() As Boolean
    BOF = m_adorecordset.BOF
End Property
Public Property Get FldValue(ByVal mFldName As Variant) As Variant
Attribute FldValue.VB_UserMemId = 0
    FldValue = m_adorecordset.Fields(mFldName)
End Property
Public Property Let FldValue(ByVal mFldName As Variant, ByVal vnewvalue As Variant)
'On Error GoTo errh
    m_adorecordset.Fields(mFldName) = vnewvalue
 '   Exit Property
'errh:
'    Exit Property
'???
End Property

'*** Jenny
Public Property Get FldPicValue(ByVal mFldName As Variant) As Picture
    Set FldPicValue = ColumnToFile(mFldName)
End Property
Public Property Let FldPicValue(ByVal mFldName As Variant, ByVal vnewvalue As Picture)
    FileToColumn mFldName, vnewvalue
End Property
Private Function ColumnToFile(ByVal mFldName As Variant) As Picture
  '*** MSDN Help Topic
  'HOWTO: Use GetChunk and AppendChunk Methods of RDO Object
  '*** MSDN Help Topic
  Dim Col As ADODB.Field
  Dim DiskFile As String
  Dim BlockSize As Long
  Dim ColSize As Long
  BlockSize = 102400
  'DiskFile = App.Path & "\ImageColToFile.bmp"
  DiskFile = "C:\ImageColToFile.bmp"
  Set Col = m_adorecordset.Fields(mFldName)
  ColSize = m_adorecordset(mFldName).ActualSize
  
  Dim NumBlocks As Integer
  Dim LeftOver As Long
  Dim ByteData() As Byte   '*** Byte array for LongVarBinary
  Dim DestFileNum As Integer, i As Integer

  '*** Remove any existing destination file
  If Len(Dir$(DiskFile)) > 0 Then
    Kill DiskFile
  End If

  DestFileNum = FreeFile
  
  Open DiskFile For Binary As DestFileNum
  'Open DiskFile For Binary Access Write As DestFileNum
  'Open DiskFile For Binary As #DestFileNum
  'Open DiskFile For Binary Access Read Lock Write As #DestFileNum

  NumBlocks = ColSize \ BlockSize
  LeftOver = ColSize Mod BlockSize
  Select Case Col.Type
    Case adLongVarBinary
    'adVarBinary
      ReDim ByteData(0, LeftOver)   '***
      'Debug.Print mFldName
      'If NumBlocks = 0 Then
      '  ByteData() = Col.OriginalValue
      'Else
        ByteData() = Col.GetChunk(LeftOver)
      'End If
      'Debug.Print "After: " + mFldName
      Put DestFileNum, , ByteData()
      For i = 1 To NumBlocks
        ReDim ByteData(0, BlockSize)  '***
        ByteData() = Col.GetChunk(BlockSize)
        Put DestFileNum, , ByteData()
      Next i
    Case Else
      Set ColumnToFile = LoadPicture()
  End Select
  'Close #DestFileNum
  Close DestFileNum

  On Error GoTo ErrorHandler
  Set ColumnToFile = LoadPicture(DiskFile)
  
  GoTo PicCleanUp
  'Exit Function   '*** Exit function to avoid error handler.
ErrorHandler:   ' Error-handling routine.
  Select Case err.Number   ' Evaluate error number.
    Case 481, 50003
      '*** 481:   Invalid Picture error
      '*** 50003: Loading a GIF file that does not specifically contain an Image Object into Visual Basic (for example, through the PictureBox control), generates the run-time error 50003
      Set ColumnToFile = LoadPicture()
    Case Else
      '*** Handle other situations here....
      MsgBox "Invalid Picture File"
  End Select
  
PicCleanUp:
  If Len(Dir$(DiskFile)) > 0 Then
    Kill DiskFile
  End If
  Set Col = Nothing
End Function
Private Sub FileToColumn(ByVal mFldName As Variant, ByVal mPic As Picture)
  '*** moves a disk file to a ChunkRequired column in the table
  
  'Dim Col As ADODB.Field
  Dim BlockSize As Long
  Dim DiskFile As String
  
  '*** A Byte array is used to avoid a UNICODE string
  Dim ByteData() As Byte   '*** Byte array for LongVarBinary
  Dim strData As String    '*** String for LongVarChar
  Dim NumBlocks As Integer
  Dim filelength As Long
  Dim LeftOver As Long
  Dim SourceFile As Integer
  Dim i As Integer
  
  DiskFile = "C:\ImageFileToCol.BMP"
  '*** Remove any existing destination file
  If Len(Dir$(DiskFile)) > 0 Then
    Kill DiskFile
  End If
  If mPic = LoadPicture() Then m_adorecordset.Fields(mFldName).AppendChunk "0": Exit Sub
  SavePicture mPic, DiskFile   '*** Save picture to file.
  BlockSize = 102400
  
  SourceFile = FreeFile
  Open DiskFile For Binary Access Read As SourceFile
  filelength = LOF(SourceFile) ' Get the length of the file
  If filelength = 0 Then
    Close SourceFile
    'MsgBox DiskFile & " empty or not found."
  Else
    '*** Calculate number of blocks to read and left over bytes
    NumBlocks = filelength \ BlockSize
    LeftOver = filelength Mod BlockSize
    'm_adorecordset.Fields(mFldName).AppendChunk Null

    Select Case m_adorecordset.Fields(mFldName).Type
      Case adLongVarBinary
        'adVarBinary
        m_adorecordset.Fields(mFldName).AppendChunk Null
        '*** Read the left over amount of LONGVARBINARY data
        ReDim ByteData(0, LeftOver)
        Get SourceFile, , ByteData()
        m_adorecordset.Fields(mFldName).AppendChunk ByteData()
        ReDim ByteData(0, BlockSize)
        For i = 1 To NumBlocks
          Get SourceFile, , ByteData()
          m_adorecordset.Fields(mFldName).AppendChunk ByteData()
        Next i
        'Close SourceFile
      'Case Else
      '  MsgBox "not a chunkrequired column."
    End Select
    Close SourceFile
  End If
  
End Sub
'*** Jenny
Public Sub WriteToCSVFile(ByVal FileName As String)
  Dim fFile As Integer
  fFile = FreeFile

  Dim i As Long
  Open FileName For Output As #fFile
   Dim strString As String
   'Store header
        For i = 0 To m_adorecordset.Fields.Count - 1
             strString = strString & m_adorecordset.Fields(i).Name & vbTab
        Next
        Print #fFile, strString
         strString = ""
         i = 0
   'Write Records to file
      Do Until m_adorecordset.EOF
          'Loop across all fields in recordset delimiting them with a tab key
           For i = 0 To m_adorecordset.Fields.Count - 1
               strString = strString & m_adorecordset(i) & vbTab
           Next
           'Output the resulting string to the text file
           Print #fFile, strString
           'Reset the string to empty for the next record
           strString = ""
           'Move to the next record in the recordset
           m_adorecordset.MoveNext
      Loop
      'Close the recordset and the text file
      m_adorecordset.Close
      Close #fFile

'Destroy the instance of the recordset from memory
End Sub
Public Property Get FldType(ByVal mFldName As Variant) As EnMDOFldType
    Select Case m_adorecordset.Fields(mFldName).Type
    Case Is = adChar, adVarChar
        FldType = mdoftString
    Case Is = adInteger
        FldType = mdoftInt
    Case Is = adDBTimeStamp
        FldType = mdoftDateTime
    Case Is = adDouble
        FldType = mdoftFloat
    Case Is = adNumeric
        FldType = mdoftNumeric
    End Select
End Property
Public Property Get FldDets(ByVal mFldName As Variant) As tFldDets
    FldDets.fdFldName = m_adorecordset.Fields(mFldName).Name
    FldDets.fdFldSize = m_adorecordset.Fields(mFldName).DefinedSize
    FldDets.fdFldType = FldType(mFldName)
End Property
Public Property Get FldCount() As Long
    FldCount = m_adorecordset.Fields.Count
End Property
Public Sub Update()
    'Debug.Print m_adorecordset!TdIdNo
    m_adorecordset.Update
    me_EditMode = mdoemNone
End Sub

Public Sub Addnew()
    If me_EditMode <> mdoemNone Then err.Raise vbObjectError + 512, "", "RowSet not in mdoemNone mode.Cannot do this operation"
    m_adorecordset.Addnew
    me_EditMode = mdoemAddNew
End Sub
Public Sub edit()
    If me_EditMode <> mdoemNone Then err.Raise vbObjectError + 512, "", "RowSet not in mdoemNone mode.Cannot do this operation"
'    m_adorecordset.edit
    me_EditMode = mdoemEditInProgress
End Sub
Public Sub CancelUpdate()
     If Not (m_adorecordset.BOF Or m_adorecordset.EOF) Then m_adorecordset.CancelUpdate    '*** (Jenny 22/02/08)
    'm_adorecordset.CancelUpdate    '*** (Bef 22/02/08)
    me_EditMode = mdoemNone
End Sub
Public Sub Delete()
    If me_EditMode <> mdoemNone Then err.Raise vbObjectError + 512, "", "RowSet not in mdoemNone mode.Cannot do this operation"
    
    m_adorecordset.Delete
    me_EditMode = mdoemNone
End Sub
Public Sub MoveFirst()
    If me_EditMode <> mdoemNone Then err.Raise vbObjectError + 512, "", "RowSet not in mdoemNone mode.Cannot do this operation"
    m_adorecordset.MoveFirst
End Sub
Public Sub MoveLast()
    If me_EditMode <> mdoemNone Then err.Raise vbObjectError + 512, "", "RowSet not in mdoemNone mode.Cannot do this operation"
    m_adorecordset.MoveLast
End Sub
Public Sub MoveNext()
    If me_EditMode <> mdoemNone Then err.Raise vbObjectError + 512, "", "RowSet not in mdoemNone mode.Cannot do this operation"
    m_adorecordset.MoveNext
End Sub
Public Sub MovePrevious()
    If me_EditMode <> mdoemNone Then err.Raise vbObjectError + 512, "", "RowSet not in mdoemNone mode.Cannot do this operation"
    m_adorecordset.MovePrevious
End Sub
Public Sub Resync(ByVal pv_ResyncOpt As EnMDOResyncOpt)
    Exit Sub
    If pv_ResyncOpt = mdoryAll Then m_adorecordset.Resync adAffectAll
    If pv_ResyncOpt = mdoryCurrent Then m_adorecordset.Resync adAffectCurrent
End Sub
Private Sub Class_Initialize()
    me_EditMode = mdoemNone
End Sub
Private Sub Class_Terminate()
    If Not (m_adorecordset Is Nothing) Then
'        m_adorecordset.Close
        If m_adorecordset.State = adStateOpen Then m_adorecordset.Close
        '??? this check has been temporarily stopped so that u cmoe to know points where
        'mdorowset is not properly closed
    End If
    
    Set m_adorecordset = Nothing
End Sub
Public Sub CloseRes()
    If Not (m_adorecordset Is Nothing) Then
'        m_adorecordset.Close
        If m_adorecordset.State = adStateOpen Then m_adorecordset.Close
        m_adorecordset.ActiveConnection = Nothing
        '??? this check has been temporarily stopped so that u cmoe to know points where
        'mdorowset is not properly closed
    End If
'    Set m_adorecordset = Nothing
End Sub

