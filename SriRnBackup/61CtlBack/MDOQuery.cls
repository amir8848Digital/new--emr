VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDOQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim wo_param As ADODB.Parameter
Dim mo_Command As ADODB.Command
Dim mo_adoCommand As ADODB.Command
Public Property Set adoCommand(vnewvalue As ADODB.Command)
    Set mo_adoCommand = vnewvalue
End Property
Public Property Let QryParameters(ByVal ParamNo As Integer, ByVal vnewvalue As Variant)
    mo_adoCommand.Parameters(ParamNo) = vnewvalue
'??? conv to datatype reqd ???
End Property
Public Property Get QryParameters(ByVal ParamNo As Integer) As Variant
    QryParameters = mo_adoCommand.Parameters(ParamNo)
'??? conv to datatype reqd ???
End Property
Public Sub QryExecute()
    mo_adoCommand.Execute
End Sub

Public Sub QryClose()
'  mo_adoCommand.Close
  Set mo_adoCommand.ActiveConnection = Nothing
  Set mo_adoCommand = Nothing
End Sub

Public Property Let QryParametersByName(ByVal ParamName As String, ByVal vnewvalue As Variant)
    mo_adoCommand.Parameters(ParamName) = vnewvalue
'??? conv to datatype reqd ???
End Property

Public Property Get QryParametersByName(ByVal ParamName As String) As Variant
    QryParametersByName = mo_adoCommand.Parameters(ParamName)
'??? conv to datatype reqd ???
End Property

