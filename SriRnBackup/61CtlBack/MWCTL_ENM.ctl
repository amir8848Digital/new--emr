VERSION 5.00
Begin VB.UserControl MWCTL_ENM 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "MWCTL_ENM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum en_FldDataType
    fdtCharType = 0
    fdtNumber = 1
    fdtFloat = 2
    fdtDateType = 3
    fdtConvertToUpper = 4
    fdtTimeType = 5
End Enum
Public Enum en_TxtAlign
    tLeftAlign = 1
    tRightAlign = 2
    tMultiLine = 3
End Enum
Public Enum en_HlpMode
    hmHlpAvlbl = 0
    hmHlpNotavlbl = 1
    hmHlpBrowsing = 2
End Enum
Public Enum en_GrpLevelType
    gltGroup1 = 0
    gltGroup2 = 1
    gltGroup3 = 2
    gltsort1 = 3
    gltsort2 = 4
    gltsort3 = 5
End Enum
Public Enum en_fgAction
    fgaOthers = 0
    fgaShiftTabFrmCol = 1
    fgaTabFrmCol = 2
    fgaViaCode = 3
End Enum
Public Enum en_fgMode
    fgminit = 1
    fgmdel = 2
    fgmInitOver = 3
    fgmnorm = 0
    fgmCopy = 4
    fgmCopyOver = 5
    fgmServerCopy = 6
End Enum
Public Enum en_RepDstOpt
    reponscreen
    repInPrint
    repInFile
    repInExport
End Enum
Public Enum en_XMode
    xInit = 0
    xNorm = 1
    xClear = 2
    xCopy = 3
End Enum
Public Enum en_XEntType
    xetEntry = 0
    xetReport = 1
End Enum
Public Enum en_XDeleteMode
    xdelNormal = 0
    xdelChildRecs = 1
End Enum
Public Enum en_XDataMode
    xdatBound = 0
    xdatUnbound = 1
End Enum
Public Enum en_XAddFndMode
    xaddmode = 0
    xFndMode = 1
End Enum
Public Enum en_AddSaveOpt
    asAll = 0
    asCurrent = 1
End Enum
Public Enum en_ErrorType
    etWarning = 0
    etInfo = 1
    etError = 2
End Enum
Public Enum en_fgColStyle
    fgcinvisible = 0
    fgcNormal = 1
    fgcReadOnly = 2
End Enum


'*** Jenny
Public Enum en_GrdLineStyle
    flxGrdNone = 0
    flxGrdFlat = 1
    flxGrdInset = 2
    flxGrdRaised = 3
End Enum
Public Enum en_GrdAppearance
    flxGrdAppFlat = 0
    flxGrdApp3D = 1
End Enum
'*** Jenny

