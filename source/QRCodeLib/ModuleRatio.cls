VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModuleRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PreLightRatio4   As Long
Public PreDarkRatio1    As Long
Public PreLightRatio1   As Long
Public CenterDarkRatio3 As Long
Public FolLightRatio1   As Long
Public FolDarkRatio1    As Long
Public FolLightRatio4   As Long

Public Function PenaltyImposed() As Boolean

    If PreDarkRatio1 = 0 Then
        PenaltyImposed = False
        Exit Function
    End If

    If PreDarkRatio1 = PreLightRatio1 And _
       PreDarkRatio1 = FolLightRatio1 And _
       PreDarkRatio1 = FolDarkRatio1 And _
       PreDarkRatio1 * 3 = CenterDarkRatio3 Then

        PenaltyImposed = PreLightRatio4 >= PreDarkRatio1 * 4 Or _
                         FolLightRatio4 >= PreDarkRatio1 * 4
    Else
        PenaltyImposed = False
    End If

End Function
