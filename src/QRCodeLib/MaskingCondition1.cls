VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaskingCondition1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IMaskingCondition

Private Function IMaskingCondition_Evaluate(ByVal r As Long, ByVal c As Long) As Boolean
    IMaskingCondition_Evaluate = r Mod 2 = 0
End Function
