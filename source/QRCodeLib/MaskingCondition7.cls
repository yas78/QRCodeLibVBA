VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaskingCondition7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IMaskingCondition

Private Function IMaskingCondition_Evaluate(ByVal r As Long, ByVal c As Long) As Boolean
    IMaskingCondition_Evaluate = ((r + c) Mod 2 + (r * c) Mod 3) Mod 2 = 0
End Function
