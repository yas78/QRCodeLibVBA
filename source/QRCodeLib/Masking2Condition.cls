VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Masking2Condition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMaskingCondition

Private Function IMaskingCondition_Evaluate(ByVal r As Long, ByVal c As Long) As Boolean
    
    IMaskingCondition_Evaluate = c Mod 3 = 0

End Function
