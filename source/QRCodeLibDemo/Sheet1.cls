VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub btnOpenForm_Click()
    
    Dim frm As Form1
    Set frm = New Form1
    
    Call frm.Show(vbModeless)

End Sub
