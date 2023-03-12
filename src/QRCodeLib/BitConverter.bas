Attribute VB_Name = "BitConverter"
Option Private Module
Option Explicit

Public Function GetBytes(ByVal arg As Variant, Optional ByVal reverse As Boolean = False) As Byte()
    Dim ret() As Byte
    Dim temp As Byte

    Select Case VarType(arg)
        Case VbVarType.vbByte
            ReDim ret(0)
            ret(0) = arg
        Case VbVarType.vbInteger
            ReDim ret(1)
            ret(0) = arg And &HFF&
            ret(1) = (arg And &HFF00&) \ 2 ^ 8

            If reverse Then
                temp = ret(0)
                ret(0) = ret(1)
                ret(1) = temp
            End If
        Case VbVarType.vbLong
            ReDim ret(3)
            ret(0) = arg And &HFF&
            ret(1) = (arg And &HFF00&) \ 2 ^ (8 * 1)
            ret(2) = (arg And &HFF0000) \ 2 ^ (8 * 2)
            ret(3) = (arg And &HFF000000) \ 2 ^ (8 * 3) And &HFF&

            If reverse Then
                temp = ret(0)
                ret(0) = ret(3)
                ret(3) = temp

                temp = ret(1)
                ret(1) = ret(2)
                ret(2) = temp
            End If
#If Win64 Then
        Case VbVarType.vbLongLong
            ReDim ret(7)
            ret(0) = arg And &HFF&
            ret(1) = (arg And &HFF00&) \ 2 ^ (8 * 1)
            ret(2) = (arg And &HFF0000) \ 2 ^ (8 * 2)
            ret(3) = (arg And &HFF000000^) \ 2 ^ (8 * 3)
            ret(4) = (arg And &HFF00000000^) \ 2 ^ (8 * 4)
            ret(5) = (arg And &HFF0000000000^) \ 2 ^ (8 * 5)
            ret(6) = (arg And &HFF000000000000^) \ 2 ^ (8 * 6)
            ret(7) = (arg And &HFF00000000000000^) \ 2 ^ (8 * 7) And &HFF&

            If reverse Then
                temp = ret(0)
                ret(0) = ret(7)
                ret(7) = temp

                temp = ret(1)
                ret(1) = ret(6)
                ret(6) = temp

                temp = ret(2)
                ret(2) = ret(5)
                ret(5) = temp

                temp = ret(3)
                ret(3) = ret(4)
                ret(4) = temp
            End If
#End If
        Case Else
            Call Err.Raise(5)
    End Select

    GetBytes = ret
End Function
