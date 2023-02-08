Attribute VB_Name = "mStkPrivate"
Option Explicit

Private Function StkOn(ByVal s_stack As Collection, _
                           ByVal s_item As Variant, _
                  Optional ByRef s_pos As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the index (s_pos) when the item (s_item) is found in the
' stack (s_stack).
' ----------------------------------------------------------------------------
    Dim i   As Long
    
    If VarType(s_item) = vbObject Then
        For i = 1 To s_stack.Count
            If s_stack(i) Is s_item Then
                StkOn = True
                s_pos = i
                Exit Function
            End If
        Next i
    Else
        For i = 1 To s_stack.Count
            If s_stack(i) = s_item Then
                StkOn = True
                s_pos = i
                Exit Function
            End If
        Next i
    End If

End Function

Private Sub StkBottom(ByVal s_stack As Collection, _
                        ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the bottom item (s_item) on the stack (s_stack), provided the stack
' is not empty.
' ----------------------------------------------------------------------------
    Dim lBottom As Long
    
    If Not StkIsEmpty(s_stack) Then
        lBottom = s_stack.Count
        If VarType(s_stack(lBottom)) = vbObject Then
            Set s_item = s_stack(lBottom)
        Else
            s_item = s_stack(lBottom)
        End If
    End If
End Sub

Private Function StkIsEmpty(ByVal s_stack As Collection) As Boolean
    StkIsEmpty = s_stack Is Nothing
    If Not StkIsEmpty Then StkIsEmpty = s_stack.Count = 0
End Function

Private Sub StkItem(ByVal s_stack As Collection, _
                      ByVal s_pos As Long, _
             Optional ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (s_item) at the position (s_pos) on the stack (s_stack),
' provided the stack is not empty and the position is within the stack's size.
' ----------------------------------------------------------------------------
    
    If Not StkIsEmpty(s_stack) Then
        If s_pos <= StkSize(s_stack) Then
            If VarType(s_stack(s_pos)) = vbObject Then
                Set s_item = s_stack(s_pos)
            Else
                s_item = s_stack(s_pos)
            End If
        End If
    End If
    
End Sub

Private Sub StkPop(ByRef s_stack As Collection, _
                     ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the top item on the stack (s_item), i.e. the last one pushed on it,
' and removes it from the stack.
' ----------------------------------------------------------------------------
    Dim Pos As Long
    If Not StkIsEmpty(s_stack) Then
        StkTop s_stack, s_item, Pos
        s_stack.Remove Pos
    End If
End Sub

Private Sub StkPush(ByRef s_stack As Collection, _
                      ByVal s_item As Variant)
    If s_stack Is Nothing Then Set s_stack = New Collection
    s_stack.Add s_item
End Sub

Private Function StkSize(ByVal s_stack As Collection) As Long
    If Not StkIsEmpty(s_stack) Then StkSize = s_stack.Count
End Function

Private Sub StkTop(ByVal s_stack As Collection, _
                     ByRef s_item As Variant, _
            Optional ByRef s_pos As Long)
' ----------------------------------------------------------------------------
' Returns the top item on the stack (s_item), i.e. the last one pushed on it
' and the top item's position.
' ----------------------------------------------------------------------------
    
    If Not StkIsEmpty(s_stack) Then
        s_pos = s_stack.Count
        If VarType(s_stack(s_pos)) = vbObject Then
            Set s_item = s_stack(s_pos)
        Else
            s_item = s_stack(s_pos)
        End If
    End If
End Sub

Public Sub Test_Private_Stk_Services()
' ----------------------------------------------------------------------------
' Self-test for the 'Private' Stk.... services
' ----------------------------------------------------------------------------
    Dim MyStk As New Collection
    Dim Item    As Variant
    Dim Pos     As Long
                                Debug.Assert StkIsEmpty(MyStk)
    StkPush MyStk, "A":     Debug.Assert Not StkIsEmpty(MyStk)
    StkPop MyStk, Item:     Debug.Assert Item = "A"
                                Debug.Assert StkIsEmpty(MyStk)
    StkPush MyStk, "A"
    StkPush MyStk, "B"
    StkPush MyStk, "C"
    StkPush MyStk, "D"
                                Debug.Assert Not StkIsEmpty(MyStk)
                                Debug.Assert StkSize(MyStk) = 4
                                Debug.Assert StkOn(MyStk, "B", Pos) = True
                                Debug.Assert Pos = 2
    StkItem MyStk, 2, Item: Debug.Assert Item = "B"
    StkPop MyStk, Item:     Debug.Assert Item = "D"
    StkPop MyStk, Item:     Debug.Assert Item = "C"
    StkPop MyStk, Item:     Debug.Assert Item = "B"
    StkPop MyStk, Item:     Debug.Assert Item = "A"
                                Debug.Assert StkIsEmpty(MyStk)
    Set MyStk = Nothing
    
End Sub

