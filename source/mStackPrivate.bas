Attribute VB_Name = "mStackPrivate"
Option Explicit

Private Function IsOnStack(ByVal s_stack As Collection, _
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
                IsOnStack = True
                s_pos = i
                Exit Function
            End If
        Next i
    Else
        For i = 1 To s_stack.Count
            If s_stack(i) = s_item Then
                IsOnStack = True
                s_pos = i
                Exit Function
            End If
        Next i
    End If

End Function

Private Sub StackBottom(ByVal s_stack As Collection, _
                        ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the bottom item (s_item) on the stack (s_stack), provided the stack
' is not empty.
' ----------------------------------------------------------------------------
    Dim lBottom As Long
    
    If Not StackIsEmpty(s_stack) Then
        lBottom = s_stack.Count
        If VarType(s_stack(lBottom)) = vbObject Then
            Set s_item = s_stack(lBottom)
        Else
            s_item = s_stack(lBottom)
        End If
    End If
End Sub

Private Function StackIsEmpty(ByVal s_stack As Collection) As Boolean
    StackIsEmpty = s_stack Is Nothing
    If Not StackIsEmpty Then StackIsEmpty = s_stack.Count = 0
End Function

Private Sub StackItem(ByVal s_stack As Collection, _
                      ByVal s_pos As Long, _
             Optional ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (s_item) at the position (s_pos) on the stack (s_stack),
' provided the stack is not empty and the position is within the stack's size.
' ----------------------------------------------------------------------------
    
    If Not StackIsEmpty(s_stack) Then
        If s_pos <= StackSize(s_stack) Then
            If VarType(s_stack(s_pos)) = vbObject Then
                Set s_item = s_stack(s_pos)
            Else
                s_item = s_stack(s_pos)
            End If
        End If
    End If
    
End Sub

Private Sub StackPop(ByRef s_stack As Collection, _
                     ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the top item on the stack (s_item), i.e. the last one pushed on it,
' and removes it from the stack.
' ----------------------------------------------------------------------------
    Dim Pos As Long
    If Not StackIsEmpty(s_stack) Then
        StackTop s_stack, s_item, Pos
        s_stack.Remove Pos
    End If
End Sub

Private Sub StackPush(ByRef s_stack As Collection, _
                      ByVal s_item As Variant)
    If s_stack Is Nothing Then Set s_stack = New Collection
    s_stack.Add s_item
End Sub

Private Function StackSize(ByVal s_stack As Collection) As Long
    If Not StackIsEmpty(s_stack) Then StackSize = s_stack.Count
End Function

Private Sub StackTop(ByVal s_stack As Collection, _
                     ByRef s_item As Variant, _
            Optional ByRef s_pos As Long)
' ----------------------------------------------------------------------------
' Returns the top item on the stack (s_item), i.e. the last one pushed on it
' and the top item's position.
' ----------------------------------------------------------------------------
    
    If Not StackIsEmpty(s_stack) Then
        s_pos = s_stack.Count
        If VarType(s_stack(s_pos)) = vbObject Then
            Set s_item = s_stack(s_pos)
        Else
            s_item = s_stack(s_pos)
        End If
    End If
End Sub

Public Sub Test_Private_Stack_Services()
' ----------------------------------------------------------------------------
' Self-test for the 'Private' Stack.... services
' ----------------------------------------------------------------------------
    Dim MyStack As New Collection
    Dim Item    As Variant
    Dim Pos     As Long
                                Debug.Assert StackIsEmpty(MyStack)
    StackPush MyStack, "A":     Debug.Assert Not StackIsEmpty(MyStack)
    StackPop MyStack, Item:     Debug.Assert Item = "A"
                                Debug.Assert StackIsEmpty(MyStack)
    StackPush MyStack, "A"
    StackPush MyStack, "B"
    StackPush MyStack, "C"
    StackPush MyStack, "D"
                                Debug.Assert Not StackIsEmpty(MyStack)
                                Debug.Assert StackSize(MyStack) = 4
                                Debug.Assert IsOnStack(MyStack, "B", Pos) = True
                                Debug.Assert Pos = 2
    StackItem MyStack, 2, Item: Debug.Assert Item = "B"
    StackPop MyStack, Item:     Debug.Assert Item = "D"
    StackPop MyStack, Item:     Debug.Assert Item = "C"
    StackPop MyStack, Item:     Debug.Assert Item = "B"
    StackPop MyStack, Item:     Debug.Assert Item = "A"
                                Debug.Assert StackIsEmpty(MyStack)
    Set MyStack = Nothing
    
End Sub

