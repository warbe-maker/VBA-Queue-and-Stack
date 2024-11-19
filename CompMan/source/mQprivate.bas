Attribute VB_Name = "mQprivate"
Option Explicit

Private Function IsInQueue(ByVal i_queue As Collection, _
                           ByVal i_item As Variant, _
                  Optional ByRef i_pos As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the index (i_pos) when the item (i_item) is found in the
' queue (i_queue).
' ----------------------------------------------------------------------------
    Dim i As Long
    
    If VarType(i_item) = vbObject Then
        For i = 1 To i_queue.Count
            If i_queue(i) Is i_item Then
                IsInQueue = True
                i_pos = i
                Exit Function
            End If
        Next i
    Else
        For i = 1 To i_queue.Count
            If i_queue(i) = i_item Then
                IsInQueue = True
                i_pos = i
                Exit Function
            End If
        Next i
    End If

End Function

Private Sub Qdequeue(ByRef q_queue As Collection, _
                     ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the top item in the queue (q_item), i.e. the last one pushed on it,
' and removes it from the queue.
' ----------------------------------------------------------------------------
    Dim Pos As Long
    If Not QisEmpty(q_queue) Then
        Qfirst q_queue, q_item, Pos
        q_queue.Remove Pos
    End If
End Sub

Private Sub Qenqueue(ByRef q_queue As Collection, _
                         ByVal q_item As Variant)
    If q_queue Is Nothing Then Set q_queue = New Collection
    q_queue.Add q_item
End Sub

Private Sub Qfirst(ByVal q_queue As Collection, _
                       ByRef q_item As Variant, _
              Optional ByRef q_pos As Long)
' ----------------------------------------------------------------------------
' Returns the current first item in the queue, i.e. the one added (enqueued)
' first.
' ----------------------------------------------------------------------------
    If Not QisEmpty(q_queue) Then
        q_pos = 1
        If VarType(q_queue(q_pos)) = vbObject Then
            Set q_item = q_queue(q_pos)
        Else
            q_item = q_queue(q_pos)
        End If
    End If
End Sub

Private Function QisEmpty(ByVal q_queue As Collection) As Boolean
    QisEmpty = q_queue Is Nothing
    If Not QisEmpty Then QisEmpty = q_queue.Count = 0
End Function

Private Sub Qitem(ByVal q_queue As Collection, _
                      ByVal q_pos As Long, _
             Optional ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (q_item) at the position (q_pos) in the queue (q_queue),
' provided the queue is not empty and the position is within the queue's size.
' ----------------------------------------------------------------------------
    
    If Not QisEmpty(q_queue) Then
        If q_pos <= Qsize(q_queue) Then
            If VarType(q_queue(q_pos)) = vbObject Then
                Set q_item = q_queue(q_pos)
            Else
                q_item = q_queue(q_pos)
            End If
        End If
    End If
    
End Sub

Private Sub Qlast(ByVal q_queue As Collection, _
                      ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (q_item) in the queue which had been enqueued last in the
' queue (q_queue), provided the queue is not empty.
' ----------------------------------------------------------------------------
    Dim lSize As Long
    
    If Not QisEmpty(q_queue) Then
        lSize = q_queue.Count
        If VarType(q_queue(lSize)) = vbObject Then
            Set q_item = q_queue(lSize)
        Else
            q_item = q_queue(lSize)
        End If
    End If
End Sub

Private Function Qsize(ByVal q_queue As Collection) As Long
    If Not QisEmpty(q_queue) Then Qsize = q_queue.Count
End Function

Public Sub Test_Private_Queue_Services()
' ----------------------------------------------------------------------------
' Self-test for the 'Private' Queue.... services
' ----------------------------------------------------------------------------
    Dim MyQueue As New Collection
    Dim Item    As Variant
    Dim Pos     As Long
                                Debug.Assert QisEmpty(MyQueue)
    Qenqueue MyQueue, "A":  Debug.Assert Not QisEmpty(MyQueue)
    Qdequeue MyQueue, Item: Debug.Assert Item = "A"
                                Debug.Assert QisEmpty(MyQueue)
    Qenqueue MyQueue, "A"
    Qenqueue MyQueue, "B"
    Qenqueue MyQueue, "C"
    Qenqueue MyQueue, "D"
                                Debug.Assert Not QisEmpty(MyQueue)
                                Debug.Assert Qsize(MyQueue) = 4
                                Debug.Assert IsInQueue(MyQueue, "B", Pos) = True
                                Debug.Assert Pos = 2
    Qitem MyQueue, 2, Item: Debug.Assert Item = "B"
    Qdequeue MyQueue, Item: Debug.Assert Item = "A"
    Qdequeue MyQueue, Item: Debug.Assert Item = "B"
    Qdequeue MyQueue, Item: Debug.Assert Item = "C"
    Qdequeue MyQueue, Item: Debug.Assert Item = "D"
                                Debug.Assert QisEmpty(MyQueue)
    Set MyQueue = Nothing
    
End Sub

