VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------------------
' Class Module clsQueue: FiFo (queue) services based on a Collection as queue.
'
' Note: A queue can be seen as a tube which is open at both ends. The first
'       item put into it (is enqued) is the first on taken from it (dequeued).
'
' Public services:
' - DeQueue  Returns the first (first added) item from the queue
' - EnQueue  Queues an item
' - First    Returns the first item in the queue without dequeuing it.
' - IsQueued Returns TRUE and its position when a given item is in the queue
' - Item     Returns an item on a provided position in the queue
' - Last     Returns the botton item in the queue
' - Size     Returns the current queue's size
'
' W. Rauschenberger, Berlin Feb 2023
' ----------------------------------------------------------------------------
Private cllQueue As Collection

Private Sub Class_Initialize():                         Set cllQueue = New Collection:          End Sub

Private Sub Class_Terminate():                          Set cllQueue = Nothing:                 End Sub

Public Sub Last(ByRef b_item As Variant):               QueueLast cllQueue, b_item:             End Sub

Public Function IsEmpty() As Boolean:                   IsEmpty = QueueIsEmpty(cllQueue):       End Function

Public Function IsQueued(ByVal i_item As Variant, _
                Optional ByRef i_pos As Long)
    IsQueued = IsInQueue(cllQueue, i_item, i_pos)
End Function

Public Sub Item(ByVal i_pos As Long, _
                ByRef i_item As Variant)
    QueueItem cllQueue, i_pos, i_item
End Sub
                
Public Sub DeQueue(ByRef d_item As Variant):            QueueDequeue cllQueue, d_item:          End Sub

Public Sub EnQueue(ByVal e_item As Variant):            QueueEnqueue cllQueue, e_item:          End Sub

Public Function Size() As Long:                         Size = QueueSize(cllQueue):             End Function

Private Sub QueueLast(ByVal q_queue As Collection, _
                      ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (q_item) in the queue which had been enqueued last in the
' queue (q_queue), provided the queue is not empty.
' ----------------------------------------------------------------------------
    Dim lSize As Long
    
    If Not QueueIsEmpty(q_queue) Then
        lSize = q_queue.Count
        If VarType(q_queue(lSize)) = vbObject Then
            Set q_item = q_queue(lSize)
        Else
            q_item = q_queue(lSize)
        End If
    End If
End Sub

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

Private Sub QueueItem(ByVal q_queue As Collection, _
                      ByVal q_pos As Long, _
             Optional ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (q_item) at the position (q_pos) in the queue (q_queue),
' provided the queue is not empty and the position is within the queue's size.
' ----------------------------------------------------------------------------
    
    If Not QueueIsEmpty(q_queue) Then
        If QueueSize(q_queue) <= q_pos Then
            If VarType(q_queue(q_pos)) = vbObject Then
                Set q_item = q_queue(q_pos)
            Else
                q_item = q_queue(q_pos)
            End If
        End If
    End If
    
End Sub

Private Function QueueIsEmpty(ByVal q_queue As Collection) As Boolean
    QueueIsEmpty = q_queue Is Nothing
    If Not QueueIsEmpty Then QueueIsEmpty = q_queue.Count = 0
End Function

Private Sub QueueDequeue(ByRef q_queue As Collection, _
                     ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the top item in the queue (q_item), i.e. the last one pushed on it,
' and removes it from the queue.
' ----------------------------------------------------------------------------
    Dim pos As Long
    If Not QueueIsEmpty(q_queue) Then
        QueueFirst q_queue, q_item, pos
        q_queue.Remove pos
    End If
End Sub

Private Sub QueueEnqueue(ByRef q_queue As Collection, _
                         ByVal q_item As Variant)
    If q_queue Is Nothing Then Set q_queue = New Collection
    q_queue.Add q_item
End Sub

Private Function QueueSize(ByVal q_queue As Collection) As Long
    If Not QueueIsEmpty(q_queue) Then QueueSize = q_queue.Count
End Function

Private Sub QueueFirst(ByVal q_queue As Collection, _
                       ByRef q_item As Variant, _
              Optional ByRef q_pos As Long)
' ----------------------------------------------------------------------------
' Returns the current first item in the queue, i.e. the one added (enqueued)
' first.
' ----------------------------------------------------------------------------
    If Not QueueIsEmpty(q_queue) Then
        q_pos = 1
        If VarType(q_queue(q_pos)) = vbObject Then
            Set q_item = q_queue(q_pos)
        Else
            q_item = q_queue(q_pos)
        End If
    End If
End Sub

Public Sub First(ByRef t_item)
    QueueFirst cllQueue, t_item
End Sub



