Attribute VB_Name = "mQueue"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mQueue: FiFo (queue) services based on a Collection as queue.
'
' Note: A queue can be seen as a tube which is open at both ends. The first
'       item put into it (is enqued) is the first on taken from it (dequeued).
'
' Public services (in case no queue is provided by the calle the module's
'                  internal default queue is used):
' - DeQueue   Returns the first (first added) item from the queue
' - EnQueue   Queues an item
' - First     Returns the first item in the queue without dequeuing it.
' - IsEmpty   Returns TRUE when the queue is empty
' - IsInQueue Returns TRUE and its position when a given item is in the queue
' - Item      Returns an item on a provided position in the queue
' - Last      Returns the last item enqueued
' - Size      Returns the current queue's size
'
' W. Rauschenberger Berlin Jan 2023
' ----------------------------------------------------------------------------
Private cllQueue As New Collection

Public Sub First(ByRef f_item As Variant, _
        Optional ByRef f_queue As Collection = Nothing)
' ------------------------------------------------------------------------------
' Returns the first item in the queue without dequeuing it.
' ------------------------------------------------------------------------------
    QueueFirst UsedQueue(f_queue), f_item
End Sub

Public Sub Last(ByRef l_item As Variant, _
       Optional ByRef l_queue As Collection = Nothing)
' ------------------------------------------------------------------------------
' Returns the last item enqueued.
' ------------------------------------------------------------------------------
    QueueLast UsedQueue(l_queue), l_item
End Sub

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Public Sub DeQueue(ByRef d_item As Variant, _
          Optional ByRef d_queue As Collection = Nothing)
' ----------------------------------------------------------------------------
' Returns the first item from the queue (d_queue), in case none is provided
' from the module's internal queue.
' ----------------------------------------------------------------------------
    QueueDequeue UsedQueue(d_queue), d_item
End Sub

Public Sub EnQueue(ByVal q_item As Variant, _
          Optional ByRef q_queue As Collection = Nothing)
' ----------------------------------------------------------------------------
' Adds the item (q-var) to the queue (q_queue), in case none is provided to
' the module's internal queue.
' ----------------------------------------------------------------------------
    QueueEnqueue UsedQueue(q_queue), q_item
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
'
' Basic service:
' - Displays a debugging option button when the Conditional Compile Argument
'   'Debugging = 1'
' - Displays an optional additional "About the error:" section when a string is
'   concatenated with the error message by two vertical bars (||)
' - Displays the error message by means of VBA.MsgBox when neither of the
'   following is installed
'
' Extendend service when other Common Components are installed and indicated via
' Conditional Compile Arguments:
' - Invokes mErH.ErrMsg when the Conditional Compile Argument ErHComp = 1
' - Invokes mMsg.ErrMsg when the Conditional Compile Argument MsgComp = 1 (and
'   the mErH module is not installed / MsgComp not set)
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to turn
'          them into negative and in the error message back into a positive
'          number.
' - ErrSrc To provide an unambiguous procedure name by prefixing is with the
'          module name.
'
' See: https://github.com/warbe-maker/Common-VBA-Error-Services
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Consider extra information is provided with the error description
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mQueue." & sProc
End Function

Public Function IsEmpty(Optional ByRef q_queue As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the queue (q_queue) is empty, in case none is provided
' the module's internal queue.
' ----------------------------------------------------------------------------
    IsEmpty = QueueIsEmpty(UsedQueue(q_queue))
End Function

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

Public Function IsQueued(ByVal i_item As Variant, _
                Optional ByRef i_pos As Long, _
                Optional ByRef i_queue As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    IsQueued = IsInQueue(UsedQueue(i_queue), i_item, i_pos)
End Function

Public Sub Item(ByVal i_pos As Long, _
                ByRef i_item As Variant, _
       Optional ByRef i_queue As Collection = Nothing)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    QueueItem UsedQueue(i_queue), i_pos, i_item
End Sub

Private Sub QueueDequeue(ByRef q_queue As Collection, _
                         ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the top item in the queue (q_item), i.e. the last one pushed on it,
' and removes it from the queue.
' ----------------------------------------------------------------------------
    Dim Pos As Long
    If Not QueueIsEmpty(q_queue) Then
        QueueFirst q_queue, q_item, Pos
        q_queue.Remove Pos
    End If
End Sub

Private Sub QueueEnqueue(ByRef q_queue As Collection, _
                         ByVal q_item As Variant)
    If q_queue Is Nothing Then Set q_queue = New Collection
    q_queue.Add q_item
End Sub

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

Private Function QueueIsEmpty(ByVal q_queue As Collection) As Boolean
    QueueIsEmpty = q_queue Is Nothing
    If Not QueueIsEmpty Then QueueIsEmpty = q_queue.Count = 0
End Function

Private Sub QueueItem(ByVal q_queue As Collection, _
                      ByVal q_pos As Long, _
             Optional ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (q_item) at the position (q_pos) in the queue (q_queue),
' provided the queue is not empty and the position is within the queue's size.
' ----------------------------------------------------------------------------
    
    If Not QueueIsEmpty(q_queue) Then
        If q_pos <= QueueSize(q_queue) Then
            If VarType(q_queue(q_pos)) = vbObject Then
                Set q_item = q_queue(q_pos)
            Else
                q_item = q_queue(q_pos)
            End If
        End If
    End If
    
End Sub

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

Private Function QueueSize(ByVal q_queue As Collection) As Long
    If Not QueueIsEmpty(q_queue) Then QueueSize = q_queue.Count
End Function

Public Function Size(Optional ByRef q_queue As Collection = Nothing) As Long
' ----------------------------------------------------------------------------
' Returns the size (number of items) in the queue (q_queue), in case none is
' provided those of the module's internal queue.
' ----------------------------------------------------------------------------
    Size = QueueSize(UsedQueue(q_queue))
End Function

Private Function UsedQueue(Optional ByRef u_queue As Collection = Nothing) As Collection
' ------------------------------------------------------------------------------
' Provides the stack the caller has provided (passed with the call) or when none
' had been provided, a default stack.
' ------------------------------------------------------------------------------
    Const PROC = "UsedQueue"
    
    On Error GoTo eh
    Select Case True
        Case Not u_queue Is Nothing And TypeName(u_queue) <> "Collection"
            Err.Raise AppErr(1), ErrSrc(PROC), "The provided queue (u_queue) is not a Collection!"
        Case Not u_queue Is Nothing And TypeName(u_queue) = "Collection"
            Set UsedQueue = u_queue
        Case u_queue Is Nothing
            Set UsedQueue = cllQueue
    End Select

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub Test_Private_Queue_Services()
' ----------------------------------------------------------------------------
' Self-test for the 'Private' Queue.... services
' ----------------------------------------------------------------------------
    Dim MyQueue As New Collection
    Dim Item    As Variant
    Dim Pos     As Long
                                Debug.Assert QueueIsEmpty(MyQueue)
    QueueEnqueue MyQueue, "A":  Debug.Assert Not QueueIsEmpty(MyQueue)
    QueueDequeue MyQueue, Item: Debug.Assert Item = "A"
                                Debug.Assert QueueIsEmpty(MyQueue)
    QueueEnqueue MyQueue, "A"
    QueueEnqueue MyQueue, "B"
    QueueEnqueue MyQueue, "C"
    QueueEnqueue MyQueue, "D"
                                Debug.Assert Not QueueIsEmpty(MyQueue)
                                Debug.Assert QueueSize(MyQueue) = 4
                                Debug.Assert IsInQueue(MyQueue, "B", Pos) = True
                                Debug.Assert Pos = 2
    QueueItem MyQueue, 2, Item: Debug.Assert Item = "B"
    QueueDequeue MyQueue, Item: Debug.Assert Item = "A"
    QueueDequeue MyQueue, Item: Debug.Assert Item = "B"
    QueueDequeue MyQueue, Item: Debug.Assert Item = "C"
    QueueDequeue MyQueue, Item: Debug.Assert Item = "D"
                                Debug.Assert QueueIsEmpty(MyQueue)
    Set MyQueue = Nothing
    
End Sub

