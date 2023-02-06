Attribute VB_Name = "mTest"
Option Explicit

Private MyStack  As New Collection
Private MyQueue  As New Collection

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

Private Sub BoC(ByVal boc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(C)ode with id (boc_id) trace. Procedure to be copied as Private
' into any module potentially using the Common VBA Execution Trace Service. Has
' no effect when Conditional Compile Argument is 0 or not set at all.
' Note: The begin id (boc_id) has to be identical with the paired EoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.BoC boc_id, s
#End If
End Sub

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EoC(ByVal eoc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(C)ode id (eoc_id) trace. Procedure to be copied as Private into
' any module potentially using the Common VBA Execution Trace Service. Has no
' effect when the Conditional Compile Argument is 0 or not set at all.
' Note: The end id (eoc_id) has to be identical with the paired BoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.EoC eoc_id, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
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
    ErrSrc = "mTest." & sProc
End Function

Private Sub QueueDequeue(ByRef q_queue As Collection, _
                Optional ByRef q_item_returned As Variant, _
                Optional ByVal q_item_to_be_dequeued As Variant = Nothing, _
                Optional ByVal q_item_pos_to_be_dequeued As Long = 0)
' ----------------------------------------------------------------------------
' - When neither a specific item to be dequeued (q_item_to_be_dequeued) nor a
'   specific to be dequed item by its position (q_item_pos) is provided, the
'   service returns the top item in the queue (q_item_returned) - i.e. the
'   first one added, i.e. enqueued - and removes it from the queue.
' - When a specific item to be dequeued (q_item_to_be_dequeued) is provided
'   and it exists in the queue, this one is dequeued, returned
'   (q_item_returned) and removed from the queue.
' - When a specific to be dequeued item by its position (q_item_pos) is
'   provided - and the position is within the queue's size - this position's
'   item is returned and removed.
'
' Notes
' 1. When the argument (q_item_to_be_dequeued) is provided the argument
'    (q_item_pos_to_be_dequeued) is ignored.
' 2. All private procedures Queue... may be copied into any StandardModule
' ----------------------------------------------------------------------------
    Const PROC = "QueueDequeue"
    
    On Error GoTo eh
    Dim lPos    As Long
    Dim lNo     As Long
    
    If Not QueueIsEmpty(q_queue) Then
        If Not QueueIsNothing(q_item_to_be_dequeued) Then
            If QueueIsQueued(q_queue, q_item_to_be_dequeued, lPos, lNo) Then
                If lNo > 1 _
                Then Err.Raise AppErr(1), ErrSrc(PROC), "The specific item provided cannot be dequeued since it is not unambigous but in the queue " & lNo & " times!"
                QueueVarType q_item_to_be_dequeued, q_item_returned
                q_queue.Remove lPos
            End If
        ElseIf q_item_pos_to_be_dequeued <> 0 Then
            If q_item_pos_to_be_dequeued <= QueueSize(q_queue) Then
                QueueItem q_queue, q_item_pos_to_be_dequeued, q_item_returned
                q_queue.Remove q_item_pos_to_be_dequeued
            End If
        Else
            QueueFirst q_queue, q_item_returned
            q_queue.Remove 1
        End If
    Else
        Set q_item_returned = Nothing
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub QueueEnqueue(ByRef q_queue As Collection, _
                         ByVal q_item As Variant)
    If q_queue Is Nothing Then Set q_queue = New Collection
    q_queue.Add q_item
End Sub

Private Sub QueueFirst(ByVal q_queue As Collection, _
                       ByRef q_item_returned As Variant)
' ----------------------------------------------------------------------------
' Returns the current first item in the queue, i.e. the one added (enqueued)
' first. When the queue is empty Nothing is returned
' ----------------------------------------------------------------------------
    If Not QueueIsEmpty(q_queue) Then
        QueueVarType q_queue(1), q_item_returned
    Else
        Set q_item_returned = Nothing
    End If

End Sub

Private Function QueueIdenticalItems(ByVal q_1 As Variant, _
                                     ByVal q_2 As Variant) As Boolean
' ----------------------------------------------------------------------------
' Retunrs TRUE when item 1 is identical with item 2.
' ----------------------------------------------------------------------------
    Select Case True
        Case VarType(q_1) = vbObject And VarType(q_2) = vbObject:   QueueIdenticalItems = q_1 Is q_2
        Case VarType(q_1) <> vbObject And VarType(q_2) <> vbObject: QueueIdenticalItems = q_1 = q_2
    End Select
End Function

Private Function QueueIsEmpty(ByVal q_queue As Collection) As Boolean
    QueueIsEmpty = q_queue Is Nothing
    If Not QueueIsEmpty Then QueueIsEmpty = q_queue.Count = 0
End Function

Private Function QueueIsNothing(ByVal i_item As Variant) As Boolean
    Select Case True
        Case VarType(i_item) = vbNull:      QueueIsNothing = True
        Case VarType(i_item) = vbEmpty:     QueueIsNothing = True
        Case VarType(i_item) = vbObject:    QueueIsNothing = i_item Is Nothing
        Case IsNumeric(i_item):             QueueIsNothing = CInt(i_item) = 0
        Case VarType(i_item) = vbString:    QueueIsNothing = i_item = vbNullString
    End Select
End Function

Private Function QueueIsQueued(ByVal i_queue As Collection, _
                               ByVal i_item As Variant, _
                      Optional ByRef i_pos As Long, _
                      Optional ByRef i_no_found As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the index (i_pos) when the item (i_item) is found in the
' queue (i_queue).
' ----------------------------------------------------------------------------
    Dim i As Long
    
    i_no_found = 0
    For i = 1 To i_queue.Count
        If QueueIdenticalItems(i_queue(i), i_item) Then
            i_no_found = i_no_found + 1
            QueueIsQueued = True
            i_pos = i
        End If
    Next i

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

Private Sub QueueVarType(ByVal q_item As Variant, _
                         ByRef q_item_result As Variant)
' ----------------------------------------------------------------------------
' Returns the pr0vided item (q_item) with respect to its VarType (q_item_var).
' ----------------------------------------------------------------------------
    Set q_item_result = Nothing
    If VarType(q_item) = vbObject Then
        Set q_item_result = q_item
    Else
        q_item_result = q_item
    End If
End Sub

Private Function Test_10_ClassModule_clsQueue_VarTypes_Test_Case(ByVal v As Variant) As String
    Test_10_ClassModule_clsQueue_VarTypes_Test_Case = VarTypeString(v)
End Function

Public Sub Test_00_Regression_Test()
    Const PROC = "Test_00_Regression_test"
    
    mTrc.LogClear
    BoP ErrSrc(PROC)
    mErH.Regression = True
    mTest.Test_20_StandardModule_mQue_Default_Queue
    mTest.Test_20_StandardModule_mQue_Provided_Queue
    mTest.Test_20_StandardModule_mStck_Default_Stack
    mTest.Test_20_StandardModule_mStck_Provided_Stack
    mTest.Test_10_ClassModule_clsStack
    mTest.Test_10_ClassModule_clsQueue
    mTest.Test_10_ClassModule_clsQueue_VarTypes
    mTest.Test_10_Queue_as_Private_Services
    EoP ErrSrc(PROC)
    mTrc.Dsply
    
End Sub

Public Sub Test_10_ClassModule_clsQueue()
    Const PROC = "Test_10_ClassModule_clsQueue"
    
    Dim MyQueue         As New clsQueue
    Dim vDequeuedItem   As Variant
    Dim lQueuePos       As Long
    Dim lNo             As Long
    
    BoP ErrSrc(PROC)
    
    BoC "clsQueue.IsEmpty"
    Debug.Assert MyQueue.IsEmpty
    EoC "clsQueue.IsEmpty"
    
    BoC "clsQueue.Enqueue"
    MyQueue.EnQueue "A":                Debug.Assert MyQueue.IsQueued("A", lQueuePos)
                                        Debug.Assert lQueuePos = 1
    EoC "clsQueue.Enqueue"
    
    BoC "clsQueue.DeQueue"
    MyQueue.DeQueue vDequeuedItem:      Debug.Assert vDequeuedItem = "A": Set vDequeuedItem = Nothing
    MyQueue.DeQueue vDequeuedItem:      Debug.Assert vDequeuedItem Is Nothing
    EoC "clsQueue.DeQueue"
    
    MyQueue.EnQueue "A"                 ' 1st: a string
    MyQueue.EnQueue True                ' 2nd: a boolean
    MyQueue.EnQueue True                ' 3nd: a boolean
    MyQueue.EnQueue ThisWorkbook        ' 41h: an object
    MyQueue.EnQueue Now                 ' 5th: a Date
    
    BoC "clsQueue.Size"
    Debug.Assert MyQueue.Size = 5
    EoC "clsQueue.Size"
    
    BoC "clsQueue.First"
    MyQueue.First vDequeuedItem:        Debug.Assert vDequeuedItem = "A"
    EoC "clsQueue.First"
        
    BoC "clsQueue.Last"
    MyQueue.Last vDequeuedItem:         Debug.Assert IsDate(vDequeuedItem)
    EoC "clsQueue.Last"
    
    BoC "clsQueue.Item"
    MyQueue.Item 2, vDequeuedItem:      Debug.Assert vDequeuedItem = True
    EoC "clsQueue.Item"
    
    BoC "clsQueue.IsQueued"
    Debug.Assert MyQueue.IsQueued(ThisWorkbook, lQueuePos, lNo) = True
    EoC "clsQueue.IsQueued"
    Debug.Assert lQueuePos = 4
    Debug.Assert lNo = 1
      
    mErH.Asserted AppErr(1)
    BoC "clsQueue.IsQueued AppErr(1) asserted"
    Debug.Assert MyQueue.IsQueued(True, lQueuePos, lNo) = True
    EoC "clsQueue.IsQueued AppErr(1) asserted"
    Debug.Assert lQueuePos = 3 ' the position of the last found item identical with True
    Debug.Assert lNo = 2
    
    Debug.Assert Not MyQueue.IsEmpty
    MyQueue.DeQueue vDequeuedItem, ThisWorkbook:    Debug.Assert vDequeuedItem Is ThisWorkbook
                                                    Debug.Assert MyQueue.IsQueued(ThisWorkbook) = False
    mErH.Asserted AppErr(1)
    MyQueue.DeQueue vDequeuedItem, True
    
    MyQueue.DeQueue vDequeuedItem:      Debug.Assert vDequeuedItem = "A"
    MyQueue.DeQueue vDequeuedItem:      Debug.Assert vDequeuedItem = True
    MyQueue.DeQueue vDequeuedItem:      Debug.Assert vDequeuedItem = True
    MyQueue.DeQueue vDequeuedItem:      Debug.Assert IsDate(vDequeuedItem)
    Debug.Assert MyQueue.IsEmpty
    
    Set MyQueue = Nothing
    
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_10_ClassModule_clsQueue_VarTypes()
    Const PROC = "Test_10_ClassModule_clsQueue_VarTypes"
    
    Dim cll     As clsQueue
    Dim vItem   As Variant
    
    BoP ErrSrc(PROC)
    
    Set cll = New clsQueue
    Dim v2   As Variant: v2 = Empty
    Debug.Assert VarType(v2) = vbEmpty
    BoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQueue_VarTypes_Test_Case(v2)
    cll.EnQueue v2
    cll.DeQueue vItem
    EoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQueue_VarTypes_Test_Case(v2)
    Debug.Assert VarType(vItem) = vbEmpty
    Set cll = Nothing
      
    Set cll = New clsQueue
    Dim v3   As Variant: v3 = Null
    Debug.Assert VarType(v3) = vbNull
    BoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQueue_VarTypes_Test_Case(v3)
    cll.EnQueue v3
    cll.DeQueue vItem
    EoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQueue_VarTypes_Test_Case(v3)
    Debug.Assert VarType(vItem) = vbNull
    Set cll = Nothing
    
    Set cll = New clsQueue
    Dim v1   As Variant: v1 = vbNullString
    Debug.Assert VarType(v1) = vbString
    BoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQueue_VarTypes_Test_Case(v1)
    cll.EnQueue v1
    cll.DeQueue vItem
    EoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQueue_VarTypes_Test_Case(v1)
    Debug.Assert VarType(vItem) = vbString
    Debug.Assert vItem = vbNullString
    Set cll = Nothing
    
    Set cll = New clsQueue
    Dim o   As Object
    Debug.Assert VarType(o) = vbObject
    BoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQueue_VarTypes_Test_Case(o)
    cll.EnQueue o
    cll.DeQueue vItem
    EoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQueue_VarTypes_Test_Case(o)
    Debug.Assert VarType(vItem) = vbObject
    Set cll = Nothing
        
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_10_ClassModule_clsStack()
    Const PROC = "Test_10_ClassModule_clsStack"
    
    Dim MyStack As New clsStack
    Dim Item    As Variant
    Dim Pos     As Long
    
    BoP ErrSrc(PROC)
    
    BoC "clsStack.IsEmpty"
    Debug.Assert MyStack.IsEmpty
    EoC "clsStack.IsEmpty"
    
    BoC "cllStack.Push"
    MyStack.Push "A"
    EoC "cllStack.Push"
    
    BoC "cllStack.Pop"
    MyStack.Pop Item
    EoC "cllStack.Pop"
    Debug.Assert Item = "A"
    
    MyStack.Push "A"
    MyStack.Push "B"
    MyStack.Push "C"
    MyStack.Push "D"
    
    BoC "cllStack.Size"
    Debug.Assert MyStack.Size = 4
    EoC "cllStack.Size"
    
    BoC "cllStack.Top"
    MyStack.Top Item
    Debug.Assert Item = "D"
    EoC "cllStack.Top"
        
    BoC "cllStack.Bottom"
    MyStack.Bottom Item
    Debug.Assert Item = "A"
    EoC "cllStack.Bottom"
    
    BoC "cllStack.IsStacked"
    Debug.Assert MyStack.IsStacked("C", Pos) = True
    EoC "cllStack.IsStacked"
    Debug.Assert Pos = 3
    
    Debug.Assert Not MyStack.IsEmpty
    MyStack.Pop Item
    Debug.Assert Item = "D"
    MyStack.Pop Item
    Debug.Assert Item = "C"
    MyStack.Pop Item
    Debug.Assert Item = "B"
    MyStack.Pop Item
    Debug.Assert Item = "A"
    Debug.Assert MyStack.IsEmpty
    
    Set MyStack = Nothing
    
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_10_Queue_as_Private_Services()
' ------------------------------------------------------------------------------
' Tested are: All private services Queue.... copied from the Class Module
' clsQueue - which are those also identical in the mQueue Standatd Module.
' ------------------------------------------------------------------------------
    Const PROC = "Test_10_Queue_as_Private_Services"
    
    Dim MyQueue         As New Collection
    Dim vDequeuedItem   As Variant
    Dim lQueuePos       As Long
    Dim lNo             As Long
    
    BoP ErrSrc(PROC)
    
    BoC "QueueIsEmpty"
    Debug.Assert QueueIsEmpty(MyQueue)
    EoC "QueueIsEmpty"
    
    BoC "QueueEnqueue"
    QueueEnqueue MyQueue, "A":                          Debug.Assert QueueIsQueued(MyQueue, "A", lQueuePos)
                                                        Debug.Assert lQueuePos = 1
    EoC "QueueEnqueue"
    
    BoC "QueueDeQueue"
    QueueDequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem = "A": Set vDequeuedItem = Nothing
    QueueDequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem Is Nothing
    EoC "QueueDeQueue"
    
    QueueEnqueue MyQueue, "A"                           ' 1st: a string
    QueueEnqueue MyQueue, True                          ' 2nd: a boolean
    QueueEnqueue MyQueue, True                          ' 3nd: a boolean
    QueueEnqueue MyQueue, ThisWorkbook                  ' 41h: an object
    QueueEnqueue MyQueue, Now                           ' 5th: a Date
    
    BoC "QueueSize"
    Debug.Assert QueueSize(MyQueue) = 5
    EoC "QueueSize"
    
    BoC "QueueFirst"
    QueueFirst MyQueue, vDequeuedItem:                  Debug.Assert vDequeuedItem = "A"
    EoC "QueueFirst"
        
    BoC "QueueLast"
    QueueLast MyQueue, vDequeuedItem:                   Debug.Assert IsDate(vDequeuedItem)
    EoC "QueueLast"
    
    BoC "QueueItem"
    QueueItem MyQueue, 2, vDequeuedItem:                Debug.Assert vDequeuedItem = True
    EoC "QueueItem"
    
    BoC "QueueIsQueued"
    Debug.Assert QueueIsQueued(MyQueue, ThisWorkbook, lQueuePos, lNo) = True
    EoC "QueueIsQueued"
    Debug.Assert lQueuePos = 4
    Debug.Assert lNo = 1
      
    mErH.Asserted AppErr(1)
    BoC "QueueIsQueued AppErr(1) asserted"
    Debug.Assert QueueIsQueued(MyQueue, True, lQueuePos, lNo) = True
    EoC "QueueIsQueued AppErr(1) asserted"
    Debug.Assert lQueuePos = 3 ' the position of the last found item identical with True
    Debug.Assert lNo = 2
    
    Debug.Assert Not QueueIsEmpty(MyQueue)
    QueueDequeue MyQueue, vDequeuedItem, ThisWorkbook:  Debug.Assert vDequeuedItem Is ThisWorkbook
                                                        Debug.Assert QueueIsQueued(MyQueue, ThisWorkbook) = False
    mErH.Asserted AppErr(1)
    QueueDequeue MyQueue, vDequeuedItem, True
    
    QueueDequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem = "A"
    QueueDequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem = True
    QueueDequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem = True
    QueueDequeue MyQueue, vDequeuedItem:                Debug.Assert IsDate(vDequeuedItem)
    Debug.Assert QueueIsEmpty(MyQueue)
    
    Set MyQueue = Nothing
    
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_20_StandardModule_mQue_Default_Queue()
    Const PROC = "Test_20_StandardModule_mQue_Default_Queue"
    
    Dim Item As Variant
    Dim Pos  As Long
    
    BoP ErrSrc(PROC)
    BoC "mQueue.IsEmpty"
    Debug.Assert mQueue.IsEmpty()
    EoC "mQueue.IsEmpty"
    
    BoC "mQueue.EnQueue"
    mQueue.EnQueue "A"
    EoC "mQueue.EnQueue"
    
    BoC "mQueue.DeQueue"
    mQueue.DeQueue Item
    EoC "mQueue.DeQueue"
    Debug.Assert Item = "A"
    
    mQueue.EnQueue "A"
    mQueue.EnQueue "B"
    mQueue.EnQueue "C"
    mQueue.EnQueue "D"
    
    BoC "mQueue.Size"
    Debug.Assert mQueue.Size() = 4
    EoC "mQueue.Size"
        
    BoC "mQueue.Queued"
    Debug.Assert mQueue.IsQueued("B", Pos) = True
    EoC "mQueue.Queued"
    Debug.Assert Pos = 2
    
    mQueue.DeQueue Item:    Debug.Assert Item = "A"
    mQueue.DeQueue Item:    Debug.Assert Item = "B"
    mQueue.DeQueue Item:    Debug.Assert Item = "C"
    mQueue.DeQueue Item:    Debug.Assert Item = "D"
                            Debug.Assert mQueue.IsEmpty()
    EoP ErrSrc(PROC)
        
End Sub

Public Sub Test_20_StandardModule_mQue_Provided_Queue()
    Const PROC = "Test_20_StandardModule_mQue_Provided_Queue"

    Dim Item    As Variant
    Dim Pos     As Long

    BoP ErrSrc(PROC)
    BoC "mQueue.IsEmpty"
    Debug.Assert mQueue.IsEmpty(MyQueue)
    EoC "mQueue.IsEmpty"
    
    BoC "mQueue.EnQueue"
    mQueue.EnQueue "A", MyQueue
    EoC "mQueue.EnQueue"
    
    BoC "mQueue.DeQueue"
    mQueue.DeQueue Item, MyQueue:        Debug.Assert Item = "A"
    EoC "mQueue.DeQueue"
    
    mQueue.EnQueue "A", MyQueue
    mQueue.EnQueue "B", MyQueue
    mQueue.EnQueue "C", MyQueue
    mQueue.EnQueue "D", MyQueue
    
    BoC "mQueue.Size"
    Debug.Assert mQueue.Size(MyQueue) = 4
    EoC "mQueue.Size"
        
    BoC "mQueue.IsQueued"
    Debug.Assert mQueue.IsQueued("B", Pos, MyQueue) = True
    EoC "mQueue.IsQueued"
    Debug.Assert Pos = 2
    
    Debug.Assert Not mQueue.IsEmpty(MyQueue)
    mQueue.DeQueue Item, MyQueue
    Debug.Assert Item = "A"
    mQueue.DeQueue Item, MyQueue
    Debug.Assert Item = "B"
    mQueue.DeQueue Item, MyQueue
    Debug.Assert Item = "C"
    mQueue.DeQueue Item, MyQueue
    Debug.Assert Item = "D"
    
    Debug.Assert mQueue.IsEmpty(MyQueue)
    
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_20_StandardModule_mStck_Default_Stack()
    Const PROC = "Test_20_StandardModule_mStck_Default_Stack"

    Dim Item As Variant
    
    BoP ErrSrc(PROC)
    
    BoC "mStack.IsEmpty"
    Debug.Assert mStack.IsEmpty()
    EoC "mStack.IsEmpty"
    
    BoC "mStack.Push"
    mStack.Push "A"
    EoC "mStack.Push"
    
    BoC "mStack.Pop"
    mStack.Pop Item
    Debug.Assert Item = "A"
    EoC "mStack.Pop"
    
    mStack.Push "A"
    mStack.Push "B"
    mStack.Push "C"
    mStack.Push "D"
       
    BoC "mStack.Size"
    Debug.Assert mStack.Size() = 4
    EoC "mStack.Size"
        
    BoC "mStack.Stacked"
    Debug.Assert mStack.StackEd("B") = True
    EoC "mStack.Stacked"
    
    Debug.Assert Not mStack.IsEmpty
    mStack.Pop Item
    Debug.Assert Item = "D"
    mStack.Pop Item
    Debug.Assert Item = "C"
    mStack.Pop Item
    Debug.Assert Item = "B"
    mStack.Pop Item
    Debug.Assert Item = "A"
    Debug.Assert mStack.IsEmpty
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_20_StandardModule_mStck_Provided_Stack()
    Const PROC = "Test_20_StandardModule_mStck_Provided_Stack"
    
    Dim Item   As Variant
    Dim Pos    As Long
    
    BoP ErrSrc(PROC)
    
    BoC "mStack.IsEmpty"
    Debug.Assert mStack.IsEmpty(MyStack)
    EoC "mStack.IsEmpty"
    
    BoC "mStack.Push"
    mStack.Push "A", MyStack
    EoC "mStack.Push"
    
    BoC "mStack.Pop"
    mStack.Pop Item, MyStack
    EoC "mStack.Pop"
    Debug.Assert Item = "A"
    
    mStack.Push "A", MyStack
    mStack.Push "B", MyStack
    mStack.Push "C", MyStack
    mStack.Push "D", MyStack
    
    BoC "mStack.Size"
    Debug.Assert mStack.Size(MyStack) = 4
    EoC "mStack.Size"
        
    BoC "mStack.Stacked"
    Debug.Assert mStack.StackEd("B", Pos, MyStack) = True
    EoC "mStack.Stacked"
    Debug.Assert Pos = 2
    
    Debug.Assert Not mStack.IsEmpty(MyStack)
    mStack.Pop Item, MyStack
    Debug.Assert Item = "D"
    mStack.Pop Item, MyStack
    Debug.Assert Item = "C"
    mStack.Pop Item, MyStack
    Debug.Assert Item = "B"
    mStack.Pop Item, MyStack
    Debug.Assert Item = "A"
    Debug.Assert mStack.IsEmpty(MyStack)
    EoP ErrSrc(PROC)
    
End Sub

Private Function VarTypeString(ByVal v As Variant) As String
    Static dct  As Dictionary
    
    If dct Is Nothing Then
        Set dct = New Dictionary
        dct.Add vbEmpty, "vbEmpty           Empty (not initialized)"
        dct.Add vbNull, "vbNull            Null (no valid data)"
        dct.Add vbInteger, "vbInteger         Integer"
        dct.Add vbLong, "vbLong            Long integer"
        dct.Add vbSingle, "vbSingle          Single-precision floating-point number"
        dct.Add vbDouble, "vbDouble          Double-precision floating-point number"
        dct.Add vbCurrency, "vbCurrency        Currency value"
        dct.Add vbDate, "vbDate            Date value"
        dct.Add vbString, "vbString          String"
        dct.Add vbObject, "vbObject          Object"
        dct.Add vbError, "vbError           Error value"
        dct.Add vbBoolean, "vbBoolean         Boolean value"
        dct.Add vbVariant, "vbVariant         Variant (used only with arrays of variants)"
        dct.Add vbDataObject, "vbDataObject      A data access object"
        dct.Add vbDecimal, "vbDecimal         Decimal value"
        dct.Add vbByte, "vbByte            Byte value"
        dct.Add vbLongLong, "vbLongLong        LongLong integer (valid on 64-bit platforms only)"
        dct.Add vbUserDefinedType, "vbUserDefinedType Variants that contain user-defined types"
        dct.Add vbArray, "vbArray           Array (always added to another constant when returned by this function)"
    End If
    VarTypeString = mBasic.Align(VarType(v), 3, AlignRight) & " " & dct(VarType(v))
'    For i = 0 To dct.Count - 1
'        Debug.Print mBasic.Align(dct.Keys()(i), 3, AlignRight) & " " & dct.Items()(i)
'    Next i

End Function

