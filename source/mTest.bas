Attribute VB_Name = "mTest"
Option Explicit

Private MyStk   As New Collection
Private cllStk  As New Collection
Private MyQueue As New Collection

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

Private Sub Qdequeue(ByRef q_queue As Collection, _
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
    Const PROC = "Qdequeue"
    
    On Error GoTo eh
    Dim lPos    As Long
    Dim lNo     As Long
    
    If Not QisEmpty(q_queue) Then
        If Not QisNothing(q_item_to_be_dequeued) Then
            If QisQueued(q_queue, q_item_to_be_dequeued, lPos, lNo) Then
                If lNo > 1 _
                Then Err.Raise AppErr(1), ErrSrc(PROC), "The specific item provided cannot be dequeued since it is not unambigous but in the queue " & lNo & " times!"
                QvarType q_item_to_be_dequeued, q_item_returned
                q_queue.Remove lPos
            End If
        ElseIf q_item_pos_to_be_dequeued <> 0 Then
            If q_item_pos_to_be_dequeued <= Qsize(q_queue) Then
                Qitem q_queue, q_item_pos_to_be_dequeued, q_item_returned
                q_queue.Remove q_item_pos_to_be_dequeued
            End If
        Else
            Qfirst q_queue, q_item_returned
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

Private Sub Qenqueue(ByRef q_queue As Collection, _
                         ByVal q_item As Variant)
    If q_queue Is Nothing Then Set q_queue = New Collection
    q_queue.Add q_item
End Sub

Private Sub Qfirst(ByVal q_queue As Collection, _
                       ByRef q_item_returned As Variant)
' ----------------------------------------------------------------------------
' Returns the current first item in the queue, i.e. the one added (enqueued)
' first. When the queue is empty Nothing is returned
' ----------------------------------------------------------------------------
    If Not QisEmpty(q_queue) Then
        QvarType q_queue(1), q_item_returned
    Else
        Set q_item_returned = Nothing
    End If

End Sub

Private Function QidenticalItems(ByVal q_1 As Variant, _
                                     ByVal q_2 As Variant) As Boolean
' ----------------------------------------------------------------------------
' Retunrs TRUE when item 1 is identical with item 2.
' ----------------------------------------------------------------------------
    Select Case True
        Case VarType(q_1) = vbObject And VarType(q_2) = vbObject:   QidenticalItems = q_1 Is q_2
        Case VarType(q_1) <> vbObject And VarType(q_2) <> vbObject: QidenticalItems = q_1 = q_2
    End Select
End Function

Private Function QisEmpty(ByVal q_queue As Collection) As Boolean
    QisEmpty = q_queue Is Nothing
    If Not QisEmpty Then QisEmpty = q_queue.Count = 0
End Function

Private Function QisNothing(ByVal i_item As Variant) As Boolean
    Select Case True
        Case VarType(i_item) = vbNull:      QisNothing = True
        Case VarType(i_item) = vbEmpty:     QisNothing = True
        Case VarType(i_item) = vbObject:    QisNothing = i_item Is Nothing
        Case IsNumeric(i_item):             QisNothing = CInt(i_item) = 0
        Case VarType(i_item) = vbString:    QisNothing = i_item = vbNullString
    End Select
End Function

Private Function QisQueued(ByVal i_queue As Collection, _
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
        If QidenticalItems(i_queue(i), i_item) Then
            i_no_found = i_no_found + 1
            QisQueued = True
            i_pos = i
        End If
    Next i

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

Private Sub QvarType(ByVal q_item As Variant, _
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

Private Sub StkBottom(ByVal s_stack As Collection, _
                        ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the bottom item (s_item) on the stack (s_stack), provided the stack
' is not empty.
' ----------------------------------------------------------------------------
    Dim lBottom As Long
    
    If Not StkIsEmpty(s_stack) Then
        lBottom = 1
        If VarType(s_stack(lBottom)) = vbObject Then
            Set s_item = s_stack(lBottom)
        Else
            s_item = s_stack(lBottom)
        End If
    End If
End Sub

Public Function StkEd(ByVal s_item As Variant, _
               Optional ByRef s_pos As Long, _
               Optional ByRef s_stack As Collection = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the item (q_var) is stackd in (q_stack), when none is
' provided on the module's internal default stack.
' ------------------------------------------------------------------------------
    StkEd = StkOn(StkUsed(s_stack), s_item, s_pos)
End Function

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
    Dim lTop As Long
    
    If Not StkIsEmpty(s_stack) Then
        lTop = s_stack.Count
        If VarType(s_stack(lTop)) = vbObject Then
            Set s_item = s_stack(lTop)
        Else
            s_item = s_stack(lTop)
        End If
        s_pos = lTop
    End If
End Sub

Private Function StkUsed(Optional ByRef u_stack As Collection = Nothing) As Collection
' ------------------------------------------------------------------------------
' Provides the stack the caller has provided (passed with the call) or when none
' had been provided, a default stack.
' ------------------------------------------------------------------------------
    Const PROC = "StkUsed"
    
    On Error GoTo eh
    Select Case True
        Case Not u_stack Is Nothing And TypeName(u_stack) <> "Collection"
            Err.Raise AppErr(1), ErrSrc(PROC), "The provided stack (u_stack) is not a Collection!"
        Case Not u_stack Is Nothing And TypeName(u_stack) = "Collection"
            Set StkUsed = u_stack
        Case u_stack Is Nothing
            Set StkUsed = cllStk
    End Select

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Test_00_Regression_Test()
    Const PROC = "Test_00_Regression_test"
    
    mTrc.LogClear
    BoP ErrSrc(PROC)
    
    mErH.Regression = True
    mTest.Test_10_ClassModule_clsQ
    mTest.Test_10_ClassModule_clsQ_VarTypes
    mTest.Test_10_ClassModule_clsStk
    mTest.Test_20_StandardModule_mQue_Default_Queue
    mTest.Test_20_StandardModule_mQue_Provided_Queue
    mTest.Test_30_Queue_as_Private_Services
    mTest.Test_50_StandardModule_mStck_Default_Stk
    mTest.Test_50_StandardModule_mStck_Provided_Stk
    mTest.Test_60_Stack_as_Private_Services
    
    EoP ErrSrc(PROC)
    mTrc.Dsply
    
End Sub

Public Sub Test_10_ClassModule_clsQ()
    Const PROC = "Test_10_ClassModule_clsQ"
    
    Dim MyQueue         As New clsQ
    Dim vDequeuedItem   As Variant
    Dim lQueuePos       As Long
    Dim lNo             As Long
    
    BoP ErrSrc(PROC)
    
    BoC "clsQ.IsEmpty"
    Debug.Assert MyQueue.IsEmpty
    EoC "clsQ.IsEmpty"
    
    BoC "clsQ.Enqueue"
    MyQueue.EnQueue "A":                Debug.Assert MyQueue.IsQueued("A", lQueuePos)
                                        Debug.Assert lQueuePos = 1
    EoC "clsQ.Enqueue"
    
    BoC "clsQ.DeQueue"
    MyQueue.DeQueue vDequeuedItem:      Debug.Assert vDequeuedItem = "A": Set vDequeuedItem = Nothing
    MyQueue.DeQueue vDequeuedItem:      Debug.Assert vDequeuedItem Is Nothing
    EoC "clsQ.DeQueue"
    
    MyQueue.EnQueue "A"                 ' 1st: a string
    MyQueue.EnQueue True                ' 2nd: a boolean
    MyQueue.EnQueue True                ' 3nd: a boolean
    MyQueue.EnQueue ThisWorkbook        ' 41h: an object
    MyQueue.EnQueue Now                 ' 5th: a Date
    
    BoC "clsQ.Size"
    Debug.Assert MyQueue.Size = 5
    EoC "clsQ.Size"
    
    BoC "clsQ.First"
    MyQueue.First vDequeuedItem:        Debug.Assert vDequeuedItem = "A"
    EoC "clsQ.First"
        
    BoC "clsQ.Last"
    MyQueue.Last vDequeuedItem:         Debug.Assert IsDate(vDequeuedItem)
    EoC "clsQ.Last"
    
    BoC "clsQ.Item"
    MyQueue.Item 2, vDequeuedItem:      Debug.Assert vDequeuedItem = True
    EoC "clsQ.Item"
    
    BoC "clsQ.IsQueued"
    Debug.Assert MyQueue.IsQueued(ThisWorkbook, lQueuePos, lNo) = True
    EoC "clsQ.IsQueued"
    Debug.Assert lQueuePos = 4
    Debug.Assert lNo = 1
      
    mErH.Asserted AppErr(1)
    BoC "clsQ.IsQueued AppErr(1) asserted"
    Debug.Assert MyQueue.IsQueued(True, lQueuePos, lNo) = True
    EoC "clsQ.IsQueued AppErr(1) asserted"
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

Public Sub Test_10_ClassModule_clsQ_VarTypes()
    Const PROC = "Test_10_ClassModule_clsQ_VarTypes"
    
    Dim cll     As clsQ
    Dim vItem   As Variant
    
    BoP ErrSrc(PROC)
    
    Set cll = New clsQ
    Dim v2   As Variant: v2 = Empty
    Debug.Assert VarType(v2) = vbEmpty
    BoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQ_VarTypes_Test_Case(v2)
    cll.EnQueue v2
    cll.DeQueue vItem
    EoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQ_VarTypes_Test_Case(v2)
    Debug.Assert VarType(vItem) = vbEmpty
    Set cll = Nothing
      
    Set cll = New clsQ
    Dim v3   As Variant: v3 = Null
    Debug.Assert VarType(v3) = vbNull
    BoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQ_VarTypes_Test_Case(v3)
    cll.EnQueue v3
    cll.DeQueue vItem
    EoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQ_VarTypes_Test_Case(v3)
    Debug.Assert VarType(vItem) = vbNull
    Set cll = Nothing
    
    Set cll = New clsQ
    Dim v1   As Variant: v1 = vbNullString
    Debug.Assert VarType(v1) = vbString
    BoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQ_VarTypes_Test_Case(v1)
    cll.EnQueue v1
    cll.DeQueue vItem
    EoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQ_VarTypes_Test_Case(v1)
    Debug.Assert VarType(vItem) = vbString
    Debug.Assert vItem = vbNullString
    Set cll = Nothing
    
    Set cll = New clsQ
    Dim o   As Object
    Debug.Assert VarType(o) = vbObject
    BoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQ_VarTypes_Test_Case(o)
    cll.EnQueue o
    cll.DeQueue vItem
    EoC "Enqueue/Dequeue VarType " & Test_10_ClassModule_clsQ_VarTypes_Test_Case(o)
    Debug.Assert VarType(vItem) = vbObject
    Set cll = Nothing
        
    EoP ErrSrc(PROC)
End Sub

Private Function Test_10_ClassModule_clsQ_VarTypes_Test_Case(ByVal v As Variant) As String
    Test_10_ClassModule_clsQ_VarTypes_Test_Case = VarTypeString(v)
End Function

Public Sub Test_10_ClassModule_clsStk()
    Const PROC = "Test_10_ClassModule_clsStk"
    
    Dim MyStk As New clsStk
    Dim Item    As Variant
    Dim Pos     As Long
    
    BoP ErrSrc(PROC)
    
    BoC "clsStk.IsEmpty"
    Debug.Assert MyStk.IsEmpty
    EoC "clsStk.IsEmpty"
    
    BoC "MyStk.Push"
    MyStk.Push "A"
    EoC "MyStk.Push"
    
    BoC "MyStk.Pop"
    MyStk.Pop Item
    EoC "MyStk.Pop"
    Debug.Assert Item = "A"
    
    MyStk.Push "A"
    MyStk.Push "B"
    MyStk.Push "C"
    MyStk.Push "D"
    
    BoC "MyStk.Size"
    Debug.Assert MyStk.Size = 4
    EoC "MyStk.Size"
    
    BoC "MyStk.Top"
    MyStk.Top Item
    Debug.Assert Item = "D"
    EoC "MyStk.Top"
        
    BoC "MyStk.Bottom"
    MyStk.Bottom Item
    Debug.Assert Item = "A"
    EoC "MyStk.Bottom"
    
    BoC "MyStk.IsStked"
    Debug.Assert MyStk.IsStked("C", Pos) = True
    EoC "MyStk.IsStked"
    Debug.Assert Pos = 3
    
    Debug.Assert Not MyStk.IsEmpty
    MyStk.Pop Item
    Debug.Assert Item = "D"
    MyStk.Pop Item
    Debug.Assert Item = "C"
    MyStk.Pop Item
    Debug.Assert Item = "B"
    MyStk.Pop Item
    Debug.Assert Item = "A"
    Debug.Assert MyStk.IsEmpty
    
    Set MyStk = Nothing
    
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_30_Queue_as_Private_Services()
' ------------------------------------------------------------------------------
' Tested are: All private services Queue.... copied from the Class Module
' clsQ - which are those also identical in the mQ Standatd Module.
' ------------------------------------------------------------------------------
    Const PROC = "Test_30_Queue_as_Private_Services"
    
    Dim MyQueue         As New Collection
    Dim vDequeuedItem   As Variant
    Dim lQueuePos       As Long
    Dim lNo             As Long
    
    BoP ErrSrc(PROC)
    
    BoC "QisEmpty"
    Debug.Assert QisEmpty(MyQueue)
    EoC "QisEmpty"
    
    BoC "Qenqueue"
    Qenqueue MyQueue, "A":                          Debug.Assert QisQueued(MyQueue, "A", lQueuePos)
                                                        Debug.Assert lQueuePos = 1
    EoC "Qenqueue"
    
    BoC "QdeQueue"
    Qdequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem = "A": Set vDequeuedItem = Nothing
    Qdequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem Is Nothing
    EoC "QdeQueue"
    
    Qenqueue MyQueue, "A"                           ' 1st: a string
    Qenqueue MyQueue, True                          ' 2nd: a boolean
    Qenqueue MyQueue, True                          ' 3nd: a boolean
    Qenqueue MyQueue, ThisWorkbook                  ' 41h: an object
    Qenqueue MyQueue, Now                           ' 5th: a Date
    
    BoC "Qsize"
    Debug.Assert Qsize(MyQueue) = 5
    EoC "Qsize"
    
    BoC "Qfirst"
    Qfirst MyQueue, vDequeuedItem:                  Debug.Assert vDequeuedItem = "A"
    EoC "Qfirst"
        
    BoC "Qlast"
    Qlast MyQueue, vDequeuedItem:                   Debug.Assert IsDate(vDequeuedItem)
    EoC "Qlast"
    
    BoC "Qitem"
    Qitem MyQueue, 2, vDequeuedItem:                Debug.Assert vDequeuedItem = True
    EoC "Qitem"
    
    BoC "QisQueued"
    Debug.Assert QisQueued(MyQueue, ThisWorkbook, lQueuePos, lNo) = True
    EoC "QisQueued"
    Debug.Assert lQueuePos = 4
    Debug.Assert lNo = 1
      
    mErH.Asserted AppErr(1)
    BoC "QisQueued AppErr(1) asserted"
    Debug.Assert QisQueued(MyQueue, True, lQueuePos, lNo) = True
    EoC "QisQueued AppErr(1) asserted"
    Debug.Assert lQueuePos = 3 ' the position of the last found item identical with True
    Debug.Assert lNo = 2
    
    Debug.Assert Not QisEmpty(MyQueue)
    Qdequeue MyQueue, vDequeuedItem, ThisWorkbook:  Debug.Assert vDequeuedItem Is ThisWorkbook
                                                        Debug.Assert QisQueued(MyQueue, ThisWorkbook) = False
    mErH.Asserted AppErr(1)
    Qdequeue MyQueue, vDequeuedItem, True
    
    Qdequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem = "A"
    Qdequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem = True
    Qdequeue MyQueue, vDequeuedItem:                Debug.Assert vDequeuedItem = True
    Qdequeue MyQueue, vDequeuedItem:                Debug.Assert IsDate(vDequeuedItem)
    Debug.Assert QisEmpty(MyQueue)
    
    Set MyQueue = Nothing
    
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_20_StandardModule_mQue_Default_Queue()
    Const PROC = "Test_20_StandardModule_mQue_Default_Queue"
    
    Dim Item As Variant
    Dim Pos  As Long
    
    BoP ErrSrc(PROC)
    BoC "mQ.IsEmpty"
    Debug.Assert mQ.IsEmpty()
    EoC "mQ.IsEmpty"
    
    BoC "mQ.EnQueue"
    mQ.EnQueue "A"
    EoC "mQ.EnQueue"
    
    BoC "mQ.DeQueue"
    mQ.DeQueue Item
    EoC "mQ.DeQueue"
    Debug.Assert Item = "A"
    
    mQ.EnQueue "A"
    mQ.EnQueue "B"
    mQ.EnQueue "C"
    mQ.EnQueue "D"
    
    BoC "mQ.Size"
    Debug.Assert mQ.Size() = 4
    EoC "mQ.Size"
        
    BoC "mQ.Queued"
    Debug.Assert mQ.IsQueued("B", Pos) = True
    EoC "mQ.Queued"
    Debug.Assert Pos = 2
    
    mQ.DeQueue Item:    Debug.Assert Item = "A"
    mQ.DeQueue Item:    Debug.Assert Item = "B"
    mQ.DeQueue Item:    Debug.Assert Item = "C"
    mQ.DeQueue Item:    Debug.Assert Item = "D"
                            Debug.Assert mQ.IsEmpty()
    EoP ErrSrc(PROC)
        
End Sub

Public Sub Test_20_StandardModule_mQue_Provided_Queue()
    Const PROC = "Test_20_StandardModule_mQue_Provided_Queue"

    Dim Item    As Variant
    Dim Pos     As Long

    BoP ErrSrc(PROC)
    BoC "mQ.IsEmpty"
    Debug.Assert mQ.IsEmpty(MyQueue)
    EoC "mQ.IsEmpty"
    
    BoC "mQ.EnQueue"
    mQ.EnQueue "A", MyQueue
    EoC "mQ.EnQueue"
    
    BoC "mQ.DeQueue"
    mQ.DeQueue Item, MyQueue:        Debug.Assert Item = "A"
    EoC "mQ.DeQueue"
    
    mQ.EnQueue "A", MyQueue
    mQ.EnQueue "B", MyQueue
    mQ.EnQueue "C", MyQueue
    mQ.EnQueue "D", MyQueue
    
    BoC "mQ.Size"
    Debug.Assert mQ.Size(MyQueue) = 4
    EoC "mQ.Size"
        
    BoC "mQ.IsQueued"
    Debug.Assert mQ.IsQueued("B", Pos, MyQueue) = True
    EoC "mQ.IsQueued"
    Debug.Assert Pos = 2
    
    Debug.Assert Not mQ.IsEmpty(MyQueue)
    mQ.DeQueue Item, MyQueue
    Debug.Assert Item = "A"
    mQ.DeQueue Item, MyQueue
    Debug.Assert Item = "B"
    mQ.DeQueue Item, MyQueue
    Debug.Assert Item = "C"
    mQ.DeQueue Item, MyQueue
    Debug.Assert Item = "D"
    
    Debug.Assert mQ.IsEmpty(MyQueue)
    
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_50_StandardModule_mStck_Default_Stk()
    Const PROC = "Test_50_StandardModule_mStck_Default_Stk"

    Dim Item As Variant
    
    BoP ErrSrc(PROC)
    
    BoC "mStk.IsEmpty"
    Debug.Assert mStk.IsEmpty()
    EoC "mStk.IsEmpty"
    
    BoC "mStk.Push"
    mStk.Push "A"
    EoC "mStk.Push"
    
    BoC "mStk.Pop"
    mStk.Pop Item
    Debug.Assert Item = "A"
    EoC "mStk.Pop"
    
    mStk.Push "A"
    mStk.Push "B"
    mStk.Push "C"
    mStk.Push "D"
       
    BoC "mStk.Size"
    Debug.Assert mStk.Size() = 4
    EoC "mStk.Size"
        
    BoC "mStk.Stked"
    Debug.Assert mStk.StkEd("B") = True
    EoC "mStk.Stked"
    
    Debug.Assert Not mStk.IsEmpty
    mStk.Pop Item
    Debug.Assert Item = "D"
    mStk.Pop Item
    Debug.Assert Item = "C"
    mStk.Pop Item
    Debug.Assert Item = "B"
    mStk.Pop Item
    Debug.Assert Item = "A"
    Debug.Assert mStk.IsEmpty
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_50_StandardModule_mStck_Provided_Stk()
    Const PROC = "Test_50_StandardModule_mStck_Provided_Stk"
    
    Dim Item   As Variant
    Dim Pos    As Long
    
    BoP ErrSrc(PROC)
    
    BoC "mStk.IsEmpty"
    Debug.Assert mStk.IsEmpty(MyStk)
    EoC "mStk.IsEmpty"
    
    BoC "mStk.Push"
    mStk.Push "A", MyStk
    EoC "mStk.Push"
    
    BoC "mStk.Pop"
    mStk.Pop Item, MyStk
    EoC "mStk.Pop"
    Debug.Assert Item = "A"
    
    mStk.Push "A", MyStk
    mStk.Push "B", MyStk
    mStk.Push "C", MyStk
    mStk.Push "D", MyStk
    
    BoC "mStk.Size"
    Debug.Assert mStk.Size(MyStk) = 4
    EoC "mStk.Size"
        
    BoC "mStk.Stked"
    Debug.Assert mStk.StkEd("B", Pos, MyStk) = True
    EoC "mStk.Stked"
    Debug.Assert Pos = 2
    
    Debug.Assert Not mStk.IsEmpty(MyStk)
    mStk.Pop Item, MyStk
    Debug.Assert Item = "D"
    mStk.Pop Item, MyStk
    Debug.Assert Item = "C"
    mStk.Pop Item, MyStk
    Debug.Assert Item = "B"
    mStk.Pop Item, MyStk
    Debug.Assert Item = "A"
    Debug.Assert mStk.IsEmpty(MyStk)
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_60_Stack_as_Private_Services()
' ----------------------------------------------------------------------------
' Self-test for the 'Private' Stk.... services
' ----------------------------------------------------------------------------
    Const PROC = "Test_50_Stack_as_Private_Services"
    
    Dim MyStk As New Collection
    Dim Item    As Variant
    Dim Pos     As Long
                                
    BoP ErrSrc(PROC)
                                Debug.Assert StkIsEmpty(MyStk)
    StkPush MyStk, "A":         Debug.Assert Not StkIsEmpty(MyStk)
    StkPop MyStk, Item:         Debug.Assert Item = "A"
                                Debug.Assert StkIsEmpty(MyStk)
    StkPush MyStk, "A"
    StkPush MyStk, "B"
    StkPush MyStk, "C"
    StkPush MyStk, "D"
                                Debug.Assert Not StkIsEmpty(MyStk)
                                Debug.Assert StkSize(MyStk) = 4
                                Debug.Assert StkOn(MyStk, "B", Pos) = True
                                Debug.Assert Pos = 2
    StkItem MyStk, 2, Item:     Debug.Assert Item = "B"
    StkPop MyStk, Item:         Debug.Assert Item = "D"
    StkPop MyStk, Item:         Debug.Assert Item = "C"
    StkPop MyStk, Item:         Debug.Assert Item = "B"
    StkPop MyStk, Item:         Debug.Assert Item = "A"
                                Debug.Assert StkIsEmpty(MyStk)
    Set MyStk = Nothing
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Top(ByRef t_item As Variant, _
               ByRef t_pos As Long, _
      Optional ByRef t_stack As Collection = Nothing)
' ----------------------------------------------------------------------------
' Returns the top item (t_item) on the stack (t-stack), when none is
' provided on the module's internal default stack.
' ----------------------------------------------------------------------------
    StkTop StkUsed(t_stack), t_item, t_pos
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

