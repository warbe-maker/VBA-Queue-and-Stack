Attribute VB_Name = "mTest"
Option Explicit

Private MyStack  As New Collection
Private MyQueue  As New Collection

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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest." & sProc
End Function

Public Sub Regression_Test()
    Const PROC = "Regression_test"
    
    mTrc.LogClear
    BoP ErrSrc(PROC)
    mErH.Regression = True
    mTest.Test_StandardModule_mQue_Default_Queue
    mTest.Test_StandardModule_mQue_Provided_Queue
    mTest.Test_StandardModule_mStck_Default_Stack
    mTest.Test_StandardModule_mStck_Provided_Stack
    mTest.Test_ClassModule_clsStack
    mTest.Test_ClassModule_clsQueue
    EoP ErrSrc(PROC)
    mTrc.Dsply
    
End Sub

Public Sub Test_ClassModule_clsQueue()
    Const PROC = "Test_ClassModule_clsQueue"
    
    Dim MyQueue     As New clsQueue
    Dim QueueItem   As Variant
    Dim QueuePos    As Long
    
    BoP ErrSrc(PROC)
    
    BoC "clsQueue.IsEmpty"
    Debug.Assert MyQueue.IsEmpty
    EoC "clsQueue.IsEmpty"
    
    BoC "cllQueue.Enqueue"
    MyQueue.EnQueue "A"
    EoC "cllQueue.Enqueue"
    
    BoC "cllQueue.DeQueue"
    MyQueue.DeQueue QueueItem
    EoC "cllQueue.DeQueue"
    Debug.Assert QueueItem = "A"
    
    MyQueue.EnQueue "A"
    MyQueue.EnQueue "B"
    MyQueue.EnQueue "C"
    MyQueue.EnQueue "D"
    
    BoC "cllQueue.Size"
    Debug.Assert MyQueue.Size = 4
    EoC "cllQueue.Size"
    
    BoC "cllQueue.First"
    MyQueue.First QueueItem
    Debug.Assert QueueItem = "A"
    EoC "cllQueue.First"
        
    BoC "cllQueue.Last"
    MyQueue.Last QueueItem
    EoC "cllQueue.Last"
    Debug.Assert QueueItem = "D"
    
    BoC "cllQueue.IsQueued"
    Debug.Assert MyQueue.IsQueued("C", QueuePos) = True
    EoC "cllQueue.IsQueued"
    Debug.Assert QueuePos = 3
    
    Debug.Assert Not MyQueue.IsEmpty
    MyQueue.DeQueue QueueItem
    Debug.Assert QueueItem = "A"
    MyQueue.DeQueue QueueItem
    Debug.Assert QueueItem = "B"
    MyQueue.DeQueue QueueItem
    Debug.Assert QueueItem = "C"
    MyQueue.DeQueue QueueItem
    Debug.Assert QueueItem = "D"
    Debug.Assert MyQueue.IsEmpty
    
    Set MyQueue = Nothing
    
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_ClassModule_clsStack()
    Const PROC = "Test_ClassModule_clsStack"
    
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

Public Sub Test_StandardModule_mQue_Default_Queue()
    Const PROC = "Test_StandardModule_mQue_Default_Queue"
    
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

Public Sub Test_StandardModule_mQue_Provided_Queue()
    Const PROC = "Test_StandardModule_mQue_Provided_Queue"

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

Public Sub Test_StandardModule_mStck_Default_Stack()
    Const PROC = "Test_StandardModule_mStck_Default_Stack"

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
    Debug.Assert mStack.Stacked("B") = True
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

Public Sub Test_StandardModule_mStck_Provided_Stack()
    Const PROC = "Test_StandardModule_mStck_Provided_Stack"
    
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
    Debug.Assert mStack.Stacked("B", Pos, MyStack) = True
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

