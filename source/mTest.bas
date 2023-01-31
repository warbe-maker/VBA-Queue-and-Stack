Attribute VB_Name = "mTest"
Option Explicit

Private MyStack  As New Collection
Private MyQueue  As New Collection

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

Public Sub Test_StandardModule_mQue_Default_Queue()
    Const PROC = "Test_StandardModule_mQue_Default_Queue"
    
    BoP ErrSrc(PROC)
    BoC "mQueue.QisEmpty"
    Debug.Assert mQueue.QisEmpty()
    EoC "mQueue.QisEmpty"
    
    BoC "mQueue.Qenq"
    mQueue.Qenq() = "A"
    EoC "mQueue.Qenq"
    
    BoC "mQueue.Qdeq"
    Debug.Assert mQueue.Qdeq() = "A"
    EoC "mQueue.Qdeq"
    
    mQueue.Qenq() = "A"
    mQueue.Qenq() = "B"
    mQueue.Qenq() = "C"
    mQueue.Qenq() = "D"
    
    BoC "mQueue.Qsize"
    Debug.Assert mQueue.Qsize() = 4
    EoC "mQueue.Qsize"
        
    BoC "mQueue.Qed"
    Debug.Assert mQueue.Qed(, "B") = True
    EoC "mQueue.Qed"
    
    Debug.Assert mQueue.Qdeq() = "D"
    Debug.Assert mQueue.Qdeq() = "C"
    Debug.Assert mQueue.Qdeq() = "B"
    Debug.Assert mQueue.Qdeq() = "A"
    Debug.Assert mQueue.QisEmpty()
    EoP ErrSrc(PROC)
        
End Sub

Public Sub Test_StandardModule_mQue_Provided_Queue()
    Const PROC = "Test_StandardModule_mQue_Provided_Queue"

    BoP ErrSrc(PROC)
    BoC "mQueue.QisEmpty"
    Debug.Assert mQueue.QisEmpty(MyQueue)
    EoC "mQueue.QisEmpty"
    
    BoC "mQueue.Qenq"
    mQueue.Qenq(MyQueue) = "A"
    EoC "mQueue.Qenq"
    
    BoC "mQueue.Qdeq"
    Debug.Assert mQueue.Qdeq(MyQueue) = "A"
    EoC "mQueue.Qdeq"
    
    mQueue.Qenq(MyQueue) = "A"
    mQueue.Qenq(MyQueue) = "B"
    mQueue.Qenq(MyQueue) = "C"
    mQueue.Qenq(MyQueue) = "D"
    
    BoC "mQueue.Qsize"
    Debug.Assert mQueue.Qsize(MyQueue) = 4
    EoC "mQueue.Qsize"
        
    BoC "mQueue.Qed"
    Debug.Assert mQueue.Qed(MyQueue, "B") = True
    EoC "mQueue.Qed"
    
    Debug.Assert Not mQueue.QisEmpty(MyQueue)
    Debug.Assert mQueue.Qdeq(MyQueue) = "D"
    Debug.Assert mQueue.Qdeq(MyQueue) = "C"
    Debug.Assert mQueue.Qdeq(MyQueue) = "B"
    Debug.Assert mQueue.Qdeq(MyQueue) = "A"
    Debug.Assert mQueue.QisEmpty(MyQueue)
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_StandardModule_mStck_Default_Stack()
    Const PROC = "Test_StandardModule_mStck_Default_Stack"

    Dim StackItem As Variant
    
    BoP ErrSrc(PROC)
    
    BoC "mStack.IsEmpty"
    Debug.Assert mStack.IsEmpty()
    EoC "mStack.IsEmpty"
    
    BoC "mStack.Push"
    mStack.Push "A"
    EoC "mStack.Push"
    
    BoC "mStack.Pop"
    mStack.Pop StackItem
    Debug.Assert StackItem = "A"
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
    mStack.Pop StackItem
    Debug.Assert StackItem = "D"
    mStack.Pop StackItem
    Debug.Assert StackItem = "C"
    mStack.Pop StackItem
    Debug.Assert StackItem = "B"
    mStack.Pop StackItem
    Debug.Assert StackItem = "A"
    Debug.Assert mStack.IsEmpty
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_StandardModule_mStck_Provided_Stack()
    Const PROC = "Test_StandardModule_mStck_Provided_Stack"
    
    Dim StackItem   As Variant
    Dim StackPos    As Long
    
    BoP ErrSrc(PROC)
    
    BoC "mStack.IsEmpty"
    Debug.Assert mStack.IsEmpty(MyStack)
    EoC "mStack.IsEmpty"
    
    BoC "mStack.Push"
    mStack.Push "A", MyStack
    EoC "mStack.Push"
    
    BoC "mStack.Pop"
    mStack.Pop StackItem, MyStack
    EoC "mStack.Pop"
    Debug.Assert StackItem = "A"
    
    mStack.Push "A", MyStack
    mStack.Push "B", MyStack
    mStack.Push "C", MyStack
    mStack.Push "D", MyStack
    
    BoC "mStack.Size"
    Debug.Assert mStack.Size(MyStack) = 4
    EoC "mStack.Size"
        
    BoC "mStack.Stacked"
    Debug.Assert mStack.Stacked("B", StackPos, MyStack) = True
    EoC "mStack.Stacked"
    Debug.Assert StackPos = 2
    
    Debug.Assert Not mStack.IsEmpty(MyStack)
    mStack.Pop StackItem, MyStack
    Debug.Assert StackItem = "D"
    mStack.Pop StackItem, MyStack
    Debug.Assert StackItem = "C"
    mStack.Pop StackItem, MyStack
    Debug.Assert StackItem = "B"
    mStack.Pop StackItem, MyStack
    Debug.Assert StackItem = "A"
    Debug.Assert mStack.IsEmpty(MyStack)
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_ClassModule_clsStack()
    Const PROC = "Test_ClassModule_clsStack"
    
    Dim MyStack     As New clsStack
    Dim StackItem   As Variant
    Dim StackPos    As Long
    
    BoP ErrSrc(PROC)
    
    BoC "clsStack.IsEmpty"
    Debug.Assert MyStack.IsEmpty
    EoC "clsStack.IsEmpty"
    
    BoC "cllStack.Push"
    MyStack.Push "A"
    EoC "cllStack.Push"
    
    BoC "cllStack.Pop"
    MyStack.Pop StackItem
    EoC "cllStack.Pop"
    Debug.Assert StackItem = "A"
    
    MyStack.Push "A"
    MyStack.Push "B"
    MyStack.Push "C"
    MyStack.Push "D"
    
    BoC "cllStack.Size"
    Debug.Assert MyStack.Size = 4
    EoC "cllStack.Size"
    
    BoC "cllStack.Top"
    MyStack.Top StackItem
    Debug.Assert StackItem = "D"
    EoC "cllStack.Top"
        
    BoC "cllStack.Bottom"
    MyStack.Bottom StackItem
    Debug.Assert StackItem = "A"
    EoC "cllStack.Bottom"
    
    BoC "cllStack.IsStacked"
    Debug.Assert MyStack.IsStacked("C", StackPos) = True
    EoC "cllStack.IsStacked"
    Debug.Assert StackPos = 3
    
    Debug.Assert Not MyStack.IsEmpty
    MyStack.Pop StackItem
    Debug.Assert StackItem = "D"
    MyStack.Pop StackItem
    Debug.Assert StackItem = "C"
    MyStack.Pop StackItem
    Debug.Assert StackItem = "B"
    MyStack.Pop StackItem
    Debug.Assert StackItem = "A"
    Debug.Assert MyStack.IsEmpty
    
    Set MyStack = Nothing
    
    EoP ErrSrc(PROC)
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
    
    BoC "cllQueue.Dequeue"
    MyQueue.DeQueue QueueItem
    EoC "cllQueue.Dequeue"
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


