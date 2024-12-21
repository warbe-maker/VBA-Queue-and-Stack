Attribute VB_Name = "mTest"
Option Explicit

Private MyStk   As New Collection
Private cllStk  As New Collection
Private MyQ     As clsQ
Private TestAid As clsTestAid

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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest." & sProc
End Function

Public Sub Test_00_Regression_Test()
    Const PROC = "Test_00_Regression_test"
    
    Prepare
    mTrc.FileLocation = TestAid.TestFolder
    mTrc.FileBaseName = "RegressionTest.ExecTrace"
    mTrc.FileExtension = ".log"
    mTrc.Title = "Regression-Test mQ, mStk"
    mErH.Regression = True
    
    mBasic.BoP ErrSrc(PROC)
    mTest.Test_10_Queue
'    mTest.Test_10_Queue_VarTypes
'    mTest.Test_10_ClassModule_clsStk
    
    mBasic.EoP ErrSrc(PROC)
    TestAid.ResultSummaryLog
    
End Sub

Private Sub Prepare(Optional ByVal p_q_s As String = vbNullString)
    
    If MyQ Is Nothing Then Set MyQ = New clsQ
    If TestAid Is Nothing Then Set TestAid = New clsTestAid
    With TestAid
        If p_q_s <> vbNullString Then .TestedComp = p_q_s
        .ModeRegression = mErH.Regression
    End With
    
End Sub

Public Sub Test_10_Queue()
    Const PROC = "Test_10_Queue"
    
    Dim lPosition   As Long
    Dim lNumber     As Long
    Dim dt          As Date
    Dim v           As Variant
    
    dt = Now
    
    mBasic.BoP ErrSrc(PROC)
    Prepare "clsQ"
    
    With TestAid
        .TestedComp = "clsQ"
        .TestId = "10-1"
        .Title = "clsQ basics (Clear/IsEmpty, EnQueue, Size, First, Last, DeQueue)"
        
        ' ---------------------------------
        .Verification = "Queue is empty"
        .TestedProc = "IsEmpty"
        .TestedProcType = "Property"
        .ResultExpected = True
        MyQ.Clear
        .Result = MyQ.IsEmpty
        
        '~~ Prepare test queue
        MyQ.EnQueue "A"                 ' 1st: a string
        MyQ.EnQueue True                ' 2nd: a boolean
        MyQ.EnQueue True                ' 3nd: a boolean
        MyQ.EnQueue ThisWorkbook        ' 41h: an object
        MyQ.EnQueue dt                  ' 5th: a Date
        ' ---------------------------------
        .Verification = "Queue size"
        .TestedProc = "Size"
        .TestedProcType = "Property"
        .ResultExpected = 5
        .Result = MyQ.Size
        
        ' ---------------------------------
        .Verification = "First = ""A"""
        .TestedProc = "First"
        .TestedProcType = "Method"
        .ResultExpected = "A"
        .Result = MyQ.First
        
        ' ---------------------------------
        .Verification = "Last is a certain date"
        .TestedProc = "Last"
        .TestedProcType = "Method"
        .ResultExpected = dt
        .Result = MyQ.Last
        
        ' ---------------------------------
        .Verification = "Dequeue = ""A"""
        .TestedProc = "DeQueue"
        .TestedProcType = "Method"
        .ResultExpected = "A"
        .Result = MyQ.DeQueue
                
        ' =======================================
        '~~ Prepare test queue
        MyQ.Clear
        MyQ.EnQueue "A"                 ' 1st: a string
        MyQ.EnQueue True                ' 2nd: a boolean
        MyQ.EnQueue True                ' 3nd: a boolean
        MyQ.EnQueue ThisWorkbook        ' 41h: an object
        MyQ.EnQueue dt                  ' 5th: a Date

        .TestId = "10-2"
        .Title = "clsQ extras (Item, IsQueued)"
        ' ------------------------------------------------------
        .Verification = "Item is in queue at the given position"
        .TestedProc = "Item"
        .TestedProcType = "Method"
        .ResultExpected = ThisWorkbook
        .Result = MyQ.Item(4)
        
        ' ------------------------------------------------------
        .Verification = "Item is queued"
        .TestedProc = "IsQueued"
        .TestedProcType = "Method"
        .ResultExpected = True
        .Result = MyQ.IsQueued(ThisWorkbook, lPosition, lNumber)
    
        ' ------------------------------------------------------
        .Verification = "Item is queued once"
        .TestedProc = "IsQueued"
        .TestedProcType = "Method"
        .ResultExpected = 1
        MyQ.IsQueued ThisWorkbook, lPosition, lNumber
        .Result = lNumber
        
        ' ------------------------------------------------------
        .Verification = "Item is queued at position 4"
        .TestedProc = "IsQueued"
        .TestedProcType = "Method"
        .ResultExpected = 4
        MyQ.IsQueued ThisWorkbook, lPosition, lNumber
        .Result = lPosition
        
        ' ------------------------------------------------------
        .Verification = "Item is queued twice"
        .TestedProc = "IsQueued"
        .TestedProcType = "Method"
        .ResultExpected = 2
        MyQ.IsQueued True, lPosition, lNumber
        .Result = lNumber
    
        ' ------------------------------------------------------
        .Verification = "Item is queued twice, first at pos 2"
        .TestedProc = "IsQueued"
        .TestedProcType = "Method"
        .ResultExpected = 2
        MyQ.IsQueued True, lPosition, lNumber
        .Result = lPosition
        
        ' ------------------------------------------------------
        .Verification = "Item is not queued"
        .TestedProc = "IsQueued"
        .TestedProcType = "Method"
        .ResultExpected = False
        .Result = MyQ.IsQueued("X")
    End With
    
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_10_Queue_VarTypes()
    Const PROC = "Test_10_Queue_VarTypes"
    
    Dim cll     As clsQ
    Dim vItem   As Variant
    
    mBasic.BoP ErrSrc(PROC)
    
    Set cll = New clsQ
    Dim v2   As Variant: v2 = Empty
    Debug.Assert VarType(v2) = vbEmpty
    mBasic.BoC "Enqueue/Dequeue VarType " & Test_10_Queue_VarTypes_Test_Case(v2)
    cll.EnQueue v2
    cll.DeQueue vItem
    mBasic.EoC "Enqueue/Dequeue VarType " & Test_10_Queue_VarTypes_Test_Case(v2)
    Debug.Assert VarType(vItem) = vbEmpty
    Set cll = Nothing
      
    Set cll = New clsQ
    Dim v3   As Variant: v3 = Null
    Debug.Assert VarType(v3) = vbNull
    mBasic.BoC "Enqueue/Dequeue VarType " & Test_10_Queue_VarTypes_Test_Case(v3)
    cll.EnQueue v3
    cll.DeQueue vItem
    mBasic.EoC "Enqueue/Dequeue VarType " & Test_10_Queue_VarTypes_Test_Case(v3)
    Debug.Assert VarType(vItem) = vbNull
    Set cll = Nothing
    
    Set cll = New clsQ
    Dim v1   As Variant: v1 = vbNullString
    Debug.Assert VarType(v1) = vbString
    mBasic.BoC "Enqueue/Dequeue VarType " & Test_10_Queue_VarTypes_Test_Case(v1)
    cll.EnQueue v1
    cll.DeQueue vItem
    mBasic.EoC "Enqueue/Dequeue VarType " & Test_10_Queue_VarTypes_Test_Case(v1)
    Debug.Assert VarType(vItem) = vbString
    Debug.Assert vItem = vbNullString
    Set cll = Nothing
    
    Set cll = New clsQ
    Dim o   As Object
    Debug.Assert VarType(o) = vbObject
    mBasic.BoC "Enqueue/Dequeue VarType " & Test_10_Queue_VarTypes_Test_Case(o)
    cll.EnQueue o
    cll.DeQueue vItem
    mBasic.EoC "Enqueue/Dequeue VarType " & Test_10_Queue_VarTypes_Test_Case(o)
    Debug.Assert VarType(vItem) = vbObject
    Set cll = Nothing
        
    mBasic.EoP ErrSrc(PROC)
End Sub

Private Function Test_10_Queue_VarTypes_Test_Case(ByVal v As Variant) As String
    Test_10_Queue_VarTypes_Test_Case = VarTypeString(v)
End Function

Public Sub Test_10_ClassModule_clsStk()
    Const PROC = "Test_10_ClassModule_clsStk"
    
    Dim MyStk As New clsStk
    Dim Item    As Variant
    Dim Pos     As Long
    
    mBasic.BoP ErrSrc(PROC)
    
    mBasic.BoC "clsStk.IsEmpty"
    Debug.Assert MyStk.IsEmpty
    mBasic.EoC "clsStk.IsEmpty"
    
    mBasic.BoC "MyStk.Push"
    MyStk.Push "A"
    mBasic.EoC "MyStk.Push"
    
    mBasic.BoC "MyStk.Pop"
    MyStk.Pop Item
    mBasic.EoC "MyStk.Pop"
    Debug.Assert Item = "A"
    
    MyStk.Push "A"
    MyStk.Push "B"
    MyStk.Push "C"
    MyStk.Push "D"
    
    mBasic.BoC "MyStk.Size"
    Debug.Assert MyStk.Size = 4
    mBasic.EoC "MyStk.Size"
    
    mBasic.BoC "MyStk.Top"
    MyStk.Top Item
    Debug.Assert Item = "D"
    mBasic.EoC "MyStk.Top"
        
    mBasic.BoC "MyStk.Bottom"
    MyStk.Bottom Item
    Debug.Assert Item = "A"
    mBasic.EoC "MyStk.Bottom"
    
    mBasic.BoC "MyStk.IsStked"
    Debug.Assert MyStk.IsStked("C", Pos) = True
    mBasic.EoC "MyStk.IsStked"
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
    
    mBasic.EoP ErrSrc(PROC)
End Sub

Public Sub Test_20_Stack_Default_Stk()
    Const PROC = "Test_20_Stack_Default_Stk"

    Dim Item As Variant
    
    mBasic.BoP ErrSrc(PROC)
    
    mBasic.BoC "mStk.IsEmpty"
    Debug.Assert mStk.IsEmpty()
    mBasic.EoC "mStk.IsEmpty"
    
    mBasic.BoC "mStk.Push"
    mStk.Push "A"
    mBasic.EoC "mStk.Push"
    
    mBasic.BoC "mStk.Pop"
    mStk.Pop Item
    Debug.Assert Item = "A"
    mBasic.EoC "mStk.Pop"
    
    mStk.Push "A"
    mStk.Push "B"
    mStk.Push "C"
    mStk.Push "D"
       
    mBasic.BoC "mStk.Size"
    Debug.Assert mStk.Size() = 4
    mBasic.EoC "mStk.Size"
        
    mBasic.BoC "mStk.Stked"
    Debug.Assert mStk.StkEd("B") = True
    mBasic.EoC "mStk.Stked"
    
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
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_20_Stack_Provided_Stk()
    Const PROC = "Test_20_Stack_Provided_Stk"
    
    Dim Item   As Variant
    Dim Pos    As Long
    
    mBasic.BoP ErrSrc(PROC)
    
    mBasic.BoC "mStk.IsEmpty"
    Debug.Assert mStk.IsEmpty(MyStk)
    mBasic.EoC "mStk.IsEmpty"
    
    mBasic.BoC "mStk.Push"
    mStk.Push "A", MyStk
    mBasic.EoC "mStk.Push"
    
    mBasic.BoC "mStk.Pop"
    mStk.Pop Item, MyStk
    mBasic.EoC "mStk.Pop"
    Debug.Assert Item = "A"
    
    mStk.Push "A", MyStk
    mStk.Push "B", MyStk
    mStk.Push "C", MyStk
    mStk.Push "D", MyStk
    
    mBasic.BoC "mStk.Size"
    Debug.Assert mStk.Size(MyStk) = 4
    mBasic.EoC "mStk.Size"
        
    mBasic.BoC "mStk.Stked"
    Debug.Assert mStk.StkEd("B", Pos, MyStk) = True
    mBasic.EoC "mStk.Stked"
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
    mBasic.EoP ErrSrc(PROC)
    
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
    VarTypeString = mBasic.Align(VarType(v), enAlignRight, 3) & " " & dct(VarType(v))
'    For i = 0 To dct.Count - 1
'        Debug.Print mBasic.Align(dct.Keys()(i), 3, AlignRight) & " " & dct.Items()(i)
'    Next i

End Function

