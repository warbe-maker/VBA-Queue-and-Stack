Attribute VB_Name = "mStk"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mStk: Provides Stk services for a default stack or one
'                         declared and provided by the caller.
'
' Note: A stack can be seen as a tube which is closed at one end which means
'       that only the last item added (pushed into it, pushed on the stack
'       respecitvely) can be returned (popped). In this sense the bottom of a
'       stack is the item added first and the top is the item added last, i.e
'       the one popped next.
'
' Public services:
' - Bottom      Returns the botton item on the stack
' - IsEmpty
' - IsStked   Returns TRUE and the items position when a given item is on
'               the stack
' - Item        Returns an item on a provided position
' - Pop         Returns the top (last added) item from the stack
' - Push        Pushes an item on the stack
' - Size        Returns the current stack's size
' - Top         Returns the top item on the stack without taking it off the
'               stack, i.e. not popping it.
'
' W. Rauschenberger Berlin Jan 2013
' ----------------------------------------------------------------------------
Private cllStk As New Collection

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

Public Sub Bottom(ByRef b_item As Variant, _
         Optional ByRef b_stack As Collection = Nothing)
    StkBottom StkUsed(b_stack), b_item
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
    ErrSrc = "mStk." & sProc
End Function

Public Function IsEmpty(Optional ByRef s_stack As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the stack (s_stack) is empty, in case no stack is provided
' the module's internal stack.
' ----------------------------------------------------------------------------
    IsEmpty = StkIsEmpty(StkUsed(s_stack))
End Function

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

Public Sub Pop(ByRef p_item As Variant, _
      Optional ByRef p_stack As Collection = Nothing)
' ----------------------------------------------------------------------------
' Returns the top item from the stack (s_stack) and removes it, in case no stack
' is provided from the module's internal stack.
' ----------------------------------------------------------------------------
    StkPop StkUsed(p_stack), p_item
End Sub

Public Sub Push(ByVal p_item As Variant, _
       Optional ByRef p_stack As Collection = Nothing)
' ----------------------------------------------------------------------------
' Pushes the item (p_item) on the stack (p_stack), in case none is provided on
' the module's internal stack.
' ----------------------------------------------------------------------------
    StkUsed(p_stack).Add p_item
End Sub

Public Function Size(Optional ByRef s_stack As Collection = Nothing) As Long
' ----------------------------------------------------------------------------
' Returns the size (i.e. the number of items) on the stack (s_stack), in case
' no stack is provided of the module's internal stack.
' ----------------------------------------------------------------------------
    Size = StkSize(StkUsed(s_stack))
End Function

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

Public Sub Top(ByRef t_item As Variant, _
               ByRef t_pos As Long, _
      Optional ByRef t_stack As Collection = Nothing)
' ----------------------------------------------------------------------------
' Returns the top item (t_item) on the stack (t-stack), when none is
' provided on the module's internal default stack.
' ----------------------------------------------------------------------------
    StkTop StkUsed(t_stack), t_item, t_pos
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

