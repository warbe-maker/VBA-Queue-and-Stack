Attribute VB_Name = "mStack"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mStck: Provides Stack services for a default stack or one
'                        declared and provided by the caller.
' Public services:
' - IsEmpty Returns TRUE when the provided stack is empty, when none is
'              provided the module's default stack
' - Pop     Pops the last pushed item from the provided stack and returns
'              it, if none is provided the module's
'              default stack
' - Push    Pushes any item (object or variable) on the provided stack,
'              if none is provided on the module's default stack
' - Size    Returns the provided stack's height, i.e. size, if none is
'              provided the module's default stack
'
' W. Rauschenberger Berlin Jan 2013
' ----------------------------------------------------------------------------
Private cllStack As New Collection

Private Function UsedStack(Optional ByRef u_stack As Collection = Nothing) As Collection
    Const PROC = "UsedStack"
    
    On Error GoTo eh
    Select Case True
        Case Not u_stack Is Nothing And TypeName(u_stack) <> "Collection"
            Err.Raise AppErr(1), ErrSrc(PROC), "The provided stack (u_stack) is not a Collection!"
        Case Not u_stack Is Nothing And TypeName(u_stack) = "Collection"
            Set UsedStack = u_stack
        Case u_stack Is Nothing
            Set UsedStack = cllStack
    End Select

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Stacked(ByVal s_item As Variant, _
               Optional ByRef s_pos As Long, _
               Optional ByRef s_stack As Collection = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the item (q_var) is stackd in (q_stack), when none is
' provided on the module's internal default stack.
' ------------------------------------------------------------------------------
    Stacked = IsOnStack(UsedStack(s_stack), s_item, s_pos)
End Function

Public Sub Push(ByVal p_item As Variant, _
       Optional ByRef p_stack As Collection = Nothing)
' ----------------------------------------------------------------------------
' Pushes the item (p_item) on the stack (p_stack), in case none is provided on
' the module's internal stack.
' ----------------------------------------------------------------------------
    UsedStack(p_stack).Add p_item
End Sub

Public Sub Top(ByRef t_item As Variant, _
               ByRef t_pos As Long, _
      Optional ByRef t_stack As Collection = Nothing)
' ----------------------------------------------------------------------------
' Returns the top item (t_item) on the stack (t-stack), when none is
' provided on the module's internal default stack.
' ----------------------------------------------------------------------------
    StackTop UsedStack(t_stack), t_item, t_pos
End Sub

Public Function IsEmpty(Optional ByRef s_stack As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the stack (s_stack) is empty, in case no stack is provided
' the module's internal stack.
' ----------------------------------------------------------------------------
    IsEmpty = StackIsEmpty(UsedStack(s_stack))
End Function

Public Sub Pop(ByRef p_item As Variant, _
      Optional ByRef p_stack As Collection = Nothing)
' ----------------------------------------------------------------------------
' Returns the top item from the stack (s_stack) and removes it, in case no stack
' is provided from the module's internal stack.
' ----------------------------------------------------------------------------
    StackPop UsedStack(p_stack), p_item
End Sub

Public Function Size(Optional ByRef s_stack As Collection = Nothing) As Long
' ----------------------------------------------------------------------------
' Returns the size (i.e. the number of items) on the stack (s_stack), in case
' no stack is provided of the module's internal stack.
' ----------------------------------------------------------------------------
    Size = StackSize(UsedStack(s_stack))
End Function

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
    ErrSrc = "mStack." & sProc
End Function

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

Private Sub StackBottom(ByVal s_stack As Collection, _
                        ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the bottom item (s_item) on the stack (s_stack), provided the stack
' is not empty.
' ----------------------------------------------------------------------------
    If Not StackIsEmpty(s_stack) Then
        If VarType(s_stack(s_stack.Count)) = vbObject Then
            Set s_item = s_stack(s_stack.Count)
        Else
            s_item = s_stack(s_stack.Count)
        End If
    End If
End Sub

Private Function IsOnStack(ByVal s_stack As Collection, _
                           ByVal s_item As Variant, _
                  Optional ByRef s_pos As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the index (s_pos) when the item (s_item) is found in the
' stack (s_stack).
' ----------------------------------------------------------------------------
    Dim v   As Variant
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

Private Sub StackItem(ByVal s_stack As Collection, _
                      ByVal s_pos As Long, _
             Optional ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (s_item) at the position (s_pos) on the stack (s_stack),
' provided the stack is not empty and the position is within the stack's size.
' ----------------------------------------------------------------------------
    
    If Not StackIsEmpty(s_stack) Then
        If StackSize(s_stack) <= s_pos Then
            If VarType(s_stack(s_pos)) = vbObject Then
                Set s_item = s_stack(s_pos)
            Else
                s_item = s_stack(s_pos)
            End If
        End If
    End If
    
End Sub

Private Function StackIsEmpty(ByVal s_stack As Collection) As Boolean
    StackIsEmpty = s_stack Is Nothing
    If Not StackIsEmpty Then StackIsEmpty = s_stack.Count = 0
End Function

Private Sub StackPop(ByRef s_stack As Collection, _
                     ByRef s_item As Variant)
' ----------------------------------------------------------------------------
' Returns the top item on the stack (s_item), i.e. the last one pushed on it,
' and removes it from the stack.
' ----------------------------------------------------------------------------
    Dim pos As Long
    If Not StackIsEmpty(s_stack) Then
        StackTop s_stack, s_item, pos
        s_stack.Remove pos
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


