Attribute VB_Name = "mQueue"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mQue: Generic (typeless) queue services.
'
' Public services:
' - Qdeq        Returns the first item from the queue, in case none is
'               provided from the module's internal default queue.
' - Qed         Returns TRUE when the provided item (en) is queued, in case
'               none is provided in the  module's internal default queue.

' - Qenq        Adds a provided item to the queue, in case none is provided to
'               the module's internal default queue.
' - Qget        Returns the queue, in case none is provided the module's
'               internal default queue.
' - QisEmpty    Returns TRUE when the queue is empty, in case none
'               is provided the module's internal default queue.
' - Qsize       Returns the size (number of items) in the queue, in case none
'               is provided those of the module's internal default queue.
'
' W. Rauschenberger Berlin Jan 2013
' ----------------------------------------------------------------------------
Private cllQueue As New Collection

Public Function Qed(Optional ByRef q_queue As Collection = Nothing, _
                    Optional ByVal q_var As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the item (q_var) is queued in (q_queue), when none is
' provided in the module's internal default queue.
' ------------------------------------------------------------------------------
    Const PROC = "Qed"
    
    On Error GoTo eh
    Dim v As Variant
    
    If q_queue Is Nothing Then
        Set q_queue = cllQueue
    ElseIf TypeName(q_queue) <> "Collection" Then
        Err.Raise AppErr(1), ErrSrc(PROC), "The provided queue is not a Collection!"
    End If
    
    If VarType(q_var) = vbObject Then
        For Each v In q_queue
            If v Is q_var Then
                Qed = True
                GoTo xt
            End If
        Next v
    Else
        For Each v In q_queue
            If v = q_var Then
                Qed = True
                GoTo xt
            End If
        Next v
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Qsize(Optional ByRef q_queue As Collection = Nothing) As Long
' ----------------------------------------------------------------------------
' Returns the size (number of items) in the queue (q_queue), in case none is
' provided those of the module's internal queue.
' ----------------------------------------------------------------------------
    Const PROC = "QisEmpty"
    
    On Error GoTo eh
    If q_queue Is Nothing Then
        Set q_queue = cllQueue
    ElseIf TypeName(q_queue) <> "Collection" Then
        Err.Raise AppErr(1), ErrSrc(PROC), "The provided queue is not a Collection!"
    End If
    Set q_queue = Qget(q_queue)
    If q_queue Is Nothing Then Qsize = 0 Else Qsize = q_queue.Count

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Property Let Qenq(Optional ByRef q_queue As Collection = Nothing, _
                                  ByVal q_var As Variant)
' ----------------------------------------------------------------------------
' Adds the item (q-var) to the queue (q_queue), in case none is provided to
' the module's internal queue.
' ----------------------------------------------------------------------------
    Const PROC = "Qenq-Let"
    
    On Error GoTo eh
    
    If q_queue Is Nothing Then
        Set q_queue = cllQueue
    ElseIf TypeName(q_queue) <> "Collection" Then
        Err.Raise AppErr(1), ErrSrc(PROC), "The provided queue is not a Collection!"
    End If
    q_queue.Add q_var

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function Qdeq(Optional ByRef q_queue As Collection = Nothing) As Variant
' ----------------------------------------------------------------------------
' Returns the first item from the queue (q_queue), , in case none is provided
' from the module's internal queue.
' ----------------------------------------------------------------------------
    Const PROC = "Qdeq-Get"
    
    On Error GoTo eh
    
    If q_queue Is Nothing Then
        Set q_queue = cllQueue
    ElseIf TypeName(q_queue) <> "Collection" Then
        Err.Raise AppErr(1), ErrSrc(PROC), "The provided queue is not a Collection!"
    End If
    If QisEmpty(q_queue) Then GoTo xt
    
    If VarType(q_queue(q_queue.Count)) = vbObject Then
        Set Qdeq = q_queue(q_queue.Count)
    Else
        Qdeq = q_queue(q_queue.Count)
    End If
    q_queue.Remove q_queue.Count

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function QisEmpty(Optional ByRef q_queue As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the queue (q_queue) is empty, in case none is provided
' the module's internal queue.
' ----------------------------------------------------------------------------
    Const PROC = "QisEmpty"
    
    On Error GoTo eh
    If q_queue Is Nothing Then
        Set q_queue = cllQueue
    ElseIf TypeName(q_queue) <> "Collection" Then
        Err.Raise AppErr(1), ErrSrc(PROC), "The provided queue is not a Collection!"
    End If
    
    Set q_queue = Qget(q_queue)
    QisEmpty = q_queue Is Nothing
    If Not QisEmpty Then QisEmpty = q_queue.Count = 0
    If QisEmpty Then Set q_queue = Nothing

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Qget(Optional ByVal q_queue As Collection = Nothing) As Collection
' ----------------------------------------------------------------------------
' Returns the queue (q_queue), in case none is provided the module's internal
' queue.
' ----------------------------------------------------------------------------
    Const PROC = "QisEmpty"
    
    On Error GoTo eh
    If q_queue Is Nothing Then
        Set q_queue = cllQueue
    ElseIf TypeName(q_queue) <> "Collection" Then
        Err.Raise AppErr(1), ErrSrc(PROC), "The provided queue is not a Collection!"
    End If
    Set Qget = q_queue
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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
    ErrSrc = "mQueue." & sProc
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



