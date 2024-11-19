Attribute VB_Name = "mQ"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mQ: FiFo (queue) services based on a Collection as queue.
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
Private Const GITHUB_REPO_URL = "https://github.com/warbe-maker/VBA-Queue-and-Stack"
Private cllQueue As New Collection

Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

'***App Window Constants***
Private Const WIN_NORMAL = 1         'Open Normal
Private Const WIN_MAX = 3            'Open Maximized
Private Const WIN_MIN = 2            'Open Minimized

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

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
    Qdequeue UsedQueue(d_queue), d_item
End Sub

Public Sub EnQueue(ByVal q_item As Variant, _
          Optional ByRef q_queue As Collection = Nothing)
' ----------------------------------------------------------------------------
' Adds the item (q-var) to the queue (q_queue), in case none is provided to
' the module's internal queue.
' ----------------------------------------------------------------------------
    Qenqueue UsedQueue(q_queue), q_item
End Sub

Private Function ShellRun(ByVal sr_string As String, _
                 Optional ByVal sr_show_how As Long = WIN_NORMAL) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, sr_string, vbNullString, vbNullString, sr_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sr_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

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
    If err_source = vbNullString Then err_source = Err.source
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
    ErrSrc = "mQ." & sProc
End Function

Public Sub First(ByRef f_item As Variant, _
        Optional ByRef f_queue As Collection = Nothing)
' ------------------------------------------------------------------------------
' Returns the first item in the queue without dequeuing it.
' ------------------------------------------------------------------------------
    Qfirst UsedQueue(f_queue), f_item
End Sub

Public Function IsEmpty(Optional ByRef q_queue As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the queue (q_queue) is empty, in case none is provided
' the module's internal queue.
' ----------------------------------------------------------------------------
    IsEmpty = QisEmpty(UsedQueue(q_queue))
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
    Qitem UsedQueue(i_queue), i_pos, i_item
End Sub

Public Sub Last(ByRef l_item As Variant, _
       Optional ByRef l_queue As Collection = Nothing)
' ------------------------------------------------------------------------------
' Returns the last item enqueued.
' ------------------------------------------------------------------------------
    Qlast UsedQueue(l_queue), l_item
End Sub

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

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    Const README_URL = "/blob/master/README.md"
    
    If r_bookmark = vbNullString _
    Then ShellRun GITHUB_REPO_URL & README_URL _
    Else ShellRun GITHUB_REPO_URL & README_URL & "#" & r_bookmark
        
End Sub

Public Function Size(Optional ByRef q_queue As Collection = Nothing) As Long
' ----------------------------------------------------------------------------
' Returns the size (number of items) in the queue (q_queue), in case none is
' provided those of the module's internal queue.
' ----------------------------------------------------------------------------
    Size = Qsize(UsedQueue(q_queue))
End Function

Private Sub Test_Private_Queue_Services()
' ----------------------------------------------------------------------------
' Self-test for the 'Private' Queue.... services
' ----------------------------------------------------------------------------
    Dim MyQueue As New Collection
    Dim Item    As Variant
    Dim Pos     As Long
                                Debug.Assert QisEmpty(MyQueue)
    Qenqueue MyQueue, "A":  Debug.Assert Not QisEmpty(MyQueue)
    Qdequeue MyQueue, Item: Debug.Assert Item = "A"
                                Debug.Assert QisEmpty(MyQueue)
    Qenqueue MyQueue, "A"
    Qenqueue MyQueue, "B"
    Qenqueue MyQueue, "C"
    Qenqueue MyQueue, "D"
                                Debug.Assert Not QisEmpty(MyQueue)
                                Debug.Assert Qsize(MyQueue) = 4
                                Debug.Assert IsInQueue(MyQueue, "B", Pos) = True
                                Debug.Assert Pos = 2
    Qitem MyQueue, 2, Item: Debug.Assert Item = "B"
    Qdequeue MyQueue, Item: Debug.Assert Item = "A"
    Qdequeue MyQueue, Item: Debug.Assert Item = "B"
    Qdequeue MyQueue, Item: Debug.Assert Item = "C"
    Qdequeue MyQueue, Item: Debug.Assert Item = "D"
                                Debug.Assert QisEmpty(MyQueue)
    Set MyQueue = Nothing
    
End Sub

Private Function UsedQueue(Optional ByRef u_queue As Collection = Nothing) As Collection
' ------------------------------------------------------------------------------
' Provides the queue the caller has provided (passed with the call) or when none
' had been provided, a default queue.
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

