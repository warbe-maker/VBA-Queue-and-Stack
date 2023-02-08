VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------
' Class Module clsStk: LiFo (stack) services.
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
' W. Rauschenberger, Berlin Feb 2023
' ----------------------------------------------------------------------------
Private cllStk As Collection

Private Sub Class_Initialize():                     Set cllStk = New Collection:          End Sub

Private Sub Class_Terminate():                      Set cllStk = Nothing:                 End Sub

Public Sub Bottom(ByRef b_item As Variant):         StkBottom cllStk, b_item:           End Sub

Public Function IsEmpty() As Boolean:               IsEmpty = StkIsEmpty(cllStk):       End Function

Public Function IsStked(ByVal i_item As Variant, _
                 Optional ByRef i_pos As Long)
    IsStked = StkOn(cllStk, i_item, i_pos)
End Function

Public Sub Item(ByVal i_pos As Long, _
                ByRef i_item As Variant)
    StkItem cllStk, i_pos, i_item
End Sub
                
Public Sub Pop(ByRef p_item As Variant):            StkPop cllStk, p_item:              End Sub

Public Sub Push(ByVal p_item As Variant):           StkPush cllStk, p_item:             End Sub

Public Function Size() As Long:                     Size = StkSize(cllStk):             End Function

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

Private Function StkOn(ByVal s_stack As Collection, _
                           ByVal s_item As Variant, _
                  Optional ByRef s_pos As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the position (s_pos) when the item (s_item) is found on the
' stack (s_stack).
' ----------------------------------------------------------------------------
    Dim i As Long
    
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

Private Function StkIsEmpty(ByVal s_stack As Collection) As Boolean
    StkIsEmpty = s_stack Is Nothing
    If Not StkIsEmpty Then StkIsEmpty = s_stack.Count = 0
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

Public Sub Top(ByRef t_item):                       StkTop cllStk, t_item:              End Sub

