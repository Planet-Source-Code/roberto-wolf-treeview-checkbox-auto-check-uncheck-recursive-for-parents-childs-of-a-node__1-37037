Attribute VB_Name = "basTvw"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias _
       "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias _
       "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
       ByVal dwNewLong As Long) As Long

Public Sub tvwInitialize(ByRef Control As MSComctlLib.Treeview)

    Const TVS_CHECKBOXES = &H100
    Const GWL_STYLE = (-16)
    
    Dim CurStyle As Long
    Dim Result As Long
    
    CurStyle = GetWindowLong(Control.hwnd, GWL_STYLE)
    Result = SetWindowLong(Control.hwnd, GWL_STYLE, _
             CurStyle Or TVS_CHECKBOXES)
    
End Sub

Public Sub tvwCheckBoxes(ByRef Control As MSComctlLib.Treeview, ByRef Node As MSComctlLib.Node)

    '-- childs
    tvwCheckBoxesAux1 Control, Node
    
    '-- parent
    tvwCheckBoxesAux2 Control, Node

End Sub

Private Sub tvwCheckBoxesAux1(ByRef Control As MSComctlLib.Treeview, ByRef Node As MSComctlLib.Node)

    Dim v As Boolean, i As Long, l As Long, f As Long, c As Long
    
    If Node.Children = 0 Then Exit Sub
    
    v = Node.Checked
    f = Node.Child.Index
    l = Node.Child.LastSibling.Index
    For i = f To l
        Control.Nodes(i).Checked = v
        tvwCheckBoxesAux1 Control, Control.Nodes(i)
    Next
    
End Sub

Private Sub tvwCheckBoxesAux2(ByRef Control As MSComctlLib.Treeview, ByRef Node As MSComctlLib.Node)

    Dim n As MSComctlLib.Node
    Dim v As Boolean, i As Long, l As Long, f As Long, c As Long
    
    If Node.Parent Is Nothing Then Exit Sub
    
    v = Node.Checked
    Set n = Node.Parent
    f = n.Child.Index
    l = n.Child.LastSibling.Index
    c = 0
    For i = f To l
        If Control.Nodes(i).Checked Then
            c = c + 1
        End If
    Next
    n.Checked = CBool(c)
    tvwCheckBoxesAux2 Control, n

End Sub


