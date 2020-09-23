VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Treeview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Treeview Check / UnCheck Sample"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView tvwGroups 
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8176
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "Treeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim n As Node
    Dim i, x, y As Integer
    
    Set n = tvwGroups.Nodes.Add(Key:="a", Text:="Root")
    For i = 1 To 4
        Set n = tvwGroups.Nodes.Add("a", tvwChild, Key:="a" & i, Text:="Item " & i)
        For x = 1 To 4
            Set n = tvwGroups.Nodes.Add("a" & i, tvwChild, "a" & i & x, "SubItem " & x)
            For y = 1 To 4
                Set n = tvwGroups.Nodes.Add("a" & i & x, tvwChild, "a" & i & x & y, "SubItem " & x)
            Next
        Next
    Next

    '-- Inicializa o Treeview, e remove um Bug do 'nodecheck'
    tvwInitialize tvwGroups

End Sub

Private Sub tvwGroups_NodeCheck(ByVal Node As MSComctlLib.Node)
    tvwCheckBoxes tvwGroups, Node
End Sub
