VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucTreeVew 1.2 - Test"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   StartUpPosition =   2  'CenterScreen
   Begin Test.ucTreeView ucTreeView1 
      Height          =   7725
      Left            =   150
      TabIndex        =   52
      Top             =   135
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   13626
   End
   Begin VB.CheckBox chkEnsureVisible 
      Appearance      =   0  'Flat
      Caption         =   "Ensure visible"
      Height          =   285
      Left            =   4260
      TabIndex        =   24
      Top             =   3930
      Width           =   1425
   End
   Begin VB.CommandButton cmdFirstVisible 
      Caption         =   "First visible"
      Height          =   375
      Left            =   7080
      TabIndex        =   28
      Top             =   3960
      Width           =   1350
   End
   Begin VB.CommandButton cmdLastVisible 
      Caption         =   "Last visible"
      Height          =   375
      Left            =   7080
      TabIndex        =   31
      Top             =   5220
      Width           =   1350
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8805
      MaxLength       =   50
      TabIndex        =   44
      Top             =   6510
      Width           =   1035
   End
   Begin VB.TextBox txthRelative 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6870
      MaxLength       =   50
      TabIndex        =   42
      Text            =   "0"
      Top             =   6510
      Width           =   1035
   End
   Begin VB.OptionButton optRelation 
      Appearance      =   0  'Flat
      Caption         =   "Previous"
      Height          =   285
      Index           =   4
      Left            =   8970
      TabIndex        =   49
      Top             =   7095
      Width           =   975
   End
   Begin VB.OptionButton optRelation 
      Appearance      =   0  'Flat
      Caption         =   "Next"
      Height          =   285
      Index           =   3
      Left            =   8250
      TabIndex        =   48
      Top             =   7095
      Width           =   975
   End
   Begin VB.OptionButton optRelation 
      Appearance      =   0  'Flat
      Caption         =   "Sorted"
      Height          =   285
      Index           =   2
      Left            =   7395
      TabIndex        =   3
      Top             =   7095
      Width           =   975
   End
   Begin VB.OptionButton optRelation 
      Appearance      =   0  'Flat
      Caption         =   "First"
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   47
      Top             =   7095
      Width           =   975
   End
   Begin VB.OptionButton optRelation 
      Appearance      =   0  'Flat
      Caption         =   "Last"
      Height          =   285
      Index           =   0
      Left            =   6060
      TabIndex        =   46
      Top             =   7095
      Value           =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   7230
      TabIndex        =   50
      Top             =   7485
      Width           =   1350
   End
   Begin VB.CommandButton cmdDeleteCurrent 
      Caption         =   "Delete current"
      Height          =   375
      Left            =   4260
      TabIndex        =   39
      Top             =   7485
      Width           =   1350
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4740
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3135
      Width           =   2235
   End
   Begin VB.CheckBox chkLabelEdit 
      Appearance      =   0  'Flat
      Caption         =   "LabelEdit"
      Height          =   285
      Left            =   7035
      TabIndex        =   6
      Top             =   450
      Width           =   2820
   End
   Begin VB.CheckBox chkFullRowSelect 
      Appearance      =   0  'Flat
      Caption         =   "FullRowSelect"
      Height          =   285
      Left            =   7035
      TabIndex        =   10
      Top             =   1590
      Width           =   2820
   End
   Begin VB.CheckBox chkHideSelection 
      Appearance      =   0  'Flat
      Caption         =   "HideSelection"
      Height          =   285
      Left            =   7035
      TabIndex        =   7
      Top             =   735
      Width           =   2820
   End
   Begin VB.CheckBox chkSingleExpand 
      Appearance      =   0  'Flat
      Caption         =   "SingleExpand (+[Ctl]: no collapse)"
      Height          =   285
      Left            =   7035
      TabIndex        =   8
      Top             =   1020
      Width           =   2820
   End
   Begin VB.CheckBox chkTrackSelect 
      Appearance      =   0  'Flat
      Caption         =   "TrackSelect"
      Height          =   285
      Left            =   7035
      TabIndex        =   9
      Top             =   1305
      Width           =   2820
   End
   Begin VB.CheckBox chkHilited 
      Appearance      =   0  'Flat
      Caption         =   "Hilited"
      Height          =   285
      Left            =   9150
      TabIndex        =   22
      Top             =   3135
      Width           =   990
   End
   Begin VB.CheckBox chkGhosted 
      Appearance      =   0  'Flat
      Caption         =   "Ghosted"
      Height          =   285
      Left            =   8115
      TabIndex        =   21
      Top             =   3135
      Width           =   990
   End
   Begin VB.CheckBox chkBold 
      Appearance      =   0  'Flat
      Caption         =   "Bold"
      Height          =   285
      Left            =   7380
      TabIndex        =   20
      Top             =   3135
      Width           =   960
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4260
      TabIndex        =   38
      Top             =   6615
      Width           =   1350
   End
   Begin VB.CheckBox chkHasRootLines 
      Appearance      =   0  'Flat
      Caption         =   "HasRootLines"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4260
      TabIndex        =   2
      Top             =   735
      Width           =   2820
   End
   Begin VB.CheckBox chkCheckBoxes 
      Appearance      =   0  'Flat
      Caption         =   "CheckBoxes*"
      Height          =   285
      Left            =   4260
      TabIndex        =   5
      Top             =   1305
      Width           =   2820
   End
   Begin VB.CheckBox chkHasLines 
      Appearance      =   0  'Flat
      Caption         =   "HasLines"
      Height          =   285
      Left            =   4260
      TabIndex        =   1
      Top             =   450
      Width           =   2820
   End
   Begin VB.CheckBox chkHasButtons 
      Appearance      =   0  'Flat
      Caption         =   "HasButtons (plus/minus buttons)"
      Height          =   285
      Left            =   4260
      TabIndex        =   4
      Top             =   1020
      Width           =   2820
   End
   Begin VB.CommandButton cmdLastSibling 
      Caption         =   "Last sibling"
      Height          =   375
      Left            =   8490
      TabIndex        =   35
      Top             =   5220
      Width           =   1350
   End
   Begin VB.CommandButton cmdNextSibling 
      Caption         =   "Next sibling"
      Height          =   375
      Left            =   8490
      TabIndex        =   34
      Top             =   4800
      Width           =   1350
   End
   Begin VB.CommandButton cmdPreviousSibling 
      Caption         =   "Previous sibling"
      Height          =   375
      Left            =   8490
      TabIndex        =   33
      Top             =   4380
      Width           =   1350
   End
   Begin VB.CommandButton cmdFirstSibling 
      Caption         =   "First sibling"
      Height          =   375
      Left            =   8490
      TabIndex        =   32
      Top             =   3960
      Width           =   1350
   End
   Begin VB.CommandButton cmdRoot 
      Caption         =   "Root"
      Height          =   375
      Left            =   4260
      TabIndex        =   25
      Top             =   4380
      Width           =   1350
   End
   Begin VB.CommandButton cmdChild 
      Caption         =   "Child"
      Height          =   375
      Left            =   5670
      TabIndex        =   27
      Top             =   4800
      Width           =   1350
   End
   Begin VB.CommandButton cmdParent 
      Caption         =   "Parent"
      Height          =   375
      Left            =   5670
      TabIndex        =   26
      Top             =   4380
      Width           =   1350
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Height          =   375
      Left            =   7080
      TabIndex        =   29
      Top             =   4380
      Width           =   1350
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   7080
      TabIndex        =   30
      Top             =   4800
      Width           =   1350
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill/Add"
      Height          =   375
      Left            =   4260
      TabIndex        =   37
      Top             =   6180
      Width           =   1350
   End
   Begin VB.Label lblNote 
      Caption         =   "*TreeView should be created again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   240
      Left            =   4260
      TabIndex        =   51
      Top             =   1650
      Width           =   2625
   End
   Begin VB.Label Label1 
      Caption         =   "Key:"
      Height          =   240
      Left            =   8430
      TabIndex        =   43
      Top             =   6540
      Width           =   360
   End
   Begin VB.Label lblhRelative 
      Caption         =   "hRelative:"
      Height          =   240
      Left            =   6090
      TabIndex        =   41
      Top             =   6540
      Width           =   765
   End
   Begin VB.Label lblRelation 
      Caption         =   "Relation:"
      Height          =   240
      Left            =   6090
      TabIndex        =   45
      Top             =   6870
      Width           =   765
   End
   Begin VB.Label lblInsertNode 
      BackColor       =   &H80000010&
      Caption         =   " Insert node"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   6075
      TabIndex        =   40
      Top             =   6180
      Width           =   3765
   End
   Begin VB.Label lblAddingDeletingNodes 
      BackColor       =   &H80000010&
      Caption         =   " Adding/deleting nodes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   4260
      TabIndex        =   36
      Top             =   5730
      Width           =   5580
   End
   Begin VB.Label lblNodeNavigation 
      BackColor       =   &H80000010&
      Caption         =   " Node navigation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   4260
      TabIndex        =   23
      Top             =   3600
      Width           =   5580
   End
   Begin VB.Label lblText 
      Caption         =   "Text:"
      Height          =   255
      Left            =   4245
      TabIndex        =   18
      Top             =   3165
      Width           =   690
   End
   Begin VB.Label lblTreeViewStyles 
      BackColor       =   &H80000010&
      Caption         =   " TreeView styles:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   4260
      TabIndex        =   0
      Top             =   135
      Width           =   5580
   End
   Begin VB.Label lblKeyVal 
      Height          =   255
      Left            =   5025
      TabIndex        =   15
      Top             =   2565
      Width           =   1455
   End
   Begin VB.Label lblFullPathVal 
      Height          =   225
      Left            =   5025
      TabIndex        =   17
      Top             =   2820
      Width           =   4815
   End
   Begin VB.Label lblhNodeVal 
      Caption         =   "0"
      Height          =   255
      Left            =   5025
      TabIndex        =   13
      Top             =   2310
      Width           =   1455
   End
   Begin VB.Label lblKey 
      Caption         =   "Key:"
      Height          =   255
      Left            =   4260
      TabIndex        =   14
      Top             =   2565
      Width           =   690
   End
   Begin VB.Label lblFullPath 
      Caption         =   "Full path:"
      Height          =   255
      Left            =   4260
      TabIndex        =   16
      Top             =   2820
      Width           =   690
   End
   Begin VB.Label lblhNode 
      Caption         =   "hNode:"
      Height          =   255
      Left            =   4260
      TabIndex        =   12
      Top             =   2310
      Width           =   690
   End
   Begin VB.Label lblCurrentNode 
      BackColor       =   &H80000010&
      Caption         =   " Current node:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   4260
      TabIndex        =   11
      Top             =   1980
      Width           =   5580
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_STYLE As Long = (-16)
Private Const BS_FLAT   As Long = &H8000&
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long





Private Sub Form_Load()

  Dim oCtl As Control

    With ucTreeView1
        
        Call .Initialize
        Call .InitializeImageList 'Default size is 16x16
     
        Call .AddIcon(LoadResPicture(101, vbResIcon))
        Call .AddIcon(LoadResPicture(102, vbResIcon))
        Call .AddIcon(LoadResPicture(103, vbResIcon))
        Call .AddIcon(LoadResPicture(104, vbResIcon))
        Call .AddIcon(LoadResPicture(105, vbResIcon))
        
        .ItemHeight = 18 'Should be even (Set to -1 to restore to default).
        .ItemIndent = 33 '19 is default.
    End With
    
    For Each oCtl In fTest.Controls
        If (TypeOf oCtl Is CommandButton) Then
            Call pvFlattenButton(oCtl.hWnd)
        End If
    Next oCtl
End Sub



'==== TreeView Styles

Private Sub chkHasLines_Click()
    ucTreeView1.HasLines = -chkHasLines
    chkHasRootLines.Enabled = chkHasButtons Or chkHasLines
    chkFullRowSelect.Enabled = Not -chkHasLines
End Sub

Private Sub chkHasRootLines_Click()
    ucTreeView1.HasRootLines = -chkHasRootLines
End Sub

Private Sub chkHasButtons_Click()
    ucTreeView1.HasButtons = -chkHasButtons
    chkHasRootLines.Enabled = chkHasButtons Or chkHasLines
End Sub

Private Sub chkCheckBoxes_Click()
    ucTreeView1.CheckBoxes = -chkCheckBoxes
End Sub

Private Sub chkLabelEdit_Click()
    ucTreeView1.LabelEdit = -chkLabelEdit
End Sub

Private Sub chkHideSelection_Click()
    ucTreeView1.HideSelection = -chkHideSelection
End Sub

Private Sub chkSingleExpand_Click()
    ucTreeView1.SingleExpand = -chkSingleExpand
    chkLabelEdit.Enabled = Not -chkSingleExpand
End Sub

Private Sub chkTrackSelect_Click()
    ucTreeView1.TrackSelect = -chkTrackSelect
End Sub

Private Sub chkFullRowSelect_Click()
    ucTreeView1.FullRowSelect = -chkFullRowSelect
End Sub



'==== Current node

Private Sub txtText_Change()
    If (txtText.Tag = vbNullString) Then
        ucTreeView1.NodeText(ucTreeView1.SelectedNode) = txtText.Text
        lblFullPathVal.Caption = ucTreeView1.NodeFullPath(ucTreeView1.SelectedNode)
    End If
End Sub
 
Private Sub chkBold_Click()
    If (chkBold.Tag = vbNullString) Then
        ucTreeView1.NodeBold(ucTreeView1.SelectedNode) = -chkBold
    End If
End Sub

Private Sub chkGhosted_Click()
    If (chkGhosted.Tag = vbNullString) Then
        ucTreeView1.NodeGhosted(ucTreeView1.SelectedNode) = -chkGhosted
    End If
End Sub

Private Sub chkHilited_Click()
    If (chkHilited.Tag = vbNullString) Then
        ucTreeView1.NodeHilited(ucTreeView1.SelectedNode) = -chkHilited
    End If
End Sub



'====  Node relationship

Private Sub cmdRoot_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeRoot
End Sub

Private Sub cmdParent_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeParent(ucTreeView1.SelectedNode)
End Sub
Private Sub cmdChild_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeChild(ucTreeView1.SelectedNode)
End Sub

Private Sub cmdFirstVisible_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeFirstVisible()
End Sub
Private Sub cmdPrevious_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodePrevious(ucTreeView1.SelectedNode)
End Sub
Private Sub cmdNext_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeNext(ucTreeView1.SelectedNode)
End Sub
Private Sub cmdLastVisible_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeLastVisible()
End Sub

Private Sub cmdFirstSibling_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeFirstSibling(ucTreeView1.SelectedNode)
End Sub
Private Sub cmdPreviousSibling_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodePreviousSibling(ucTreeView1.SelectedNode)
End Sub
Private Sub cmdNextSibling_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeNextSibling(ucTreeView1.SelectedNode)
End Sub
Private Sub cmdLastSibling_Click()
    ucTreeView1.SelectedNode = ucTreeView1.NodeLastSibling(ucTreeView1.SelectedNode)
End Sub



'==== Adding/deleting nodes

Private Sub cmdFill_Click()
  
  Static lcBook As Long
  Static lKey   As Long
  
  Dim lC   As Long
  Dim lP   As Long
  Dim hB   As Long
  Dim hC   As Long
  Dim hP   As Long
    
    With ucTreeView1
        
        If (.NodeCount > 10000) Then
            Call MsgBox("Stop adding more nodes, please." & vbCrLf & vbCrLf & _
                        "Current node count: " & .NodeCount, vbExclamation)
            Exit Sub
        End If

        lcBook = lcBook + 1
        
        Call .SetRedrawMode(Enable:=False)
        
        lKey = lKey + 1
        hB = .AddNode(, , "K" & lKey, "Book #" & lcBook, 0, 1)
        
        For lC = 1 To 20
            lKey = lKey + 1
            hC = .AddNode(hB, , "K" & lKey, "Chapter #" & lC, 0, 1)
            For lP = 1 To 10
                lKey = lKey + 1
                hP = .AddNode(hC, , "K" & lKey, "Page #" & lP, 2, 2)
                lKey = lKey + 1
                Call .AddNode(hP, , "K" & lKey, "Note #1", 3, 3)
                lKey = lKey + 1
                Call .AddNode(hP, , "K" & lKey, "Note #2", 3, 3)
        Next lP, lC
        
        Call .Expand(hB)
        Call .SetRedrawMode(Enable:=True)
        
        Call .Scroll([sEnd])
    End With
End Sub

Private Sub cmdClear_Click()
    
    With ucTreeView1
        Call .SetRedrawMode(False)
        Call .Clear
        Call .SetRedrawMode(True)
    End With
    
    Call pvClearInfo
End Sub

Private Sub cmdDeleteCurrent_Click()
    
    With ucTreeView1
        Call .SetRedrawMode(False)
'       Call .HoldDeletePostProcess(True)  '(*)
        Call .DeleteNode(.SelectedNode)
'       Call .HoldDeletePostProcess(False) '(*)
        Call .SetRedrawMode(True)
        
        If (.NodeCount = 0) Then
            Call pvClearInfo
        End If
    End With
    
'(*) See ucTreeView.HoldDeletePostProcess() sub note!
End Sub

Private Sub cmdInsert_Click()
  
  Static lInsert As Long
  Dim eRelation  As Long
  Dim hNode      As Long
  
    lInsert = lInsert + 1

    Select Case True
        Case optRelation(0): eRelation = [rLast]
        Case optRelation(1): eRelation = [rFirst]
        Case optRelation(2): eRelation = [rSort]
        Case optRelation(3): eRelation = [rNext]
        Case optRelation(4): eRelation = [rPrevious]
    End Select

    With ucTreeView1
        hNode = .AddNode(Val(txthRelative.Text), eRelation, txtKey.Text, "Inserted node #" & lInsert, 4, 4)
        If (hNode) Then
            Call .EnsureVisible(hNode)
          Else
            Call MsgBox("Key already exists or unexpected error inserting node.", vbExclamation)
        End If
    End With
End Sub


'==== Raising events (See Debug window)

Private Sub ucTreeView1_GotFocus()
    Debug.Print "ucTreeView1_GotFocus"
End Sub
Private Sub ucTreeView1_LostFocus()
    Debug.Print "ucTreeView1_LostFocus"
End Sub

Private Sub ucTreeView1_Click()
    Debug.Print "ucTreeView1_Click"
End Sub

Private Sub ucTreeView1_BeforeLabelEdit(ByVal hNode As Long, Cancel As Integer)
    Debug.Print "ucTreeView1_BeforeLabelEdit"; hNode; Cancel
End Sub
Private Sub ucTreeView1_AfterLabelEdit(ByVal hNode As Long, Cancel As Integer, NewString As String)
    Debug.Print "ucTreeView1_AfterLabelEdit"; hNode; Cancel; NewString
    
    If (NewString <> vbNullString) Then
        ucTreeView1.NodeText(ucTreeView1.SelectedNode) = NewString
        Call ucTreeView1_NodeClick(ucTreeView1.SelectedNode)
    End If
End Sub

Private Sub ucTreeView1_SelectionChanged()
    Debug.Print "ucTreeView1_SelectionChanged"
    
    If (chkEnsureVisible) Then
        Call ucTreeView1.EnsureVisible(ucTreeView1.SelectedNode)
    End If
End Sub
Private Sub ucTreeView1_Collapse(ByVal hNode As Long)
    Debug.Print "ucTreeView1_Collapse"; hNode
End Sub
Private Sub ucTreeView1_ExpandBefore(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
    Debug.Print "ucTreeView1_ExpandBefore"; hNode; ExpandedOnce
End Sub
Private Sub ucTreeView1_ExpandAfter(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
    Debug.Print "ucTreeView1_ExpandAfter"; hNode; ExpandedOnce
End Sub

Private Sub ucTreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "ucTreeView1_KeyDown"; KeyCode; Shift
End Sub
Private Sub ucTreeView1_KeyPress(KeyAscii As Integer)
    Debug.Print "ucTreeView1_KeyPress"; KeyAscii
End Sub
Private Sub ucTreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print "ucTreeView1_KeyUp"; KeyCode; Shift
End Sub

Private Sub ucTreeView1_MouseEnter()
    Debug.Print "ucTreeView1_MouseEnter"
End Sub
Private Sub ucTreeView1_MouseLeave()
    Debug.Print "ucTreeView1_MouseLeave"
End Sub

Private Sub ucTreeView1_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Debug.Print "ucTreeView1_MouseDown"; Button; Shift; x; y
End Sub
Private Sub ucTreeView1_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
'   Debug.Print "ucTreeView1_MouseMove"; Button; Shift; x; y
End Sub
Private Sub ucTreeView1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Debug.Print "ucTreeView1_MouseUp"; Button; Shift; x; y
End Sub

Private Sub ucTreeView1_NodeClick(ByVal hNode As Long)
    Debug.Print "ucTreeView1_NodeClick - hNode = "; hNode
    
    lblhNodeVal.Caption = hNode
    lblKeyVal.Caption = ucTreeView1.GetNodeKey(hNode)
    lblFullPathVal.Caption = ucTreeView1.NodeFullPath(hNode)
    
    chkBold.Tag = 1
        chkBold.Value = -ucTreeView1.NodeBold(hNode)
    chkBold.Tag = vbNullString
    
    chkGhosted.Tag = 1
        chkGhosted.Value = -ucTreeView1.NodeGhosted(hNode)
    chkGhosted.Tag = vbNullString
    
    chkHilited.Tag = 1
        chkHilited.Value = -ucTreeView1.NodeHilited(hNode)
    chkHilited.Tag = vbNullString
    
    txtText.Tag = 1
        txtText.Text = ucTreeView1.NodeText(hNode)
    txtText.Tag = vbNullString
    
    txthRelative.Text = lblhNodeVal.Caption
End Sub
Private Sub ucTreeView1_NodeCheck(ByVal hNode As Long)
    Debug.Print "ucTreeView1_NodeCheck - hNode = "; hNode
End Sub
Private Sub ucTreeView1_NodeDblClick(ByVal hNode As Long)
    Debug.Print "ucTreeView1_NodeDblClick - hNode = "; hNode
End Sub

Private Sub ucTreeView1_Resize()
'   Debug.Print "ucTreeView1_Resize"
End Sub





'==== Misc.

Private Sub pvClearInfo()
    
    lblhNodeVal.Caption = "0"
    lblKeyVal.Caption = vbNullString
    lblFullPathVal.Caption = vbNullString
    
    chkBold.Value = 0
    chkGhosted.Value = 0
    chkHilited.Value = 0
    txtText.Text = vbNullString
    
    txthRelative.Text = lblhNodeVal.Caption
End Sub

Private Sub pvFlattenButton(ByVal hButton As Long)
    
  Dim lS As Long

    lS = GetWindowLong(hButton, GWL_STYLE)
    Call SetWindowLong(hButton, GWL_STYLE, lS Or BS_FLAT)
End Sub
