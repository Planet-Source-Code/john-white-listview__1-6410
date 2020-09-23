VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_down 
      Caption         =   "Down"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmd_up 
      Caption         =   "Up"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Test"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Test"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Test"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Test"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Test"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Test"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Test"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

For i = 1 To 30
    Set itmx = ListView1.ListItems.Add(i, , "Test" & i)
        itmx.SubItems(1) = "Test" & i
        itmx.SubItems(2) = "Test" & i
        itmx.SubItems(3) = "Test" & i
        itmx.SubItems(4) = "Test" & i
        itmx.SubItems(5) = "Test" & i
        itmx.SubItems(6) = "Test" & i
        itmx.SubItems(7) = "Test" & i
Next i

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Set ListView1.DropHighlight = Nothing
    
End Sub

Private Sub cmd_up_Click()

If ListView1.SelectedItem.Index = 1 Then
Set ListView1.DropHighlight = ListView1.SelectedItem


Else
If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
    Set itmx = ListView1.ListItems.Add(ListView1.SelectedItem.Index - 1, , ListView1.SelectedItem.Text)
        itmx.SubItems(1) = ListView1.SelectedItem.SubItems(1)
        itmx.SubItems(2) = ListView1.SelectedItem.SubItems(2)
        itmx.SubItems(3) = ListView1.SelectedItem.SubItems(3)
        itmx.SubItems(4) = ListView1.SelectedItem.SubItems(4)
        itmx.SubItems(5) = ListView1.SelectedItem.SubItems(5)
        itmx.SubItems(6) = ListView1.SelectedItem.SubItems(6)
 
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - 1)
    Set ListView1.DropHighlight = ListView1.SelectedItem


Else
    Set itmx = ListView1.ListItems.Add(ListView1.SelectedItem.Index - 1, , ListView1.SelectedItem.Text)
        itmx.SubItems(1) = ListView1.SelectedItem.SubItems(1)
        itmx.SubItems(2) = ListView1.SelectedItem.SubItems(2)
        itmx.SubItems(3) = ListView1.SelectedItem.SubItems(3)
        itmx.SubItems(4) = ListView1.SelectedItem.SubItems(4)
        itmx.SubItems(5) = ListView1.SelectedItem.SubItems(5)
        itmx.SubItems(6) = ListView1.SelectedItem.SubItems(6)
 
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - 2)
    Set ListView1.DropHighlight = ListView1.SelectedItem

End If
End If
End Sub


Private Sub cmd_down_Click()

If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.ListItems.Count)
    Set ListView1.DropHighlight = ListView1.SelectedItem


Else
    Set itmx = ListView1.ListItems.Add(ListView1.SelectedItem.Index + 2, , ListView1.SelectedItem.Text)
        itmx.SubItems(1) = ListView1.SelectedItem.SubItems(1)
        itmx.SubItems(2) = ListView1.SelectedItem.SubItems(2)
        itmx.SubItems(3) = ListView1.SelectedItem.SubItems(3)
        itmx.SubItems(4) = ListView1.SelectedItem.SubItems(4)
        itmx.SubItems(5) = ListView1.SelectedItem.SubItems(5)
        itmx.SubItems(6) = ListView1.SelectedItem.SubItems(6)
 
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
    Set ListView1.DropHighlight = ListView1.SelectedItem

End If
End Sub

