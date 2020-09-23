VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test of LVListBox"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fundo 
      Caption         =   " Save/Load LVListBox "
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   5775
      Begin VB.CommandButton CmdLoadSave 
         Caption         =   "Execute operation"
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox ChkLoadSave 
         Caption         =   "Load/Save with format"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   720
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.TextBox TxtFilename 
         Height          =   285
         Left            =   2880
         MaxLength       =   64
         TabIndex        =   14
         Text            =   "c:\lvlistbox.txt"
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton OptLoadSave 
         Caption         =   "Load from file"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton OptLoadSave 
         Caption         =   "Save to file"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         Caption         =   "Filename:"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame Fundo 
      Caption         =   " Interface methods and properties "
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   5775
      Begin VB.ComboBox CmbSortOrder 
         Height          =   315
         ItemData        =   "FrmTest.frx":0000
         Left            =   3840
         List            =   "FrmTest.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox ChkInterface 
         Caption         =   "Sorted"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox ChkInterface 
         Caption         =   "Multi Select"
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox ChkInterface 
         Caption         =   "Label Edit"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox ChkInterface 
         Caption         =   "Hide Selection"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox ChkInterface 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox ChkInterface 
         Caption         =   "Check Boxes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":0025
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":017F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":02D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":0873
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":09CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":0F67
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":1501
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":165B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":17B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":1C07
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":1D61
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":1EBB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin PrjTest.LVListBox LVListBox1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3201
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      MouseIcon       =   "FrmTest.frx":2015
      MultiSelect     =   -1  'True
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Caption         =   "Control test:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkInterface_Click(Index As Integer)
 Select Case Index
  Case 0
   LVListBox1.CheckBoxes = ChkInterface(Index).Value * -1
   
  Case 1
   LVListBox1.Enabled = ChkInterface(Index).Value * -1
   
  Case 2
   LVListBox1.HideSelection = ChkInterface(Index).Value * -1
   
  Case 3
   LVListBox1.LabelEdit = IIf(ChkInterface(Index).Value = 0, lvwManual, lvwAutomatic)
   
  Case 4
   LVListBox1.MultiSelect = ChkInterface(Index).Value * -1
   
  Case 5
   LVListBox1.Sorted = ChkInterface(Index).Value * -1
   CmbSortOrder.Enabled = LVListBox1.Sorted
    
   
 End Select
End Sub


Private Sub CmbSortOrder_Click()
 LVListBox1.SortOrder = CmbSortOrder.ListIndex
End Sub


Private Sub CmdLoadSave_Click()
 If OptLoadSave(0).Value = True Then
  If Trim(TxtFilename.Text) = "" Then
   MsgBox "Input the filename !", 16
   TxtFilename.SetFocus
   Exit Sub
  End If
  
  If LVListBox1.SaveToFile(TxtFilename.Text, ChkLoadSave.Value * -1) = False Then
   MsgBox "Error on save LVListBox content to file " & TxtFilename.Text, 16
  Else
   MsgBox "Content saved into " & TxtFilename.Text & " with sucess !", 48
  End If
 End If
 
 
 If OptLoadSave(1).Value = True Then
  If Trim(TxtFilename.Text) = "" Then
   MsgBox "Input the filename !", 16
   TxtFilename.SetFocus
   Exit Sub
  End If
    
  If Dir(TxtFilename.Text) = "" Then
   MsgBox "The source file " & TxtFilename.Text & " does not exist !", 16
   TxtFilename.SetFocus
   Exit Sub
  End If
  
  If LVListBox1.LoadFromFile(TxtFilename.Text, ChkLoadSave.Value * -1) = False Then
   MsgBox "Error on load LVListBox content from file " & TxtFilename.Text, 16
  Else
   MsgBox "Content loaded from " & TxtFilename.Text & " with sucess !", 48
  End If
 End If
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim L As Integer
Dim B As Boolean

 With LVListBox1
  Set .ImageList = ImageList1

  For I = 1 To 13
    L = L + 1
    If L > ImageList1.ListImages.Count Then L = 1
   .AddItem "Line item " & I & UCase(Chr(63 + I)), L, IIf(L = 1 Or L = 7 Or L = 3 Or L = 5, vbRed, -1), IIf(L = 3 Or L = 40, True, False)
  Next
  .ListIndex = .ListCount
  .Refresh
 End With
 
ChkInterface(0).Value = LVListBox1.CheckBoxes * -1
ChkInterface(1).Value = LVListBox1.Enabled * -1
ChkInterface(2).Value = LVListBox1.HideSelection * -1
ChkInterface(3).Value = IIf(LVListBox1.LabelEdit = lvwManual, 0, 1)
ChkInterface(4).Value = LVListBox1.MultiSelect * -1
ChkInterface(5).Value = LVListBox1.Sorted * -1
CmbSortOrder.Enabled = ChkInterface(5).Value * -1
CmbSortOrder.ListIndex = LVListBox1.SortOrder
End Sub

Private Sub LVListBox1_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete And LVListBox1.ListCount > 0 Then
  LVListBox1.RemoveItem LVListBox1.ListIndex
 End If
 
 If KeyCode = vbKeyF2 Then
  LVListBox1.StartEdit
 End If
End Sub
