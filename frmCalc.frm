VERSION 5.00
Begin VB.Form cmdAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XODE Multimedia UDT Calculator"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   330
      Left            =   4620
      TabIndex        =   5
      Top             =   2190
      Width           =   1740
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   330
      Left            =   2775
      TabIndex        =   4
      Top             =   2190
      Width           =   1740
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Height          =   330
      Left            =   4620
      TabIndex        =   3
      Top             =   1755
      Width           =   1740
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   2775
      TabIndex        =   2
      Top             =   1755
      Width           =   1740
   End
   Begin VB.Frame fraItem 
      Caption         =   " UDT Item "
      Height          =   1500
      Left            =   2775
      TabIndex        =   9
      Top             =   135
      Width           =   3585
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   870
         TabIndex        =   1
         Top             =   870
         Width           =   2295
      End
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "frmCalc.frx":0442
         Left            =   870
         List            =   "frmCalc.frx":0466
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   435
         Width           =   2295
      End
      Begin VB.Label lblSizeInfo 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   210
         Left            =   390
         TabIndex        =   11
         Top             =   900
         UseMnemonic     =   0   'False
         Width           =   360
      End
      Begin VB.Label lblTypeInfo 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   210
         Left            =   345
         TabIndex        =   10
         Top             =   495
         UseMnemonic     =   0   'False
         Width           =   405
      End
   End
   Begin VB.TextBox txtLength 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   4725
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ListBox lstUDT 
      Height          =   3840
      Left            =   225
      TabIndex        =   7
      Top             =   225
      Width           =   2295
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      Caption         =   "XODE Multimedia, Pascal 'Gherkin' vd Heiden (www.xodemultimedia.com)"
      Height          =   210
      Left            =   225
      TabIndex        =   13
      Top             =   4785
      Width           =   5295
   End
   Begin VB.Label lblNote 
      Caption         =   $"frmCalc.frx":04B9
      Height          =   465
      Left            =   225
      TabIndex        =   12
      Top             =   4275
      Width           =   6315
   End
   Begin VB.Label lblLengthInfo 
      AutoSize        =   -1  'True
      Caption         =   "UDT String Length:"
      Height          =   210
      Left            =   3210
      TabIndex        =   8
      Top             =   3765
      UseMnemonic     =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "cmdAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'These will hold the created UDT for processing
Private udtType As Collection
Private udtSize As Collection

Private Function GetUDTLength(ByVal StringCharLength As Long) As Long
     Dim Length As Long
     Dim i As Long
     
     'Go for all items in the UDT
     For i = 1 To udtType.Count
          
          'Check what item we deal with
          Select Case udtType.Item(i)
               
               Case 0    'Byte, can appear anywhere in the UDT
                    Length = Length + udtSize.Item(i)
                    
               Case 9    'String, can appear anywhere in the UDT, must get Unicode?
                    Length = Length + udtSize.Item(i) * StringCharLength
                    
               Case Else 'All other types must be padded to be aligned with their size
                    
                    'Add padding
                    If (Length Mod udtSize.Item(i)) > 0 Then _
                         Length = Length + udtSize.Item(i) - (Length Mod udtSize.Item(i))
                    
                    'Add item
                    Length = Length + udtSize.Item(i)
          End Select
     Next i
     
     'Return the length
     GetUDTLength = Length
End Function


Private Sub cmbType_Change()
     
     'Check if size is custom
     If cmbType.ItemData(cmbType.ListIndex) = -1 Then
          
          'Custom size
          txtSize.Enabled = True
          lblSizeInfo.Enabled = True
          txtSize = 1
          
          'Set focus on the size field
          txtSize.SetFocus
          txtSize.SelStart = 0
          txtSize.SelLength = Len(txtSize)
     Else
          
          'Fixed size
          txtSize.Enabled = False
          lblSizeInfo.Enabled = False
          txtSize = cmbType.ItemData(cmbType.ListIndex)
     End If
End Sub


Private Sub cmbType_Click()
     cmbType_Change
End Sub


Private Sub cmbType_KeyUp(KeyCode As Integer, Shift As Integer)
     cmbType_Change
End Sub


Private Sub cmdAdd_Click()
     
     'Only continue if a size is given
     If Trim$(txtSize.Text) <> "" Then
          
          'Add the Type
          udtType.Add cmbType.ListIndex
          udtSize.Add Val(txtSize.Text)
          If cmbType.ItemData(cmbType.ListIndex) = -1 Then
               lstUDT.AddItem cmbType.List(cmbType.ListIndex) & " * " & Val(txtSize.Text)
          Else
               lstUDT.AddItem cmbType.List(cmbType.ListIndex)
          End If
          
          'Recalculate the results
          txtLength = GetUDTLength(1)
     End If
End Sub

Private Sub cmdClear_Click()
     
     'Clear all
     lstUDT.Clear
     Set udtSize = New Collection
     Set udtType = New Collection
     
     'Recalculate the results
     txtLength = GetUDTLength(1)
End Sub

Private Sub cmdModify_Click()
     
     'Only continue if a selection is made
     If lstUDT.ListIndex >= 0 Then
          
          'Modify the Type
          udtType.Add cmbType.ListIndex, , , lstUDT.ListIndex + 1
          udtType.Remove lstUDT.ListIndex + 1
          udtSize.Add Val(txtSize.Text), , , lstUDT.ListIndex + 1
          udtSize.Remove lstUDT.ListIndex + 1
          If cmbType.ItemData(cmbType.ListIndex) = -1 Then
               lstUDT.List(lstUDT.ListIndex) = cmbType.List(cmbType.ListIndex) & " * " & Val(txtSize.Text)
          Else
               lstUDT.List(lstUDT.ListIndex) = cmbType.List(cmbType.ListIndex)
          End If
          
          'Recalculate the results
          txtLength = GetUDTLength(1)
     End If
End Sub

Private Sub cmdRemove_Click()
     
     'Only continue if a selection is made
     If lstUDT.ListIndex >= 0 Then
          
          'Remove the Type
          udtType.Remove lstUDT.ListIndex + 1
          udtSize.Remove lstUDT.ListIndex + 1
          lstUDT.RemoveItem lstUDT.ListIndex
          
          'Recalculate the results
          txtLength = GetUDTLength(1)
     End If
End Sub


Private Sub Form_Load()
     
     'New UDT
     cmdClear_Click
     
     'Select first type
     cmbType.ListIndex = 0
End Sub


Private Sub lstUDT_Click()
     
     'Reflect the settings
     cmbType.ListIndex = udtType.Item(lstUDT.ListIndex + 1)
     txtSize.Text = udtSize.Item(lstUDT.ListIndex + 1)
     txtSize.SelStart = 0
     txtSize.SelLength = Len(txtSize)
End Sub


