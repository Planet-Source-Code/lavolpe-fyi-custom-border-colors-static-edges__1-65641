VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   Caption         =   "Custom Borders for Static Edge Controls"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   405
      Left            =   2970
      ScaleHeight     =   345
      ScaleWidth      =   2910
      TabIndex        =   44
      Top             =   5460
      Width           =   2970
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   405
      Left            =   1395
      TabIndex        =   43
      Top             =   5460
      Width           =   1350
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   7455
      TabIndex        =   41
      Top             =   1695
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Format          =   22806529
      CurrentDate     =   38880
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   5130
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sunken Sytem Color"
      Height          =   255
      Index           =   6
      Left            =   6915
      TabIndex        =   37
      Top             =   4365
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Raised System Color"
      Height          =   255
      Index           =   5
      Left            =   6915
      TabIndex        =   36
      Top             =   4005
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Flat Dual System Color"
      Height          =   255
      Index           =   4
      Left            =   6915
      TabIndex        =   35
      Top             =   3645
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Raised Custom Color"
      Height          =   255
      Index           =   3
      Left            =   6915
      TabIndex        =   34
      Top             =   3285
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sunken Custom Color"
      Height          =   255
      Index           =   2
      Left            =   6915
      TabIndex        =   33
      Top             =   2925
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Flat Dual Color"
      Height          =   255
      Index           =   1
      Left            =   6915
      TabIndex        =   32
      Top             =   2565
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Flat Single Color"
      Height          =   255
      Index           =   0
      Left            =   6915
      TabIndex        =   31
      Top             =   2205
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   570
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   585
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   4440
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   585
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   3000
      Style           =   1  'Simple Combo
      TabIndex        =   17
      Text            =   "Combo2"
      Top             =   585
      Width           =   1275
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Index           =   0
      Left            =   5880
      TabIndex        =   16
      Top             =   585
      Width           =   1485
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Index           =   1
      Left            =   5880
      TabIndex        =   13
      Top             =   1125
      Width           =   1485
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1440
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   1965
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   2540
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Items"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   5
      Left            =   4440
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1125
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   2985
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   1125
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1125
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset All to Default"
      Height          =   495
      Left            =   6945
      TabIndex        =   5
      Top             =   4890
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   3510
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3510
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1125
      Width           =   1275
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Index           =   0
      Left            =   3495
      TabIndex        =   1
      Top             =   1965
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1965
      Width           =   1275
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1440
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   3510
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   2540
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Items"
         Object.Width           =   1235
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1440
      Index           =   0
      Left            =   4980
      TabIndex        =   11
      Top             =   1965
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   2540
      _Version        =   393217
      Indentation     =   2
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1440
      Index           =   1
      Left            =   4980
      TabIndex        =   12
      Top             =   3510
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   2540
      _Version        =   393217
      Indentation     =   19
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Index           =   1
      Left            =   7440
      TabIndex        =   14
      Top             =   1125
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Index           =   0
      Left            =   7440
      TabIndex        =   15
      Top             =   585
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   1
      Left            =   3495
      TabIndex        =   40
      Top             =   5130
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Date Picker"
      Height          =   255
      Index           =   11
      Left            =   7470
      TabIndex        =   42
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Top row is the VB control un-modified, the second row is the same control, same properties with modified borders"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   180
      TabIndex        =   38
      Top             =   90
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "Tree View"
      Height          =   255
      Index           =   9
      Left            =   4980
      TabIndex        =   30
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "File List Box"
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   29
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "List View"
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   28
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "List Box"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Image Combo"
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   26
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Drive Combo"
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   25
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Combo Style 0"
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   24
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Combo Style 1"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   23
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Combo Style 2"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   22
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Text Box"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsBorders As cBorders

Private Sub Command1_Click()
    ' reset borders
    clsBorders.ReSetBorder List1(1).hwnd
    clsBorders.ReSetBorder Text1(1).hwnd
    clsBorders.ReSetBorder File1(1).hwnd
    clsBorders.ReSetBorder Combo1(1).hwnd
    clsBorders.ReSetBorder Combo1(3).hwnd
    clsBorders.ReSetBorder Combo1(5).hwnd
    clsBorders.ReSetBorder ListView1(1).hwnd
    clsBorders.ReSetBorder TreeView1(1).hwnd
    clsBorders.ReSetBorder Drive1(1).hwnd
    clsBorders.ReSetBorder ImageCombo1(1).hwnd
    clsBorders.ReSetBorder ProgressBar1(1).hwnd
    clsBorders.ReSetBorder DTPicker1.hwnd
End Sub

Private Sub SetCustomBorder()

    ' In the SetBorder call below, following parameters should be explained
    
    ' ///// Border Styles \\\\\
    ' bsFlat1Color. Solid 1-pixel border, 1 color (i.e., flat).
    '       Uses Shadow only
    ' bsFlat2Color. Left/Top borders are 1 color, right/bottom are another
    '       Uses Shadow & Highlight only
    ' bsSunken. Left/Top outer border are Shadow, Right/Bottom outer are HighLight
    '           Left/Top inner border are DarkShadow, Right/Bottom inner are LightShadow
    ' bsRaised. Left/Top outer border are HighLight, Right/Bottom outer are DarkShadow
    '           Left/Top inner border are LightShadow, Right/Bottom outer are Shadow
    
    ' ///// Colors \\\\\ vb system colors can be passed
    ' Shadow: 2nd darkest of 4 color borders; color for a single color border
    ' DarkShadow: the darkest of 4 color borders
    ' LightShadow: 2nd lightest of 4 color borders
    ' Highlight: lightest of 4 color borders
    ' Special values for the above 4 colors
    '   -1 = AutoShade. DarkShadow, LightShadow & Highlight are shades of Shadow
    '           DarkShadow = Shadow darkened by 90% or black whichever is greater
    '           LightShadow = Shadow lightened by 85% of its lightest value (white)
    '           Highlight = Shadow lightened by 100% (or vbWhite)
    '   -2 = System Colors: vb3DDKShadow, vbButtonShadow, vb3DLight, vbHighlight respectively
    '   -3 & -4 (Reserved) are used by the class to fake single borders on combo boxes
    
    ' ///// Control Type \\\\\
    ' Some controls have their borders drawn be VB on the control's client area whereas
    '   others are drawn in the non-client area as expected. Think of a form with no
    '   borders but you want borders so you draw it on the form (non-client area).
    '   VB combo boxes are very much like that scenario. Therefore, the control type
    '   needs to be known in advance so the class can handle those special cases.
    ' ctComboBox: use for comboboxes and drivecombo
    ' ctImageCombo: use for the image combobox
    ' ctListBox: use for listboxes and file listboxes
    ' ctOther: use for other controls like treeview, listview, progressbar, etc
    
    
    Dim borderType As BorderStyleOptions
    Dim Colors(0 To 3) As Long
    
    ' for all custom color samples, we will let the class autoshade the
    '   necessary colors based off of the passed primary color (vbBlue-ish)
    '   however, you can dictate any of the 4 colors
    Colors(0) = &HFF8080
    Colors(1) = bsAutoShade: Colors(2) = bsAutoShade: Colors(3) = bsAutoShade
    
    Select Case True
        Case Option1(0).Value ' flat single color
            borderType = bsFlat1Color
        Case Option1(1).Value ' flat two color
            borderType = bsFlat2Color
        Case Option1(2).Value ' sunken custom
            borderType = bsSunken
        Case Option1(3).Value ' raised custom
            borderType = bsRaised
        Case Option1(4).Value ' flat dual using system colors
            Colors(0) = bsSysDefault
            Colors(3) = bsSysDefault
            borderType = bsFlat2Color
        Case Option1(5).Value ' raised using system colors
            Colors(0) = bsSysDefault: Colors(1) = bsSysDefault
            Colors(2) = bsSysDefault: Colors(3) = bsSysDefault
            borderType = bsRaised
        Case Option1(6).Value ' sunken using system colors
            Colors(0) = bsSysDefault: Colors(1) = bsSysDefault
            Colors(2) = bsSysDefault: Colors(3) = bsSysDefault
            borderType = bsSunken
    End Select
    
    If clsBorders Is Nothing Then Set clsBorders = New cBorders
    clsBorders.SetBorder List1(1).hwnd, borderType, ctListBox, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder Text1(1).hwnd, borderType, ctTextBox, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder File1(1).hwnd, borderType, ctListBox, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder Combo1(1).hwnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder Combo1(3).hwnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder Combo1(5).hwnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder ListView1(1).hwnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder TreeView1(1).hwnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder Drive1(1).hwnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder ImageCombo1(1).hwnd, borderType, ctImageCombo, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder ProgressBar1(1).hwnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder DTPicker1.hwnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder Check1.hwnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
    clsBorders.SetBorder Picture1.hwnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
End Sub

Private Sub Form_Load()

' add some dummy list items
Dim I As Integer, xItem As Node, J As Integer

For I = 1 To 10
    List1(0).AddItem "Item " & I
    List1(1).AddItem "Item " & I
    ListView1(0).ListItems.Add , , "Item " & I
    ListView1(1).ListItems.Add , , "Item " & I
    Set xItem = TreeView1(0).Nodes.Add(, , , "Item " & I)
        TreeView1(0).Nodes.Add xItem, tvwChild, , "Sub Item"
    Set xItem = TreeView1(1).Nodes.Add(, , , "Item " & I)
        TreeView1(1).Nodes.Add xItem, tvwChild, , "Sub Item"
    For J = 0 To Combo1.UBound
        Combo1(J).AddItem "Item " & I
    Next
Next
For J = 0 To Combo1.UBound
    Combo1(J).ListIndex = J
Next
ProgressBar1(0).Value = 60
ProgressBar1(1).Value = 60
Picture1.Print "  Picture Box Control"
Command1.Enabled = False

End Sub

Private Sub Option1_Click(Index As Integer)
    Command1.Enabled = True
    SetCustomBorder
End Sub
