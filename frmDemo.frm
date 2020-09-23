VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDemo 
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar H1 
      Height          =   255
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   26
      Top             =   4920
      Value           =   1
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   6240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "settings.ini ?"
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "colors"
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "change and save app.title"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "save"
         Height          =   375
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Text            =   "Settings Demo"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "new title:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame frmPath 
      Caption         =   "save a path"
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "save"
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.FileListBox File1 
         Height          =   480
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "CLICK HERE"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         Height          =   765
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   340
      Left            =   5880
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "store misc. settings"
      Height          =   3135
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "save changes"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "dont do anything!"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         Caption         =   "hide pathsaving"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "hide exit button"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "hide useless menu"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "save a string"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command3 
         Caption         =   "save"
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtCustom 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Text            =   "My custom Text"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "changes on HScrollbar are stored at app.exit"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No need for registry!     All Settings stored in application path!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3960
      TabIndex        =   25
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "resizeable form - resize me and change my screen position - all changes will be stored at program exit!"
      Height          =   1575
      Left            =   5880
      TabIndex        =   24
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Menu mnumain 
      Caption         =   "Menu"
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuuseless 
         Caption         =   "Useless Menu"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'keeps code clean

Private Sub cmdExit_Click()
On Error Resume Next
'saving screen position and forms height and width
If Me.WindowState <> vbMinimized Then 'check if form is not minimized
Call WriteToINI("settings", "Mainleft", Me.Left, App.Path & "/settings.ini")
Call WriteToINI("settings", "Maintop", Me.Top, App.Path & "/settings.ini")
Call WriteToINI("settings", "Mainwidth", Me.Width, App.Path & "/settings.ini")
Call WriteToINI("settings", "Mainheight", Me.Height, App.Path & "/settings.ini")
End If
'saving scrollbar value
Call WriteToINI("settings", "Scrollbarvalue", H1.Value, App.Path & "/settings.ini")

Unload Me: End
End Sub

Private Sub Command2_Click()
On Error Resume Next
Call WriteToINI("settings", "title", txtTitle.Text, App.Path & "/settings.ini")
End Sub

Private Sub Command3_Click()
On Error Resume Next
Call WriteToINI("settings", "CustomText", txtCustom.Text, App.Path & "/settings.ini")
End Sub

Private Sub Command4_Click()
On Error Resume Next
'storing all checkvalues and option values to text
Call WriteToINI("settings", "check1", Check1.Value, App.Path & "/settings.ini")
Call WriteToINI("settings", "check2", Check2.Value, App.Path & "/settings.ini")
Call WriteToINI("settings", "check3", Check3.Value, App.Path & "/settings.ini")
Call WriteToINI("settings", "check4", Check4.Value, App.Path & "/settings.ini")

Call WriteToINI("settings", "option1", Option1.Value, App.Path & "/settings.ini")
Call WriteToINI("settings", "option2", Option2.Value, App.Path & "/settings.ini")
End Sub

Private Sub Command5_Click()
On Error Resume Next
If txtPath.Text = "" Then
MsgBox "please choose a file first!", vbInformation, "Error"
Exit Sub
End If
Call WriteToINI("settings", "filepath", txtPath.Text, App.Path & "/settings.ini")
End Sub

Private Sub Command6_Click()


 On Error GoTo ErrHandler
 
 cdl.Color = frmDemo.BackColor
 cdl.CancelError = True
 cdl.Flags = cdlCCFullOpen Or cdlCCRGBInit
 On Error Resume Next
 cdl.ShowColor
 If Err.Number = 0 Then
    On Error GoTo ErrHandler

    frmDemo.BackColor = cdl.Color
Call WriteToINI("settings", "DemoColor", frmDemo.BackColor, App.Path & "/settings.ini")
MsgBox "Color changed and setting saved!", vbOKOnly, App.Title

 Else
    On Error GoTo ErrHandler

 End If

Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

ErrHandler:
 Screen.MousePointer = vbNormal
 MsgBox "Unexpected Error " & Err.Number & Err.Description, vbCritical
 Resume Exit_
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim iReturnValue
    iReturnValue = Shell("Notepad " & (App.Path + "\Settings.ini"), 1)
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
txtPath.Text = File1.Path '& "\" & File1.FileName
End Sub

Private Sub Form_Load()

If App.PrevInstance Then
MsgBox "program already running! shutting down!", vbInformation, "Error"
Unload Me: End
Exit Sub
End If


On Error GoTo ErrHandler  'Error handler

'Note:
'If the Program cant find the Settings.ini File you will get an error and the Program will continue
'with its default settings!
'the "Settings.ini" file should be stored in the App.Path
'By the way: You dont have to call it "settings.ini" you also can call it "billgates.sux" or "omfg.dll" ;-)

'Loading settings

'Programs screen position and its height, width
Me.Left = GetFromINI("settings", "Mainleft", App.Path & "/settings.ini")
Me.Top = GetFromINI("settings", "Maintop", App.Path & "/settings.ini")
Me.Width = GetFromINI("settings", "Mainwidth", App.Path & "/settings.ini")
Me.Height = GetFromINI("settings", "mainheight", App.Path & "/settings.ini")

'Loading  Backgroundcolor
frmDemo.BackColor = GetFromINI("settings", "DemoColor", App.Path & "/settings.ini")

'Loading custom text
txtCustom.Text = GetFromINI("settings", "CustomText", App.Path & "/settings.ini")

'Loading application title
frmDemo.Caption = GetFromINI("settings", "title", App.Path & "/settings.ini")
txtTitle.Text = GetFromINI("settings", "title", App.Path & "/settings.ini")

'Loading scrollbarvalue
H1.Value = GetFromINI("settings", "scrollbarvalue", App.Path & "/settings.ini")

'loading the checkboxes and optionvalues
Check1.Value = GetFromINI("settings", "check1", App.Path & "/settings.ini")
Check2.Value = GetFromINI("settings", "check2", App.Path & "/settings.ini")
Check3.Value = GetFromINI("settings", "check3", App.Path & "/settings.ini")
Check4.Value = GetFromINI("settings", "check4", App.Path & "/settings.ini")

Option1.Value = GetFromINI("settings", "option1", App.Path & "/settings.ini")
Option2.Value = GetFromINI("settings", "option2", App.Path & "/settings.ini")

'restoring settings by the checkvalues
If Check1.Value = Checked Then
mnumain.Visible = False
End If

If Check2.Value = Checked Then
cmdExit.Visible = False
End If

If Check3.Value = Checked Then
frmPath.Visible = False
End If

'restoring filepath ***usefull for mp3players***
txtPath.Text = GetFromINI("settings", "filepath", App.Path & "/settings.ini")
Dir1.Path = GetFromINI("settings", "filepath", App.Path & "/settings.ini")
Drive1.Drive = GetFromINI("settings", "filepath", App.Path & "/settings.ini")


Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

ErrHandler:
 Screen.MousePointer = vbNormal
 MsgBox "ErrorNumber: " & Err.Number & vbCrLf & "in: " & Err.Description, vbExclamation, App.Title & " " & App.Major & " " & App.Minor & " " & App.Revision
 Resume Exit_
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdExit_Click
End Sub

Private Sub txtTitle_Change()
On Error Resume Next
frmDemo.Caption = txtTitle.Text
End Sub
