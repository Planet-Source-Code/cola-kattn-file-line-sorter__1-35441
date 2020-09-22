VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text file sorter"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowse2 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort text lines alphabetically."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4575
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Enter output file:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter source file:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is made by [Tech]
'You may use it for whatever you want, no copyrights.
'The idea is pretty easy (and good?) the program just reads every line from the source file
'and adds them to a list file that is set to sort the list alphabetically.
'Just adjust the form height to see the listbox.

Private Sub cmdBrowse_Click()
CommonDialog1.ShowOpen
txtFileName.Text = CommonDialog1.FileName
End Sub

Private Sub cmdBrowse2_Click()
CommonDialog1.ShowOpen
txtOutput.Text = CommonDialog1.FileName
End Sub

Private Sub cmdSort_Click()
On Error GoTo ErrorHandle
List1.Clear
Dim LineBuffer As String
Open txtFileName For Input As #1
    Do While Not EOF(1)
        Input #1, LineBuffer
        List1.AddItem LineBuffer
    Loop
Close #1
Open txtOutput.Text For Output As #2
    For I = 1 To List1.ListCount
        Print #2, List1.List(I)
    Next I
Close #2
MsgBox "Done", vbInformation
ErrorHandle:
MsgBox "Invalid filename or path", vbCritical, "Error"
End Sub
