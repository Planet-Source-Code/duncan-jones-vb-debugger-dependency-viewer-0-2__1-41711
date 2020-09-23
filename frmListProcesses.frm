VERSION 5.00
Begin VB.Form frmListProcesses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select process to debug"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstProcesses 
      Height          =   5325
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   5640
      Width           =   975
   End
End
Attribute VB_Name = "frmListProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()

Me.Hide
Unload Me

End Sub

Private Sub cmdSelect_Click()

DebugProcess.Name = lstProcesses.List(lstProcesses.ListIndex)
AppSettings.DebugeeAppName = DebugProcess.Name
AppSettings.ProcessId = lstProcesses.ItemData(lstProcesses.ListIndex)
Call cmdcancel_Click

End Sub

Private Sub Form_Load()

Call FillProcessList(lstProcesses)

End Sub

Private Sub lstProcesses_Click()

Me.cmdSelect.Enabled = lstProcesses.ListIndex >= 0

End Sub
