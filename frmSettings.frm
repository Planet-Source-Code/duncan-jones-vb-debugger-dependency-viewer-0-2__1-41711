VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Pause the debugee when: - "
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkSetting 
         Caption         =   "An RIP event"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   4095
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "A debug string is posted"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   4095
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "A DLL is unloaded"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   4095
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "A DLL is loaded"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   4095
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "The process exits"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   4095
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "The process is spawned"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "A thread is created"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   4095
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "An unhandled exception"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CheckboxIndexes
    chk_BreakOnException     'EXCEPTION_DEBUG_EVENT
    chk_BreakOnCreateThread   'CREATE_THREAD_DEBUG_EVENT
    chk_BreakOnCreateprocess    'CREATE_PROCESS_DEBUG_EVENT     EXIT_THREAD_DEBUG_EVENT = 4&
    chk_BreakOnExitProcess        'EXIT_PROCESS_DEBUG_EVENT
    chk_BreakOnLoadDll        'LOAD_DLL_DEBUG_EVENT
    chk_BreakOnUnloadDll    'UNLOAD_DLL_DEBUG_EVENT
    chk_BreakOnDebugString 'OUTPUT_DEBUG_STRING_EVENT
    chk_BreakOnRipEvent
End Enum
Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdOk_Click()

'\\ Save the settings
With AppSettings
    .BreakOnException = (chkSetting(chk_BreakOnException).Value = vbChecked)
    .BreakOnCreateThread = (chkSetting(chk_BreakOnCreateThread).Value = vbChecked)
    .BreakOnCreateprocess = (chkSetting(chk_BreakOnCreateprocess).Value = vbChecked)
    .BreakOnExitProcess = (chkSetting(chk_BreakOnExitProcess).Value = vbChecked)
    .BreakOnLoadDll = (chkSetting(chk_BreakOnLoadDll).Value = vbChecked)
    .BreakOnUnloadDll = (chkSetting(chk_BreakOnUnloadDll).Value = vbChecked)
    .BreakOnDebugString = (chkSetting(chk_BreakOnDebugString).Value = vbChecked)
    .BreakOnRipEvent = (chkSetting(chk_BreakOnRipEvent).Value = vbChecked)
End With

'\\ and exit
Call cmdCancel_Click

End Sub

Private Sub Form_Load()

'\\ Load the settings on the screen
With AppSettings
    chkSetting(chk_BreakOnException).Value = IIf(.BreakOnException, vbChecked, vbUnchecked)
    chkSetting(chk_BreakOnCreateThread).Value = IIf(.BreakOnCreateThread, vbChecked, vbUnchecked)
    chkSetting(chk_BreakOnCreateprocess).Value = IIf(.BreakOnCreateprocess, vbChecked, vbUnchecked)
    chkSetting(chk_BreakOnExitProcess).Value = IIf(.BreakOnExitProcess, vbChecked, vbUnchecked)
    chkSetting(chk_BreakOnLoadDll).Value = IIf(.BreakOnLoadDll, vbChecked, vbUnchecked)
    chkSetting(chk_BreakOnUnloadDll).Value = IIf(.BreakOnUnloadDll, vbChecked, vbUnchecked)
    chkSetting(chk_BreakOnDebugString).Value = IIf(.BreakOnDebugString, vbChecked, vbUnchecked)
    chkSetting(chk_BreakOnRipEvent).Value = IIf(.BreakOnRipEvent, vbChecked, vbUnchecked)
End With

End Sub
