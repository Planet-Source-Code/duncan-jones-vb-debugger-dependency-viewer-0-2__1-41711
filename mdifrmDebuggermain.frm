VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm mdifrmDebuggermain 
   BackColor       =   &H00AA020B&
   Caption         =   "MCL VB Debugger and application dependency walker"
   ClientHeight    =   6585
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8640
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdifrmDebuggermain.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDebugMessages 
      Align           =   2  'Align Bottom
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   8580
      TabIndex        =   1
      Top             =   5370
      Width           =   8640
      Begin VB.TextBox txtEvents 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   0
         Width           =   8535
      End
   End
   Begin VB.PictureBox picModuleList 
      Align           =   3  'Align Left
      Height          =   5370
      Left            =   0
      ScaleHeight     =   5310
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   0
      Width           =   1875
      Begin ComctlLib.TreeView tvwDetailItems 
         Height          =   5295
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   9340
         _Version        =   327682
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         Appearance      =   0
      End
   End
   Begin MSComDlg.CommonDialog cdlgFileOpen 
      Left            =   2400
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select executable to debug"
      Filter          =   "*.exe|*.exe"
      FilterIndex     =   1
   End
   Begin VB.Menu mnuSession 
      Caption         =   "&Session"
      Begin VB.Menu mnuSessionSettings 
         Caption         =   "Se&ttings"
      End
      Begin VB.Menu mnuSessionSelectDebugee 
         Caption         =   "Select &debugee"
         Begin VB.Menu mnuDebugExisting 
            Caption         =   "Debug &existing process"
         End
         Begin VB.Menu mnuDebugNew 
            Caption         =   "Start and debug new process"
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStartDebugging 
         Caption         =   "&Start debugging"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuContinue 
         Caption         =   "&Continue"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStopDebugging 
         Caption         =   "Sto&p debugging"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSessionExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewModules 
         Caption         =   "&Modules"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewDebugEvents 
         Caption         =   "&Debug Events"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "&Windows"
         WindowList      =   -1  'True
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTile 
         Caption         =   "&Tile"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdifrmDebuggermain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTerminate As Boolean


Private WithEvents mSettings As cDebugAppSettings
Attribute mSettings.VB_VarHelpID = -1
Private mDebugging As Boolean

Public Property Get Debugging() As Boolean

    Debugging = mDebugging
    
End Property

Public Property Let Debugging(ByVal newval As Boolean)

    mDebugging = newval
    '\\ Set the menu states consistent with the current debug status
    mnuStartDebugging.Enabled = Not mDebugging
    mnuStopDebugging.Enabled = mDebugging
    mnuSessionExit.Enabled = Not mDebugging
    
    If Not Debugging Then
        bContinueOK = False
    End If
    
End Property


Public Sub LogEvent(ByVal EventType As String, ByVal ExtraData As String)

txtEvents.Text = EventType & String$(3, 9) & ExtraData & vbCrLf & txtEvents.Text

End Sub

Public Sub RefreshModulesList()

tvwDetailItems.Nodes.Clear

'\\ Header node - the process being debugged...
With tvwDetailItems
    .Nodes.Add , , DebugProcess.ICollectionItem_Key, DebugProcess.Name
    Dim cmod As cModule
    For Each cmod In DebugProcess.Modules
        .Nodes.Add DebugProcess.ICollectionItem_Key, tvwChild, cmod.ICollectionItem_Key, cmod.Name
    Next cmod
End With

End Sub

Public Property Get Terminated() As Boolean

    Terminated = mTerminate
    
End Property


Private Sub MDIForm_Load()

Set mSettings = AppSettings

'\\ If a command line exists, use it...
If Command$ <> "" Then
    mSettings.SetupFromCommandLine Command$
End If



End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'\\ If we are debugging an application, offer to close it...

'\\ Then terminate
mTerminate = True
Set mSettings = Nothing

End Sub


Private Sub mnuCascade_Click()

Call Arrange(vbCascade)

End Sub

Private Sub mnuContinue_Click()

bWaitDebugee = False

End Sub

Private Sub mnuDebugExisting_Click()

If AppSettings.DebugeeAppName = "" Then
    frmListProcesses.Show vbModal
End If

End Sub

Private Sub mnuDebugNew_Click()

cdlgFileOpen.ShowOpen
mSettings.DebugeeAppName = cdlgFileOpen.FileName

End Sub

Private Sub mnuHelpAbout_Click()

frmAbout.Show vbModal

End Sub


Private Sub mnuSessionExit_Click()

Unload Me

End Sub

Private Sub mnuSessionSettings_Click()

frmSettings.Show vbModal

End Sub

Private Sub mnuStartDebugging_Click()

Debugging = True
DebugMain.DebugLoop mSettings.ProcessId

End Sub

Private Sub mnuStopDebugging_Click()

Debugging = False

End Sub

Private Sub mnuTile_Click()

Call Arrange(vbTileHorizontal)

End Sub

Private Sub mnuViewDebugEvents_Click()

mnuViewDebugEvents.Checked = Not mnuViewDebugEvents.Checked
    picDebugMessages.Visible = mnuViewDebugEvents.Checked

End Sub

Private Sub mnuViewModules_Click()

mnuViewModules.Checked = Not mnuViewModules.Checked
picModuleList.Visible = (mnuViewModules.Checked)

End Sub

Private Sub mSettings_DebugeeNameChanged()

'\\ If there is a name, disable the select menus, otherwise make 'em enabled...
If mSettings.DebugeeAppName <> "" Then
    mnuSessionSelectDebugee.Enabled = False
    mnuStartDebugging.Caption = "&Start debugging " & mSettings.DebugeeAppName
    mnuStopDebugging.Caption = "Sto&p debugging " & mSettings.DebugeeAppName
    Debugging = False
Else
    mnuSessionSelectDebugee.Enabled = True
    mnuStartDebugging.Enabled = False
    mnuStopDebugging.Enabled = False
End If

End Sub


Private Sub picDebugMessages_Resize()

If picDebugMessages.Width > 10 Then
    txtEvents.Width = picDebugMessages.Width - 10
End If

End Sub


Private Sub picModuleList_Resize()

If picModuleList.Height > 20 Then
    Me.tvwDetailItems.Height = picModuleList.Height - 10
End If

End Sub

Private Sub tvwDetailItems_DblClick()

Dim bFound As Boolean
Dim fThis As Form
Dim fnew As frmPEImage

If Not tvwDetailItems.SelectedItem Is Nothing Then
    For Each fThis In Forms
        If fThis.Tag = tvwDetailItems.SelectedItem.Key Then
            bFound = True
            fThis.Show
        End If
    Next fThis
        If Not bFound Then
            Set fnew = New frmPEImage
            If tvwDetailItems.SelectedItem.Parent Is Nothing Then
                fnew.PortableImage = DebugProcess.Image
            Else
                Dim cmod As cModule
                For Each cmod In DebugProcess.Modules
                    If cmod.ICollectionItem_Key = tvwDetailItems.SelectedItem.Key Then
                        fnew.PortableImage = cmod.Image
                        Exit For
                    End If
                Next cmod
            End If
            fnew.Tag = tvwDetailItems.SelectedItem.Key
            fnew.Show
        End If
    'Next fThis
End If

End Sub


