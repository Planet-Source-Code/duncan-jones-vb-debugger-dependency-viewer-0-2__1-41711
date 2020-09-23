VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPEImage 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   7065
   Begin VB.Frame fraResources 
      Caption         =   "Resources"
      Height          =   4335
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   6735
   End
   Begin VB.Frame fraSections 
      Caption         =   "Section:"
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6735
      Begin VB.ComboBox cmbSections 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   0
         Width           =   3015
      End
      Begin VB.CheckBox chkCharacteristics 
         Appearance      =   0  'Flat
         Caption         =   "Contains Code"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox chkCharacteristics 
         Appearance      =   0  'Flat
         Caption         =   "Contains Initialised Data"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   720
         Width           =   3615
      End
      Begin VB.CheckBox chkCharacteristics 
         Appearance      =   0  'Flat
         Caption         =   "Contains Uninitialsed Data"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Top             =   960
         Width           =   3615
      End
      Begin VB.CheckBox chkCharacteristics 
         Appearance      =   0  'Flat
         Caption         =   "Readable"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   6
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CheckBox chkCharacteristics 
         Appearance      =   0  'Flat
         Caption         =   "Writeable"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   5
         Top             =   1560
         Width           =   3615
      End
      Begin VB.CheckBox chkCharacteristics 
         Appearance      =   0  'Flat
         Caption         =   "Executable"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   4
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label lblPhysicalAddress 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Physical Address :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblVirtualAddress 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Virtual Address :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Byte Alignment :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblByteAlign 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.ListBox lstExports 
      Height          =   4350
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6735
   End
   Begin ComctlLib.TreeView tvwImports 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   327682
      Indentation     =   353
      LabelEdit       =   1
      Style           =   4
      Appearance      =   1
   End
   Begin ComctlLib.TabStrip tsOther 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8916
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Imports"
            Key             =   "IMPORT"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exports"
            Key             =   "EXPORTS"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sections"
            Key             =   "SECTIONS"
            Object.Tag             =   "Sections"
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Resources"
            Key             =   "RESOURCES"
            Object.Tag             =   "Resources"
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPEImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mImage As cPortableExecutableImage

Public Property Let PortableImage(ByVal newImage As cPortableExecutableImage)

If Not newImage Is Nothing Then
    Caption = "Module:- " & newImage.Name
    Set mImage = newImage
    With mImage
        Dim cSect As cSection
        cmbSections.Clear
        For Each cSect In .Sections
            cmbSections.AddItem cSect.SectionName
        Next cSect
        tvwImports.Nodes.Clear
        lstExports.Clear
        If newImage.ImportDependencies.Count > 0 Then
            Dim impd As cImportDependency
            Dim impf As cFunction
            For Each impd In newImage.ImportDependencies
                Call tvwImports.Nodes.Add(, , impd.Name, impd.Name)
                For Each impf In impd.Functions
                    Call tvwImports.Nodes.Add(impd.Name, tvwChild, impd.Name & ":" & impf.Name, impf.Name & " (" & Hex(impf.Ordinal) & ") at " & impf.ProcAddress)
                Next impf
            Next impd
        End If
        If newImage.ExportedFunctions.Count > 0 Then
            For Each impf In newImage.ExportedFunctions
                lstExports.AddItem impf.Name & Chr$(9) & Chr$(9) & impf.Ordinal & Chr$(9) & impf.ProcAddress
                lstExports.ItemData(lstExports.NewIndex) = impf.ProcAddress
            Next impf
        End If
    End With
Else
    cmbSections.Clear
    Set mImage = Nothing
End If

End Property

Private Sub cmbSections_Click()

    Dim cSect As cSection
    For Each cSect In mImage.Sections
        If cmbSections.Text = cSect.SectionName Then
            With cSect
                lblPhysicalAddress.Caption = .PhysicalAddress
                lblVirtualAddress.Caption = .VirtualAddress
                lblByteAlign.Caption = .ByteAlignment
                chkCharacteristics(0).Value = IIf(.ContainsCode, vbChecked, vbUnchecked)
                chkCharacteristics(1).Value = IIf(.ContainsInitialisedData, vbChecked, vbUnchecked)
                chkCharacteristics(2).Value = IIf(.ContainsUninitialisedData, vbChecked, vbUnchecked)
                chkCharacteristics(3).Value = IIf(.Readable, vbChecked, vbUnchecked)
                chkCharacteristics(4).Value = IIf(.Writeable, vbChecked, vbUnchecked)
                chkCharacteristics(5).Value = IIf(.Executable, vbChecked, vbUnchecked)
            End With
            Exit For
        End If
    Next cSect
    
End Sub

Private Sub Form_Load()

Call tsOther_Click

End Sub

Private Sub Form_Resize()

If Width > 60 Then
    tsOther.Width = Width - 60
    If tsOther.Width > (fraResources.Left * 2) Then
        fraResources.Width = tsOther.Width - (fraResources.Left * 2)
        fraSections.Width = fraResources.Width
        tvwImports.Width = fraResources.Width
        lstExports.Width = fraResources.Width
    End If
End If
If Height > 200 Then
    tsOther.Height = Height - 200
    If tsOther.Height > (fraResources.Top * 2) Then
        fraResources.Height = tsOther.Height - (fraResources.Top * 2)
        fraSections.Height = fraResources.Height
        tvwImports.Height = fraResources.Height
        lstExports.Height = fraResources.Height
    End If
End If


End Sub

Private Sub tsOther_Click()

tsOther.ZOrder vbBringToFront
If tsOther.SelectedItem.Key = "IMPORT" Then
    tvwImports.ZOrder vbBringToFront
ElseIf tsOther.SelectedItem.Key = "RESOURCES" Then
    fraResources.ZOrder vbBringToFront
ElseIf tsOther.SelectedItem.Key = "SECTIONS" Then
    fraSections.ZOrder vbBringToFront
Else
    lstExports.ZOrder vbBringToFront
End If

End Sub
