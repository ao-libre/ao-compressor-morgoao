VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMain 
   Caption         =   "Compresor de recursos graficos"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApplyPatch 
      Caption         =   "Parchear"
      Height          =   735
      Left            =   3960
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame StatusFrame 
      Caption         =   "StatusFrame"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   5415
      Begin MSComctlLib.ProgressBar StatusBar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton cmdPatch 
      Caption         =   "Crear Parche"
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer"
      Height          =   735
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Text            =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Working Version :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApplyPatch_Click()
    Dim NewResourcePath As String
    Dim OldResourcePath As String
    Dim PatchPaths As String
    
    Dim NewVersion As Long
    Dim OldVersion As Long
    
    OldVersion = CLng(txtVersion.Text)
    NewVersion = OldVersion + 1
    
    NewResourcePath = App.Path & RESOURCE_PATH & NewVersion & "\"
    OldResourcePath = App.Path & RESOURCE_PATH & OldVersion & "\"
    PatchPaths = App.Path & PATCH_PATH & OldVersion & " to " & NewVersion & "\"
    
    'Check if the old resource file exists
    If Not FileExist(OldResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & OldResourcePath, , "Error"
        Exit Sub
    End If
    
    'We look if there's a patch to apply to the current version.
    If Not FileExist(PatchPaths, vbDirectory) Or Not FileExist(PatchPaths & GRH_PATCH_FILE, vbArchive) Then
        MsgBox "No existe el archivo .patch", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Look if the new resource file already exists.
    If FileExist(NewResourcePath, vbDirectory) Then
        If (MsgBox("Ya se encuentra un archivo parcheado, ¿Desea reparchear?", vbInformation + vbYesNo, "Error") = vbNo) Then Exit Sub
    Else
        MkDir NewResourcePath
    End If
    
    'Copy the Graphics.Ao file to the new ResourcesPath so we patch that file
    Call FileCopy(OldResourcePath & GRH_RESOURCE_FILE, NewResourcePath & GRH_RESOURCE_FILE)
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Parcheando de " & OldVersion & " a " & NewVersion
    
    'Patch!
#If SeguridadAlkon Then
    Dim FileMD5 As String
    FileMD5 = UCase(InputBox("Ingrese el MD5 del Graphics.AO final.", "Crear Parche", "00112233445566778899AABBDDEEFF00"))
    If Apply_Patch(NewResourcePath, PatchPaths, FileMD5, frmMain.StatusBar) Then
#Else
    If Apply_Patch(NewResourcePath, PatchPaths, frmMain.StatusBar) Then
#End If
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
        'Delete the Graphics File, its dummy
        Call Kill(NewResourcePath & GRH_RESOURCE_FILE)
    End If
    
    
    'Hide status
    Me.Height = 2055
    
End Sub

Private Sub cmdCompress_Click()
    Dim SourcePath As String
    Dim OutputPath As String
    
    SourcePath = App.Path & GRAPHIC_PATH
    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    
    'Check if the version already exists
    If FileExist(OutputPath & GRH_RESOURCE_FILE, vbNormal) Then
        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        If Not FileExist(OutputPath, vbDirectory) Then
            'Create this version folder
            MkDir OutputPath
        End If
    End If
    
    'Show status
    Me.Height = 2880
    StatusFrame.Caption = "Comprimiendo..."
    
    'Compress!
    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub cmdExtract_Click()
    Dim ResourcePath As String
    Dim OutputPath As String

    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "\"
    
    'Check if the resource file exists
    If Not FileExist(ResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Extrayendo..."
    
    'Extract!
    If Extract_Files(ResourcePath, OutputPath, StatusBar) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub cmdPatch_Click()
    Dim NewResourcePath As String
    Dim OldResourcePath As String
    Dim OutputPath As String
    
    Dim NewVersion As Long
    Dim OldVersion As Long
    
    NewVersion = CLng(txtVersion.Text)
    OldVersion = NewVersion - 1 'we patch from the last version
    
    NewResourcePath = App.Path & RESOURCE_PATH & NewVersion & "\"
    OldResourcePath = App.Path & RESOURCE_PATH & OldVersion & "\"
    OutputPath = App.Path & PATCH_PATH & OldVersion & " to " & NewVersion & "\"
    
    'Check if the new resource file exists
    If Not FileExist(NewResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & NewResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the old resource file exists
    If Not FileExist(OldResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos de la version anterior." & vbCrLf & OldResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("El parche ya se encuentra realizado. ¿Desea reparchear?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Armando el parche de " & OldVersion & " a " & NewVersion
    
    'Patch!
    If Make_Patch(NewResourcePath, OldResourcePath, OutputPath, StatusBar) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

