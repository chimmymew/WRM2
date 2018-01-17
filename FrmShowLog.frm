VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmShowLog 
   Caption         =   "WRM2 Log"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12375
   Icon            =   "FrmShowLog.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   12375
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   2160
      Top             =   6180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.rtf"
      DialogTitle     =   "Save Log As"
      Filter          =   "*.rtf|*.rtf"
   End
   Begin VB.CommandButton CmdSaveAs 
      Caption         =   "Save As"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   6060
      Width           =   1275
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   10504
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"FrmShowLog.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier MonoThai"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmShowLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSaveAs_Click()
cdialog.ShowSave
RichTextBox1.SaveFile (cdialog.FileName)
End Sub

Private Sub Form_Resize()
RichTextBox1.Width = Me.Width - 200
RichTextBox1.Height = Me.Height - 1100
CmdSaveAs.Top = Me.Height - 1000
End Sub
