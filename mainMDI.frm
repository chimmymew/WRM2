VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm mainMDI 
   BackColor       =   &H8000000C&
   Caption         =   "WRM2 Database & Report Manager"
   ClientHeight    =   9195
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   14355
   Icon            =   "mainMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   4740
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   960
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.rtf"
      DialogTitle     =   "Open WRM Log"
      Filter          =   "*.rtf|*.rtf"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":1E2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   1588
      ButtonWidth     =   1746
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open Log"
            Object.ToolTipText     =   "Open WRM Log - �Դ WRM Log"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import Log"
            Object.ToolTipText     =   "Import WRM Log - ����� WRM Log"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete All"
            Object.ToolTipText     =   "Delete from Database - ź�����ŷ�����"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Language"
            Object.ToolTipText     =   "Language switch - ����¹����"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Database"
            Object.ToolTipText     =   "Database control - ἧ�Ǻ����ҹ������"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Exit - �͡�ҡ�����"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8820
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   20180
            MinWidth        =   20180
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpenWRMLog 
         Caption         =   "Open WRM log"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuShowData 
         Caption         =   "Show Data Panel"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import WRM log"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export Excel CSV"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Delete All Data"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuLang 
      Caption         =   "Language-����"
      Begin VB.Menu mnuEng 
         Caption         =   "English"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuThai 
         Caption         =   "������"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuApp 
         Caption         =   "WRM2 DB Report Manager"
      End
   End
End
Attribute VB_Name = "mainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LineData As String
Dim fld() As String
Dim Processdate As String
Dim SQL As String
Dim rsWarn As ADODB.Recordset
Dim Warning As String
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then
        Select Case Lang
        Case "Thai"
 
        If MsgBox("��ҹ��ͧ����͡�ҡ���������������?", vbYesNo Or vbQuestion) = vbNo Then Cancel = True
        
        Case "English"
         If MsgBox("Are you sure you want to quit?", vbYesNo Or vbQuestion) = vbNo Then Cancel = True
        End Select
        
    End If
End Sub

Private Sub mnuApp_Click()
frmAbout.Show
End Sub

Private Sub mnuDeleteAll_Click()
Select Case Lang
    Case "Thai"
    reply = MsgBox(Qt(1), vbYesNo, "�׹�ѹ���ź������")
    If reply = vbYes Then
    cn.Execute "delete from dose"
    MsgBox "ź�����ŷ���������", vbInformation, "��èѴ��ðҹ������"
    End If
    
    Case "English"
    reply = MsgBox(Qe(1), vbYesNo, "Confirm Delete")
    If reply = vbYes Then
    cn.Execute "delete from dose"
    MsgBox "Delete all data from database", vbInformation, "Database Manager"
    End If
End Select
End Sub

Private Sub mnuEng_Click()
Lang = "English"
MnuFile.Caption = "File"
mnuOpenWRMLog.Caption = "Open WRM2 Log"
mnuExit.Caption = "Exit"
mnuData.Caption = "Data"
mnuShowData.Caption = "Show Data Panel"
mnuImport.Caption = "Import WRM2 Log"
mnuExport.Caption = "Export Excel CSV"
mnuDeleteAll.Caption = "Delete All Data"
mnuAbout.Caption = "About"

FrmShowLog.Caption = "Show Log File"
FrmShowLog.CmdSaveAs.Caption = "Save As"

frmDataPanel.Caption = "Database Control Panel"
frmDataPanel.Label(0).Caption = "Serial Number"
frmDataPanel.Label(1).Caption = "User"
frmDataPanel.Label(2).Caption = "from Date"
frmDataPanel.Label(3).Caption = "to Date"
frmDataPanel.Label(4).Caption = "Accumulate dose"
frmDataPanel.Label(5).Caption = "Raport Organization"
frmDataPanel.txtOrganization.Text = "Thailand Institute of Nuclear Technology (Public Organization)"
frmDataPanel.cmdExport.Caption = "Export"
frmDataPanel.cmdCreateReport.Caption = "Create Report"

frmDataPanel.ListView1.ColumnHeaders(1).Text = "Date-Time"
frmDataPanel.ListView1.ColumnHeaders(2).Text = "Serial No."
frmDataPanel.ListView1.ColumnHeaders(3).Text = "User"
frmDataPanel.ListView1.ColumnHeaders(4).Text = "Location"
frmDataPanel.ListView1.ColumnHeaders(5).Text = "Dose"
frmDataPanel.ListView1.ColumnHeaders(6).Text = "Dose Rate"
frmDataPanel.ListView1.ColumnHeaders(7).Text = "Warning"

Toolbar1.Buttons(1).Caption = "Open Log"
Toolbar1.Buttons(2).Caption = "Import Log"
Toolbar1.Buttons(3).Caption = "Delete All"
Toolbar1.Buttons(4).Caption = "Language"
Toolbar1.Buttons(5).Caption = "Database"
Toolbar1.Buttons(6).Caption = "Exit"

End Sub

Private Sub mnuExit_Click()
Select Case Lang
        Case "Thai"
 
        If MsgBox("��ҹ��ͧ����͡�ҡ���������������?", vbYesNo Or vbQuestion) = vbYes Then End
        
        Case "English"
         If MsgBox("Are you sure you want to quit?", vbYesNo Or vbQuestion) = vbYes Then End
        End Select
End Sub

Private Sub mnuExport_Click()
On Error GoTo errdet:
If Lang = "English" Then frmDataPanel.cdialog.DialogTitle = "Export to Excel CSV" Else frmDataPanel.cdialog.DialogTitle = "���͡��ѧ Excel CSV"
frmDataPanel.cdialog.ShowSave

Open frmDataPanel.cdialog.FileName For Output As #1
If Lang = "English" Then Print #1, "Dosimeter Report" Else Print #1, "��§ҹ�ѧ�ջ�Шӵ�Ǻؤ��"
If Lang = "English" Then Print #1, "Organization," + frmDataPanel.txtOrganization.Text Else Print #1, "˹��§ҹ," + frmDataPanel.txtOrganization.Text
If Lang = "English" Then Print #1, "SerialNumber," + frmDataPanel.cmbSerial.Text Else Print #1, "�����Ţ����ͧ," + frmDataPanel.cmbSerial.Text
If Lang = "English" Then Print #1, "User," + frmDataPanel.cmbUser.Text Else Print #1, "�����ҹ," + frmDataPanel.cmbUser.Text

If Lang = "English" Then Print #1, "Report from," + Format(frmDataPanel.DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + ",to," + Format(frmDataPanel.DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") Else Print #1, "��§ҹ������ѹ���," + Format(frmDataPanel.DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + ",�֧�ѹ���," + Format(frmDataPanel.DTPicker1.Value, "dd/MM/yyyy HH:mm:ss")

Print #1, frmDataPanel.ListView1.ColumnHeaders(1) + "," + frmDataPanel.ListView1.ColumnHeaders(2) + "," + frmDataPanel.ListView1.ColumnHeaders(3) + "," + frmDataPanel.ListView1.ColumnHeaders(4) + "," + frmDataPanel.ListView1.ColumnHeaders(5) + "," + frmDataPanel.ListView1.ColumnHeaders(6) + "," + frmDataPanel.ListView1.ColumnHeaders(7)
For i = 1 To ListView1.ListItems.Count - 1
Print #1, frmDataPanel.ListView1.ListItems(i).Text + "," + frmDataPanel.ListView1.ListItems(i).SubItems(1) + "," + frmDataPanel.ListView1.ListItems(i).SubItems(2) + "," + frmDataPanel.ListView1.ListItems(i).SubItems(3) + "," + frmDataPanel.ListView1.ListItems(i).SubItems(4) + "," + frmDataPanel.ListView1.ListItems(i).SubItems(5) + "," + frmDataPanel.ListView1.ListItems(i).SubItems(6)
Next
Close #1
Exit Sub
errdet:
Close #1
MsgBox Err.Description, vbInformation, "WRM Data Export"
End Sub

Private Sub mnuImport_Click()
On Error GoTo errdet:
If Lang = "English" Then cdialog.DialogTitle = "Import WRM Log to Database" Else cdialog.DialogTitle = "����� WRM Log ���ҹ������"
cdialog.ShowOpen
Open cdialog.FileName For Input As #2
FrmShowLog.RichTextBox1.Text = ""
While Not EOF(2)

Line Input #2, LineData
fld = Split(LineData, ",")
If UBound(fld) > 4 Then
Processdate = Left(Right(cdialog.FileName, 6), 2) + "/" + Left(Right(cdialog.FileName, 8), 2) + "/" + Left(Right(cdialog.FileName, 12), 4) + " "

        For i = 1 To UBound(fld) - 1
            If IsNull(fld(i)) Then fld(i) = " "
        Next
  
       Set rsWarn = cn.Execute("select * from warning")
           Warning = ""
            While Not rsWarn.EOF
                If InStr(fld(11), rsWarn.Fields(1).Value) Then
                If Lang = "English" Then Warning = Warning + " " + rsWarn.Fields(2) Else Warning = Warning + " " + rsWarn.Fields(3)
                End If
            rsWarn.MoveNext
            Wend
            
            If Warning = "" Then
            If Lang = "English" Then Warning = "Status OK" Else Warning = "ʶҹл���"
            End If
            
SQL = "insert into dose (DoseDateTime,SerialNumber,UserName,Location,RadiationDose,DoseUnit,DoseCount,DoseRate,DoseRateUnit,DoseRateCount,Warning) values("
SQL = SQL + "#" + Processdate + fld(1) + "#,'" + fld(3) + "','" + fld(2) + "','" + fld(4) + "'," + fld(5) + ",'" + fld(6) + "'," + fld(7) + "," + fld(8) + ",'" + fld(9) + "'," + fld(10) + ",'" + Warning + "');"

Set rs = cn.Execute("select * from dose where SerialNumber='" + fld(3) + "' and DoseDateTime =" + "#" + Processdate + fld(1) + "#")

            If rs.EOF Then
                    cn.Execute SQL
                    If Lang = "English" Then StatusBar1.Panels(1).Text = "Process Log: " + cdialog.FileName + "at " + Processdate + fld(1) + " Serial:" + fld(3) Else StatusBar1.Panels(1).Text = "�����ż� Log: " + cdialog.FileName + " ��� " + Processdate + fld(1) + " �Ţ����ͧ:" + fld(3)
                    If Lang = "English" Then FrmShowLog.RichTextBox1.Text = FrmShowLog.RichTextBox1.Text + "Importing " + cdialog.FileName + "at " + Processdate + fld(1) + " Serial:" + fld(3) + vbCrLf Else FrmShowLog.RichTextBox1.Text = FrmShowLog.RichTextBox1.Text + "�������� " + cdialog.FileName + "�ѹ��� " + Processdate + fld(1) + " �����Ţ����ͧ:" + fld(3) + vbCrLf
                   Else
                    If Lang = "English" Then FrmShowLog.RichTextBox1.Text = FrmShowLog.RichTextBox1.Text + "Duplicate record " + Processdate + fld(1) + " Serial:" + fld(3) + vbCrLf Else FrmShowLog.RichTextBox1.Text = FrmShowLog.RichTextBox1.Text + "����¹���  " + Processdate + fld(1) + " �����Ţ����ͧ:" + fld(3) + vbCrLf
            End If

End If
Wend




Close #2
If Lang = "English" Then MsgBox "Import Complete!", vbInformation, "WRM Database Importer" Else MsgBox "������������!", vbInformation, "��ù���Ұҹ������"

Exit Sub
errdet:
Close #2
MsgBox Err.Description, vbInformation, "WRM Database Importer"

End Sub

Private Sub mnuOpenWRMLog_Click()
On Error GoTo errdet:
cdialog.ShowOpen
FrmShowLog.RichTextBox1.LoadFile (cdialog.FileName)
FrmShowLog.Show
If Lang = "English" Then StatusBar1.Panels(1).Text = "Open Log: " + cdialog.FileName Else StatusBar1.Panels(1).Text = "�Դ Log: " + cdialog.FileName
Exit Sub
errdet:
MsgBox Err.Description, vbInformation, "WRM2 DBReport Manager"
End Sub

Private Sub mnuShowData_Click()
frmDataPanel.Show
End Sub

Private Sub mnuThai_Click()
Lang = "Thai"
MnuFile.Caption = "���������"
mnuOpenWRMLog.Caption = "�Դ�����¡�úѹ�֡ WRM2"
mnuExit.Caption = "�͡�ҡ�����"
mnuData.Caption = "������"
mnuShowData.Caption = "�ʴ�ἧ�Ǻ���������"
mnuImport.Caption = "����������¡�úѹ�֡ WRM2"
mnuExport.Caption = "���͡�����ٻẺ Excel CSV"
mnuDeleteAll.Caption = "ź�����ŷ�����㹰ҹ������"
mnuAbout.Caption = "����ǡѺ"

FrmShowLog.Caption = "�ʴ� Log File"
FrmShowLog.CmdSaveAs.Caption = "�ѹ�֡��"


frmDataPanel.Caption = "˹�ҵ�ҧ�Ǻ����ҹ������"
frmDataPanel.Label(0).Caption = "��������ͧ��Ǩ�Ѵ"
frmDataPanel.Label(1).Caption = "���ͼ����ҹ"
frmDataPanel.Label(2).Caption = "�ҡ�ѹ���"
frmDataPanel.Label(3).Caption = "�֧�ѹ���"
frmDataPanel.Label(4).Caption = "����ҳ�ѧ������"
frmDataPanel.Label(5).Caption = "����˹��§ҹ/ͧ���"
frmDataPanel.txtOrganization.Text = "ʶҺѹ෤����չ����������觪ҵ� (ͧ������Ҫ�)"
frmDataPanel.cmdExport.Caption = "���͡"
frmDataPanel.cmdCreateReport.Caption = "���ҧ��§ҹ"

frmDataPanel.ListView1.ColumnHeaders(1).Text = "�ѹ���-����"
frmDataPanel.ListView1.ColumnHeaders(2).Text = "��������ͧ"
frmDataPanel.ListView1.ColumnHeaders(3).Text = "�����"
frmDataPanel.ListView1.ColumnHeaders(4).Text = "ʶҹ���"
frmDataPanel.ListView1.ColumnHeaders(5).Text = "����ҳ�ѧ��"
frmDataPanel.ListView1.ColumnHeaders(6).Text = "�ѵ���ѧ��"
frmDataPanel.ListView1.ColumnHeaders(7).Text = "��ͤ�����͹"

Toolbar1.Buttons(1).Caption = "�Դ Log"
Toolbar1.Buttons(2).Caption = "����� Log"
Toolbar1.Buttons(3).Caption = "ź������"
Toolbar1.Buttons(4).Caption = "����"
Toolbar1.Buttons(5).Caption = "�ҹ������"
Toolbar1.Buttons(6).Caption = "�͡"
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(2).Text = Date
StatusBar1.Panels(3).Text = Time
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button
    Case "Open Log"
        mnuOpenWRMLog_Click
    Case "Import Log"
           mnuImport_Click
    Case "Delete All"
        mnuDeleteAll_Click
    Case "Language"
        If Lang = "English" Then mnuThai_Click Else mnuEng_Click
    Case "Database"
        frmDataPanel.Show
    Case "Exit"
        mnuExit_Click
    
    
    Case "�Դ Log"
        mnuOpenWRMLog_Click
    Case "����� Log"
        mnuImport_Click
    Case "ź������"
        mnuDeleteAll_Click
    Case "����"
        If Lang = "English" Then mnuThai_Click Else mnuEng_Click
    Case "�ҹ������"
        frmDataPanel.Show
    Case "�͡"
        mnuExit_Click
        
        
        
End Select

End Sub
