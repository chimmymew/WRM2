VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDataPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Control Panel"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12135
   Icon            =   "frmDataPanel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12135
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   9720
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.csv"
      DialogTitle     =   "Export to CSV"
      Filter          =   "*.csv|*.csv"
   End
   Begin VB.CommandButton cmdCreateReport 
      Caption         =   "Create Report"
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
      Left            =   8220
      TabIndex        =   14
      Top             =   6660
      Width           =   1335
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
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
      Left            =   6780
      TabIndex        =   13
      Top             =   6660
      Width           =   1335
   End
   Begin VB.TextBox txtOrganization 
      BeginProperty Font 
         Name            =   "AngsanaUPC"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1860
      TabIndex        =   12
      Text            =   "Thailand Institute of Nuclear Technology (Public Organization)"
      Top             =   6660
      Width           =   4695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5595
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   9869
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date-Time"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Serial No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dose"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Dose Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Warning"
         Object.Width           =   7832
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   4380
      TabIndex        =   5
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   163905539
      CurrentDate     =   43093
   End
   Begin VB.ComboBox cmbUser 
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   540
      Width           =   1635
   End
   Begin VB.ComboBox cmbSerial 
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1635
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   4380
      TabIndex        =   7
      Top             =   540
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   163905539
      CurrentDate     =   43093
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Report Organization"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblAccumulate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0.000 mSv"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   6660
      TabIndex        =   9
      Top             =   420
      Width           =   3555
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Accumulate dose"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   7440
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "to Date"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3120
      TabIndex        =   6
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Date from"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   180
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial Number"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1395
   End
End
Attribute VB_Name = "frmDataPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fromDate As String
Dim toDate As String
Dim iList As ListItem
Dim rsWarn As ADODB.Recordset
Dim Warning As String
Dim Unit As String
Dim Accumulation As Single

Private Sub cmbSerial_Click()
ListView1.ListItems.Clear
Accumulation = 0
'------------------------------- All Serial & User -----------------------------------------

If cmbSerial.Text = "*" And cmbUser.Text = "*" Then
Set rs = cn.Execute("select * from dose where DoseDateTime between #" + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + "# and #" + Format(DTPicker2.Value, "dd/MM/yyyy HH:mm:ss") + "# order by dosedatetime asc")
 While Not rs.EOF
 
 If rs.Fields(0).Value <> vbNull Then
 Warning = ""
 
If rs.Fields(5) <> vbNull Then Accumulation = rs.Fields(5)
If rs.Fields(6) <> vbNull Then Unit = rs.Fields(6)

Set iList = ListView1.ListItems.Add(, , Format(rs.Fields(1).Value, "dd/MM/yyyy HH:mm:ss"))
iList.ForeColor = vbBlue
 If rs.Fields(2).Value <> vbNull Then iList.SubItems(1) = rs.Fields(2).Value Else iList.SubItems(1) = " "
 iList.ListSubItems(1).ForeColor = vbBlack
 If rs.Fields(3).Value <> vbNull Then iList.SubItems(2) = rs.Fields(3).Value Else iList.SubItems(2) = " "
 iList.ListSubItems(2).ForeColor = vbBlack
 If rs.Fields(4).Value <> vbNull Then iList.SubItems(3) = rs.Fields(4).Value Else iList.SubItems(3) = " "
 iList.ListSubItems(3).ForeColor = vbBlack
  If rs.Fields(5).Value <> vbNull And rs.Fields(6).Value <> vbNull Then iList.SubItems(4) = Format(rs.Fields(5).Value, "0.000") + " " + rs.Fields(6).Value Else iList.SubItems(4) = " "
 iList.ListSubItems(4).ForeColor = vbBlack
 If rs.Fields(8).Value <> vbNull And rs.Fields(9).Value <> vbNull Then iList.SubItems(5) = Format(rs.Fields(8).Value, "0.000") + " " + rs.Fields(9).Value Else iList.SubItems(5) = " "
 iList.ListSubItems(5).ForeColor = vbBlack
 
         If rs.Fields(11) = "Status OK" Or rs.Fields(11) = "สถานะปกติ" Then
         iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = QBColor(2)
          Else
          iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = vbRed
          iList.ForeColor = vbRed
          End If
          

End If
 
 rs.MoveNext
Wend
End If

'---------------------------- All Serial specify User ------------------------------------------

If cmbSerial.Text = "*" And cmbUser.Text <> "*" Then
Set rs = cn.Execute("select * from dose where DoseDateTime between #" + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + "# and #" + Format(DTPicker2.Value, "dd/MM/yyyy HH:mm:ss") + "# and UserName = '" + cmbUser.Text + "' order by dosedatetime asc")
 While Not rs.EOF
 
 If rs.Fields(0).Value <> vbNull Then
 Warning = ""
 
If rs.Fields(5) <> vbNull Then Accumulation = rs.Fields(5)
If rs.Fields(6) <> vbNull Then Unit = rs.Fields(6)

Set iList = ListView1.ListItems.Add(, , Format(rs.Fields(1).Value, "dd/MM/yyyy HH:mm:ss"))
iList.ForeColor = vbBlue
 If rs.Fields(2).Value <> vbNull Then iList.SubItems(1) = rs.Fields(2).Value Else iList.SubItems(1) = " "
 iList.ListSubItems(1).ForeColor = vbBlack
 If rs.Fields(3).Value <> vbNull Then iList.SubItems(2) = rs.Fields(3).Value Else iList.SubItems(2) = " "
 iList.ListSubItems(2).ForeColor = vbBlack
 If rs.Fields(4).Value <> vbNull Then iList.SubItems(3) = rs.Fields(4).Value Else iList.SubItems(3) = " "
 iList.ListSubItems(3).ForeColor = vbBlack
  If rs.Fields(5).Value <> vbNull And rs.Fields(6).Value <> vbNull Then iList.SubItems(4) = Format(rs.Fields(5).Value, "0.000") + " " + rs.Fields(6).Value Else iList.SubItems(4) = " "
 iList.ListSubItems(4).ForeColor = vbBlack
 If rs.Fields(8).Value <> vbNull And rs.Fields(9).Value <> vbNull Then iList.SubItems(5) = Format(rs.Fields(8).Value, "0.000") + " " + rs.Fields(9).Value Else iList.SubItems(5) = " "
 iList.ListSubItems(5).ForeColor = vbBlack
 
     If rs.Fields(11) = "Status OK" Or rs.Fields(11) = "สถานะปกติ" Then
         iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = QBColor(2)
          Else
          iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = vbRed
          iList.ForeColor = vbRed
          End If
    
    
 End If
 rs.MoveNext
Wend
End If

'-----------------------------------------All User specify Serial-------------------------------------------------

If cmbSerial.Text <> "*" And cmbUser.Text = "*" Then
Set rs = cn.Execute("select * from dose where DoseDateTime between #" + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + "# and #" + Format(DTPicker2.Value, "dd/MM/yyyy HH:mm:ss") + "# and SerialNumber = '" + cmbSerial.Text + "' order by dosedatetime asc")
 While Not rs.EOF
 
 If rs.Fields(0).Value <> vbNull Then
 
 Warning = ""
If rs.Fields(5) <> vbNull Then Accumulation = rs.Fields(5)
If rs.Fields(6) <> vbNull Then Unit = rs.Fields(6)


Set iList = ListView1.ListItems.Add(, , Format(rs.Fields(1).Value, "dd/MM/yyyy HH:mm:ss"))
iList.ForeColor = vbBlue
 If rs.Fields(2).Value <> vbNull Then iList.SubItems(1) = rs.Fields(2).Value Else iList.SubItems(1) = " "
 iList.ListSubItems(1).ForeColor = vbBlack
 If rs.Fields(3).Value <> vbNull Then iList.SubItems(2) = rs.Fields(3).Value Else iList.SubItems(2) = " "
 iList.ListSubItems(2).ForeColor = vbBlack
 If rs.Fields(4).Value <> vbNull Then iList.SubItems(3) = rs.Fields(4).Value Else iList.SubItems(3) = " "
 iList.ListSubItems(3).ForeColor = vbBlack
  If rs.Fields(5).Value <> vbNull And rs.Fields(6).Value <> vbNull Then iList.SubItems(4) = Format(rs.Fields(5).Value, "0.000") + " " + rs.Fields(6).Value Else iList.SubItems(4) = " "
 iList.ListSubItems(4).ForeColor = vbBlack
 If rs.Fields(8).Value <> vbNull And rs.Fields(9).Value <> vbNull Then iList.SubItems(5) = Format(rs.Fields(8).Value, "0.000") + " " + rs.Fields(9).Value Else iList.SubItems(5) = " "
 iList.ListSubItems(5).ForeColor = vbBlack
 
      If rs.Fields(11) = "Status OK" Or rs.Fields(11) = "สถานะปกติ" Then
         iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = QBColor(2)
          Else
          iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = vbRed
          iList.ForeColor = vbRed
          End If
          
   End If
 rs.MoveNext
Wend
End If

'-----------------------------------------specify User specify Serial-------------------------------------------------

If cmbSerial.Text <> "*" And cmbUser.Text <> "*" Then
Set rs = cn.Execute("select * from dose where DoseDateTime between #" + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + "# and #" + Format(DTPicker2.Value, "dd/MM/yyyy HH:mm:ss") + "# and SerialNumber = '" + cmbSerial.Text + "' and UserName ='" + cmbUser.Text + "' order by dosedatetime asc")
 While Not rs.EOF
 
 If rs.Fields(0).Value <> vbNull Then
 
 Warning = ""
If rs.Fields(5) <> vbNull Then Accumulation = rs.Fields(5)
If rs.Fields(6) <> vbNull Then Unit = rs.Fields(6)


Set iList = ListView1.ListItems.Add(, , Format(rs.Fields(1).Value, "dd/MM/yyyy HH:mm:ss"))
iList.ForeColor = vbBlue
 If rs.Fields(2).Value <> vbNull Then iList.SubItems(1) = rs.Fields(2).Value Else iList.SubItems(1) = " "
 iList.ListSubItems(1).ForeColor = vbBlack
 If rs.Fields(3).Value <> vbNull Then iList.SubItems(2) = rs.Fields(3).Value Else iList.SubItems(2) = " "
 iList.ListSubItems(2).ForeColor = vbBlack
 If rs.Fields(4).Value <> vbNull Then iList.SubItems(3) = rs.Fields(4).Value Else iList.SubItems(3) = " "
 iList.ListSubItems(3).ForeColor = vbBlack
  If rs.Fields(5).Value <> vbNull And rs.Fields(6).Value <> vbNull Then iList.SubItems(4) = Format(rs.Fields(5).Value, "0.000") + " " + rs.Fields(6).Value Else iList.SubItems(4) = " "
 iList.ListSubItems(4).ForeColor = vbBlack
 If rs.Fields(8).Value <> vbNull And rs.Fields(9).Value <> vbNull Then iList.SubItems(5) = Format(rs.Fields(8).Value, "0.000") + " " + rs.Fields(9).Value Else iList.SubItems(5) = " "
 iList.ListSubItems(5).ForeColor = vbBlack
 
         If rs.Fields(11) = "Status OK" Or rs.Fields(11) = "สถานะปกติ" Then
         iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = QBColor(2)
          Else
          iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = vbRed
          iList.ForeColor = vbRed
          End If
 
 End If
 rs.MoveNext
Wend
End If


lblAccumulate.Caption = Format(Accumulation, "0.000") + " " + Unit


End Sub

Private Sub cmbUser_click()
ListView1.ListItems.Clear
Accumulation = 0
'------------------------------- All Serial & User -----------------------------------------

If cmbSerial.Text = "*" And cmbUser.Text = "*" Then
Set rs = cn.Execute("select * from dose where DoseDateTime between #" + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + "# and #" + Format(DTPicker2.Value, "dd/MM/yyyy HH:mm:ss") + "# order by dosedatetime asc")
 While Not rs.EOF
 
 If rs.Fields(0).Value <> vbNull Then
 Warning = ""
 
If rs.Fields(5) <> vbNull Then Accumulation = rs.Fields(5)
If rs.Fields(6) <> vbNull Then Unit = rs.Fields(6)

Set iList = ListView1.ListItems.Add(, , Format(rs.Fields(1).Value, "dd/MM/yyyy HH:mm:ss"))
iList.ForeColor = vbBlue
 If rs.Fields(2).Value <> vbNull Then iList.SubItems(1) = rs.Fields(2).Value Else iList.SubItems(1) = " "
 iList.ListSubItems(1).ForeColor = vbBlack
 If rs.Fields(3).Value <> vbNull Then iList.SubItems(2) = rs.Fields(3).Value Else iList.SubItems(2) = " "
 iList.ListSubItems(2).ForeColor = vbBlack
 If rs.Fields(4).Value <> vbNull Then iList.SubItems(3) = rs.Fields(4).Value Else iList.SubItems(3) = " "
 iList.ListSubItems(3).ForeColor = vbBlack
  If rs.Fields(5).Value <> vbNull And rs.Fields(6).Value <> vbNull Then iList.SubItems(4) = Format(rs.Fields(5).Value, "0.000") + " " + rs.Fields(6).Value Else iList.SubItems(4) = " "
 iList.ListSubItems(4).ForeColor = vbBlack
 If rs.Fields(8).Value <> vbNull And rs.Fields(9).Value <> vbNull Then iList.SubItems(5) = Format(rs.Fields(8).Value, "0.000") + " " + rs.Fields(9).Value Else iList.SubItems(5) = " "
 iList.ListSubItems(5).ForeColor = vbBlack
 
         If rs.Fields(11) = "Status OK" Or rs.Fields(11) = "สถานะปกติ" Then
         iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = QBColor(2)
          Else
          iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = vbRed
          iList.ForeColor = vbRed
          End If
 
 End If
 
 
 rs.MoveNext
Wend
End If

'---------------------------- All Serial specify User ------------------------------------------

If cmbSerial.Text = "*" And cmbUser.Text <> "*" Then
Set rs = cn.Execute("select * from dose where DoseDateTime between #" + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + "# and #" + Format(DTPicker2.Value, "dd/MM/yyyy HH:mm:ss") + "# and UserName = '" + cmbUser.Text + "' order by dosedatetime asc")
 While Not rs.EOF
 
 If rs.Fields(0).Value <> vbNull Then
 Warning = ""
 
If rs.Fields(5) <> vbNull Then Accumulation = rs.Fields(5)
If rs.Fields(6) <> vbNull Then Unit = rs.Fields(6)

Set iList = ListView1.ListItems.Add(, , Format(rs.Fields(1).Value, "dd/MM/yyyy HH:mm:ss"))
iList.ForeColor = vbBlue
 If rs.Fields(2).Value <> vbNull Then iList.SubItems(1) = rs.Fields(2).Value Else iList.SubItems(1) = " "
 iList.ListSubItems(1).ForeColor = vbBlack
 If rs.Fields(3).Value <> vbNull Then iList.SubItems(2) = rs.Fields(3).Value Else iList.SubItems(2) = " "
 iList.ListSubItems(2).ForeColor = vbBlack
 If rs.Fields(4).Value <> vbNull Then iList.SubItems(3) = rs.Fields(4).Value Else iList.SubItems(3) = " "
 iList.ListSubItems(3).ForeColor = vbBlack
  If rs.Fields(5).Value <> vbNull And rs.Fields(6).Value <> vbNull Then iList.SubItems(4) = Format(rs.Fields(5).Value, "0.000") + " " + rs.Fields(6).Value Else iList.SubItems(4) = " "
 iList.ListSubItems(4).ForeColor = vbBlack
 If rs.Fields(8).Value <> vbNull And rs.Fields(9).Value <> vbNull Then iList.SubItems(5) = Format(rs.Fields(8).Value, "0.000") + " " + rs.Fields(9).Value Else iList.SubItems(5) = " "
 iList.ListSubItems(5).ForeColor = vbBlack
 
         If rs.Fields(11) = "Status OK" Or rs.Fields(11) = "สถานะปกติ" Then
         iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = QBColor(2)
          Else
          iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = vbRed
          iList.ForeColor = vbRed
          End If
 
 End If
 
 rs.MoveNext
Wend
End If

'-----------------------------------------All User specify Serial-------------------------------------------------

If cmbSerial.Text <> "*" And cmbUser.Text = "*" Then
Set rs = cn.Execute("select * from dose where DoseDateTime between #" + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + "# and #" + Format(DTPicker2.Value, "dd/MM/yyyy HH:mm:ss") + "# and SerialNumber = '" + cmbSerial.Text + "' order by dosedatetime asc")
 While Not rs.EOF
 
 If rs.Fields(0).Value <> vbNull Then
 
 Warning = ""
If rs.Fields(5) <> vbNull Then Accumulation = rs.Fields(5)
If rs.Fields(6) <> vbNull Then Unit = rs.Fields(6)


Set iList = ListView1.ListItems.Add(, , Format(rs.Fields(1).Value, "dd/MM/yyyy HH:mm:ss"))
iList.ForeColor = vbBlue
 If rs.Fields(2).Value <> vbNull Then iList.SubItems(1) = rs.Fields(2).Value Else iList.SubItems(1) = " "
 iList.ListSubItems(1).ForeColor = vbBlack
 If rs.Fields(3).Value <> vbNull Then iList.SubItems(2) = rs.Fields(3).Value Else iList.SubItems(2) = " "
 iList.ListSubItems(2).ForeColor = vbBlack
 If rs.Fields(4).Value <> vbNull Then iList.SubItems(3) = rs.Fields(4).Value Else iList.SubItems(3) = " "
 iList.ListSubItems(3).ForeColor = vbBlack
  If rs.Fields(5).Value <> vbNull And rs.Fields(6).Value <> vbNull Then iList.SubItems(4) = Format(rs.Fields(5).Value, "0.000") + " " + rs.Fields(6).Value Else iList.SubItems(4) = " "
 iList.ListSubItems(4).ForeColor = vbBlack
 If rs.Fields(8).Value <> vbNull And rs.Fields(9).Value <> vbNull Then iList.SubItems(5) = Format(rs.Fields(8).Value, "0.000") + " " + rs.Fields(9).Value Else iList.SubItems(5) = " "
 iList.ListSubItems(5).ForeColor = vbBlack
 
         If rs.Fields(11) = "Status OK" Or rs.Fields(11) = "สถานะปกติ" Then
         iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = QBColor(2)
          Else
          iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = vbRed
          iList.ForeColor = vbRed
          End If
 
 End If
 
 rs.MoveNext
Wend
End If


'-----------------------------------------specify User specify Serial-------------------------------------------------

If cmbSerial.Text <> "*" And cmbUser.Text <> "*" Then
Set rs = cn.Execute("select * from dose where DoseDateTime between #" + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + "# and #" + Format(DTPicker2.Value, "dd/MM/yyyy HH:mm:ss") + "# and SerialNumber = '" + cmbSerial.Text + "' and UserName ='" + cmbUser.Text + "' order by dosedatetime asc")
 While Not rs.EOF
 
 If rs.Fields(0).Value <> vbNull Then
 
 Warning = ""
If rs.Fields(5) <> vbNull Then Accumulation = rs.Fields(5)
If rs.Fields(6) <> vbNull Then Unit = rs.Fields(6)


Set iList = ListView1.ListItems.Add(, , Format(rs.Fields(1).Value, "dd/MM/yyyy HH:mm:ss"))
iList.ForeColor = vbBlue
 If rs.Fields(2).Value <> vbNull Then iList.SubItems(1) = rs.Fields(2).Value Else iList.SubItems(1) = " "
 iList.ListSubItems(1).ForeColor = vbBlack
 If rs.Fields(3).Value <> vbNull Then iList.SubItems(2) = rs.Fields(3).Value Else iList.SubItems(2) = " "
 iList.ListSubItems(2).ForeColor = vbBlack
 If rs.Fields(4).Value <> vbNull Then iList.SubItems(3) = rs.Fields(4).Value Else iList.SubItems(3) = " "
 iList.ListSubItems(3).ForeColor = vbBlack
  If rs.Fields(5).Value <> vbNull And rs.Fields(6).Value <> vbNull Then iList.SubItems(4) = Format(rs.Fields(5).Value, "0.000") + " " + rs.Fields(6).Value Else iList.SubItems(4) = " "
 iList.ListSubItems(4).ForeColor = vbBlack
 If rs.Fields(8).Value <> vbNull And rs.Fields(9).Value <> vbNull Then iList.SubItems(5) = Format(rs.Fields(8).Value, "0.000") + " " + rs.Fields(9).Value Else iList.SubItems(5) = " "
 iList.ListSubItems(5).ForeColor = vbBlack
 
        If rs.Fields(11) = "Status OK" Or rs.Fields(11) = "สถานะปกติ" Then
         iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = QBColor(2)
          Else
          iList.SubItems(6) = rs.Fields(11)
          iList.ListSubItems(6).ForeColor = vbRed
          iList.ForeColor = vbRed
          End If
 
 End If
 
 rs.MoveNext
Wend
End If


lblAccumulate.Caption = Format(Accumulation, "0.000") + " " + Unit
End Sub

Private Sub cmdCreateReport_Click()
If Lang = "English" Then
Report.Sections("Section4").Controls.Item("lblReportTitle").Caption = "Personal Dosimeter Report"
Report.Sections("Section4").Controls.Item("lblUser").Caption = "User : " + cmbUser.Text
Report.Sections("Section4").Controls.Item("lblSerial").Caption = "Serial : " + cmbSerial.Text
Report.Sections("Section4").Controls.Item("lblOrganization").Caption = "Organization : " + txtOrganization.Text
Report.Sections("Section4").Controls.Item("lblAccumulate").Caption = "Accumulation dose : " + lblAccumulate.Caption

Report.Sections("Section2").Controls.Item("lblDateTime").Caption = "Date-Time"
Report.Sections("Section2").Controls.Item("lblSerialData").Caption = "Serial"
Report.Sections("Section2").Controls.Item("lblUserData").Caption = "User"
Report.Sections("Section2").Controls.Item("lblLocation").Caption = "Location"
Report.Sections("Section2").Controls.Item("lblDose").Caption = "Dose "
Report.Sections("Section2").Controls.Item("lblDoseRate").Caption = "Dose rate "
Report.Sections("Section2").Controls.Item("lblWarning").Caption = "Warning"

Report.Sections("Section3").Controls.Item("lblPage").Caption = "Page"
Report.Sections("Section3").Controls.Item("lblOf").Caption = "of"
Report.Caption = "Report Viewer"

Else
Report.Sections("Section4").Controls.Item("lblReportTitle").Caption = "รายงานเครื่องวัดรังสีประจำตัวบุคคล"
Report.Sections("Section4").Controls.Item("lblUser").Caption = "ชื่อผู้ใช้งาน : " + cmbUser.Text
Report.Sections("Section4").Controls.Item("lblSerial").Caption = "เลขเครื่อง: " + cmbSerial.Text
Report.Sections("Section4").Controls.Item("lblOrganization").Caption = "หน่วยงาน : " + txtOrganization.Text
Report.Sections("Section4").Controls.Item("lblAccumulate").Caption = "ปริมาณรังสีสะสม : " + lblAccumulate.Caption

Report.Sections("Section2").Controls.Item("lblDateTime").Caption = "วันที่-เวลา"
Report.Sections("Section2").Controls.Item("lblSerialData").Caption = "เลขเครื่อง"
Report.Sections("Section2").Controls.Item("lblUserData").Caption = "ชื่อผู้ใช้งาน"
Report.Sections("Section2").Controls.Item("lblLocation").Caption = "ตำแหน่ง"
Report.Sections("Section2").Controls.Item("lblDose").Caption = "ปริมาณรังสี"
Report.Sections("Section2").Controls.Item("lblDoseRate").Caption = "อัตราการได้รับ"
Report.Sections("Section2").Controls.Item("lblWarning").Caption = "แจ้งเตือน"

Report.Sections("Section3").Controls.Item("lblPage").Caption = "หน้า"
Report.Sections("Section3").Controls.Item("lblOf").Caption = "จาก"

Report.Caption = "หน้าต่างรายงาน"
End If


Set Report.DataSource = rs



Report.Show
End Sub

Private Sub cmdExport_Click()
On Error GoTo errdet:
If Lang = "English" Then cdialog.DialogTitle = "Export to Excel CSV" Else cdialog.DialogTitle = "ส่งออกไปยัง Excel CSV"
cdialog.ShowSave

Open cdialog.FileName For Output As #1
If Lang = "English" Then Print #1, "Dosimeter Report" Else Print #1, "รายงานรังสีประจำตัวบุคคล"
If Lang = "English" Then Print #1, "Organization," + txtOrganization.Text Else Print #1, "หน่วยงาน," + txtOrganization.Text
If Lang = "English" Then Print #1, "SerialNumber," + cmbSerial.Text Else Print #1, "หมายเลขเครื่อง," + cmbSerial.Text
If Lang = "English" Then Print #1, "User," + cmbUser.Text Else Print #1, "ผู้ใช้งาน," + cmbUser.Text

If Lang = "English" Then Print #1, "Report from," + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + ",to," + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") Else Print #1, "รายงานตั้งแต่วันที่," + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss") + ",ถึงวันที่," + Format(DTPicker1.Value, "dd/MM/yyyy HH:mm:ss")

Print #1, ListView1.ColumnHeaders(1) + "," + ListView1.ColumnHeaders(2) + "," + ListView1.ColumnHeaders(3) + "," + ListView1.ColumnHeaders(4) + "," + ListView1.ColumnHeaders(5) + "," + ListView1.ColumnHeaders(6) + "," + ListView1.ColumnHeaders(7)
For i = 1 To ListView1.ListItems.Count - 1
Print #1, ListView1.ListItems(i).Text + "," + ListView1.ListItems(i).SubItems(1) + "," + ListView1.ListItems(i).SubItems(2) + "," + ListView1.ListItems(i).SubItems(3) + "," + ListView1.ListItems(i).SubItems(4) + "," + ListView1.ListItems(i).SubItems(5) + "," + ListView1.ListItems(i).SubItems(6)
Next
Close #1
Exit Sub
errdet:
Close #1
MsgBox Err.Description, vbInformation, "WRM Data Export"
End Sub

Private Sub Form_Load()
Set rsWarn = New ADODB.Recordset
fromDate = "01/" + Trim(Str(Month(Now))) + "/" + Trim(Str(Year(Now))) + " 00:00:00"
toDate = Trim(Str(Day(Now))) + "/" + Trim(Str(Month(Now))) + "/" + Trim(Str(Year(Now))) + " 23:59:59"

DTPicker1.Value = fromDate
DTPicker2.Value = toDate


cmbUser.Clear
cmbUser.AddItem "*"
Set rs = cn.Execute("select distinct UserName from dose")

While Not rs.EOF
If rs.Fields(0).Value <> vbNull Then cmbUser.AddItem rs.Fields(0).Value
rs.MoveNext
Wend

cmbSerial.Clear
cmbSerial.AddItem "*"
Set rs = cn.Execute("select distinct SerialNumber from dose")

While Not rs.EOF
cmbSerial.AddItem rs.Fields(0).Value
rs.MoveNext
Wend

End Sub

