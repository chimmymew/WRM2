Attribute VB_Name = "wrm_module"
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public Lang As String
Public Qt(20) As String
Public Qe(20) As String



Sub main()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

Lang = "English"
Qt(0) = "��ҹ��ͧ����͡�ҡ���������������?"
Qt(1) = "��ҹ��ͧ���ź�����ŷ������͡�ҡ�ҹ����������������?" + vbCrLf + "��á�зӹ����������ö��͹��Ѻ���ա"
Qe(0) = "Are you sure to exit WRM2 DBReport Manager program?"
Qe(1) = "Are you sure to DELETE all data from the database?" + vbCrLf + "This process cannot undo."

cn.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + App.Path + "\wrm2.mdb;Persist Security Info=False;")
mainMDI.Show



End Sub
