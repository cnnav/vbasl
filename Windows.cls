VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "windows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'*˵���� ����ģ���빤���ڵ�Window��ģ����������ģ�ʹ��ʱ��һ����
'*       ��ģ����СFisher��www.officefans.net��ԭ��������������ʹ�á����ơ��޸�
'*������ ����ȫ�����룬�����Ϸ�����ע��ԭ���ߺͳ�����лл������
'*******************************************************************************

Private colWindows As Collection
Property Get Count() As Long
    Count = colWindows.Count
End Property

Property Get Item(ByVal vID As Variant) As window
Attribute Item.VB_UserMemId = 0
    Set Item = colWindows(vID)
End Property


Private Sub Class_Initialize()
    Set colWindows = New Collection
End Sub


Private Sub Class_Terminate()
    Dim oWin As window
    For Each oWin In colWindows
        Set oWin = Nothing
    Next
    Set colWindows = Nothing
End Sub

Public Sub Add(ByRef oWin As window)
    colWindows.Add oWin, "s" & oWin.hwnd
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colWindows.[_NewEnum]
End Function