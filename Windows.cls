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
'*说明： 此类模块与工程内的Window类模块是相关联的，使用时请一起导入
'*       本模块由小Fisher（www.officefans.net）原创，您可以自由使用、复制、修改
'*　　　 以下全部代码，在网上发表请注明原作者和出处，谢谢合作。
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
