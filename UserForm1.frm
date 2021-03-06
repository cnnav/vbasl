VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_CHILD = &H40000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const IDC_ARROW As Long = 32512&

Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const R2_NOT = 6 '  Dn
Private Const R2_NOTXORPEN = 10  '  DPxn
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long



Private WithEvents Timer1 As xTimer
Attribute Timer1.VB_VarHelpID = -1
Dim bFrameVisible As Boolean
Dim hwnd As Long
Dim hdc As Long
Dim hPen As Long, hPenPrev As Long
Dim hBrush As Long, hBrushPrev As Long
Dim hCursor As Long
Dim oMouseIcon As StdPicture
Dim oRect As RECT

Const PAGE_GENERAL = 0
Const PAGE_POSITION = 1
Const PAGE_STYLE = 2
Const PAGE_PICTURE = 3

Dim StyleName()
Dim StyleLong()
Dim oWin As window
Dim bEnableEvents As Boolean
Dim lTimeButtonDown As Long
Dim pStart As POINTAPI
Dim lOldStyle As Long

 

Private Sub chkEnabled_Click()
    If Not oWin Is Nothing And bEnableEvents Then oWin.Enabled = chkEnabled.Value
End Sub

Private Sub chkTasckBar_Click()
    If Not oWin Is Nothing And bEnableEvents Then oWin.ShowInTaskBar = chkTasckBar.Value
End Sub

Private Sub chkTopMost_Click()
    If Not oWin Is Nothing And bEnableEvents Then oWin.TopMost = chkTopMost.Value
End Sub

Private Sub chkVisible_Click()
    If Not oWin Is Nothing And bEnableEvents Then oWin.Visible = chkVisible.Value

End Sub

Private Sub cmdApplyStyle_Click()
    Dim lStyle As Long
    Dim i As Integer

    If Not oWin Is Nothing And bEnableEvents Then
        lStyle = oWin.Style
 
        For i = 0 To lstStyle.ListCount - 2
            
            If lstStyle.Selected(i) Then
                lStyle = lStyle Or lstStyle.List(i, 1)
            Else
                lStyle = lStyle And Not CLng(lstStyle.List(i, 1))
            End If
        Next
 
        If lStyle <> oWin.Style Then
            oWin.Style = lStyle
        End If
    End If
End Sub

Private Sub cmdResetStyle_Click()
    If Not oWin Is Nothing Then
        oWin.Style = lOldStyle
        For i = 0 To lstStyle.ListCount - 2
            lstStyle.Selected(i) = CLng(lstStyle.List(i, 1)) And lOldStyle
        Next
    End If
End Sub

Private Sub cmdSave_Click()
    If Not imgWindowPic.Picture Is Nothing Then
        Dim sTargetFile As String
        sTargetFile = GetSaveAsFileName
        If Len(sTargetFile) > 0 Then
            If LCase(Right(sTargetFile, 4)) <> ".bmp" Then
                sTargetFile = sTargetFile & ".bmp"
            End If
            SavePicture imgWindowPic.Picture, sTargetFile
        End If
    End If
End Sub
 

Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim w As window
    Dim oNode As Node
    Set w = New window
    w.hwnd = 0
    ClearTreeView TreeView1
    With TreeView1.Nodes
        Set oNode = .Add(, , "hwnd" & w.hwnd, w.ClassName & "(" & w.Caption & ")")
        Set oNode.Tag = w
        ListChildWindow w, oNode
        .Item(1).Expanded = True
        .Item(1).Selected = True
        .Item(1).Text = "(Widown系统)"
        TreeView1_NodeClick .Item(1)
    End With
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Set Image1.Picture = LoadPicture("")
    Me.Repaint
    hdc = GetDC(0)
    hPen = CreatePen(0, 2, vbBlack)
    hPenPrev = SelectObject(hdc, hPen)
    SetROP2 hdc, R2_NOTXORPEN
    Timer1.Interval = 400
    Timer1.Enabled = True
    lTimeButtonDown = GetTickCount
    GetCursorPos pStart
End Sub


Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then
        SetCursor hCursor
    End If
End Sub

Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim lTimeButtonUp As Long
    Dim pEnd As POINTAPI
    lTimeButtonUp = GetTickCount
    GetCursorPos pEnd
    
    Timer1.Enabled = False
    If bFrameVisible Then
        DrawBorder hwnd
    End If
    SelectObject hdc, hPenPrev
    ReleaseDC 0, hdc
    DeleteObject hPen
    Set Image1.Picture = oMouseIcon
    Me.Repaint
    
 
    If (lTimeButtonUp - lTimeButtonDown < 1000) And (Abs(pEnd.X - pStart.X) < 10) And (Abs(pEnd.Y - pStart.Y) < 10) Then Exit Sub

    Dim w As window
    Dim oNode As Node
    Set w = New window
    w.hwnd = hwnd
    ClearTreeView TreeView1
    With TreeView1.Nodes
        Set oNode = .Add(, , "hwnd" & w.hwnd, w.ClassName & "(" & w.Caption & ")")
        Set oNode.Tag = w
        ListChildWindow w, oNode
        .Item(1).Expanded = True
        .Item(1).Selected = True
        TreeView1_NodeClick .Item(1)
    End With

End Sub

Private Sub ClearTreeView(tvw As TreeView)
    Dim oNode As Node
    For Each oNode In tvw.Nodes
        Set oNode.Tag = Nothing
    Next
    tvw.Nodes.Clear
End Sub

Private Sub ListChildWindow(w As window, oNode As Node)
    bEnableEvents = False
    Dim childWin As window
    Dim childNode As Node
    For Each childWin In w.Children
        Set childNode = TreeView1.Nodes.Add(oNode.Key, tvwChild, "hwnd" & childWin.hwnd, childWin.ClassName & "(" & childWin.Caption & ")")
        Set childNode.Tag = childWin
        If childWin.Children.Count > 0 Then
            ListChildWindow childWin, childNode
        End If
    Next
    bEnableEvents = True
End Sub


Private Sub imgBigIcon_Click()
    Dim oPic As StdPicture
    If Not oWin Is Nothing Then
        Set oPic = GetIconFromFile
        If Not oPic Is Nothing Then
            imgBigIcon.Picture = oPic
            oWin.BigIcon = oPic
            Frame4.Repaint
        End If
    End If
End Sub


Private Sub imgSmallIcon_Click()
    Dim oPic As StdPicture
    If Not oWin Is Nothing Then
        Set oPic = GetIconFromFile
        If Not oPic Is Nothing Then
            imgSmallIcon.Picture = oPic
            oWin.SmallIcon = oPic
            Frame4.Repaint
        End If
    End If

End Sub
 
 

Private Sub lstStyle_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub optPic_Click()
    If Not oWin Is Nothing Then
        If optPic Then
            imgWindowPic.Picture = oWin.Picture
        Else
            imgWindowPic.Picture = oWin.ScreenPicture
        End If
        MultiPage1.Pages(1).Repaint
    End If
End Sub

Private Sub optScreenPic_Click()
    
    If Not oWin Is Nothing Then
        If optPic Then
            imgWindowPic.Picture = oWin.Picture
        Else
            imgWindowPic.Picture = oWin.ScreenPicture
        End If
        MultiPage1.Pages(1).Repaint
    End If
End Sub


 

Private Sub spZoom_Change()
    txtZoom.Text = spZoom.Value & "%"
    Frame1.Zoom = spZoom.Value
 
End Sub

Private Sub Timer1_Timer()

    Dim hwndNew As Long
    Dim p As POINTAPI
    GetCursorPos p
    hwndNew = WindowFromPoint(p.X, p.Y)
    If hwndNew <> hwnd Then
        If bFrameVisible Then
            DrawBorder hwnd
        End If
        hwnd = hwndNew
    End If
    DrawBorder hwnd

End Sub



Private Sub DrawBorder(hwnd As Long)

    Dim oRect As RECT
    GetWindowRect hwnd, oRect
    If oRect.Left < 2 Then oRect.Left = 2
    If oRect.Right > GetSystemMetrics(SM_CXSCREEN) - 2 Then oRect.Right = GetSystemMetrics(SM_CXSCREEN) - 2
    If oRect.Top < 2 Then oRect.Top = 2
    If oRect.Top > GetSystemMetrics(SM_CYSCREEN) - 2 Then oRect.Top = GetSystemMetrics(SM_CYSCREEN) - 2
    

    Rectangle hdc, oRect.Left, oRect.Top, oRect.Right + 2, oRect.Bottom + 2
    bFrameVisible = Not bFrameVisible
    Set w = Nothing
End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lStyle As Long
    Dim i As Integer
    Dim lStyleOther As Long
    bEnableEvents = False
    Set oWin = Node.Tag
    
    With oWin
        txtCaption = .Caption
        txtText = .Text
        txtClassName = .ClassName
        txtHwnd = .hwnd
        txtParentHwnd = .Parent.hwnd
        txtPID = .PID
        txtThreadID = .ThreadID
        txtScreenPos = "(" & .ScreenLeft & ", " & .ScreenTop & ")"
        txtWindowPos = "(" & .WindowLeft & ", " & .WindowTop & ")"
        txtLeft = .Left
        txtTop = .Top
        txtWidth = .Width
        txtHeight = .Height

        chkVisible = .Visible
        chkEnabled = .Enabled
        chkTopMost = .TopMost
        chkTasckBar = .ShowInTaskBar
        Set imgBigIcon.Picture = .BigIcon
        Set imgSmallIcon.Picture = .SmallIcon
        Frame4.Repaint
        
        lStyle = oWin.Style
        lOldStyle = lStyle
        For i = 0 To lstStyle.ListCount - 2
            lstStyle.Selected(i) = CLng(lstStyle.List(i, 1)) And lStyle
            lStyleOther = lStyle And Not StyleLong(i)
        Next

        lstStyle.List(lstStyle.ListCount - 1, 1) = "&H" & Hex(lStyleOther)
        lstStyle.Selected(lstStyle.ListCount - 1) = True
    End With
    If optPic Then
        imgWindowPic.Picture = oWin.Picture
    Else
        imgWindowPic.Picture = oWin.ScreenPicture
    End If
    Frame1.ScrollHeight = imgWindowPic.Height
    Frame1.ScrollWidth = imgWindowPic.Width
    bEnableEvents = True
    Me.Repaint
    
End Sub


Private Sub txtCaption_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not oWin Is Nothing Then
        oWin.Caption = txtCaption
    End If
End Sub

Private Sub txtCaption_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not oWin Is Nothing Then
        If KeyCode = 13 Then
            oWin.Caption = txtCaption
        ElseIf KeyCode = 27 Then
            txtCaption = oWin.Caption
        End If
    End If
End Sub

 

Private Sub txtText_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not oWin Is Nothing Then
        oWin.Text = txtText
    End If
End Sub

Private Sub txtText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not oWin Is Nothing Then
        If KeyCode = 13 Then
            oWin.Text = txtText
        ElseIf KeyCode = 27 Then
            txtCText = oWin.Text
        End If
    End If
End Sub

Private Sub txtLeft_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    MoveWindow
End Sub


Private Sub txtLeft_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        MoveWindow
        KeyCode = 0
    ElseIf KeyCode = 27 Then
        If Not oWin Is Nothing Then
            txtLeft = oWin.Left
        End If
    End If
End Sub

Private Sub txtText_Change()

End Sub


Private Sub txtTop_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    MoveWindow
End Sub



Private Sub txtTop_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        MoveWindow
        KeyCode = 0
    ElseIf KeyCode = 27 Then
        If Not oWin Is Nothing Then
            txtTop = oWin.Top
        End If
    End If
End Sub

Private Sub txtWidth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    MoveWindow
End Sub

Private Sub txtWidth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        MoveWindow
        KeyCode = 0
    ElseIf KeyCode = 27 Then
        If Not oWin Is Nothing Then
            txtWidth = oWin.Width
        End If
    End If
End Sub

Private Sub txtHeight_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    MoveWindow
End Sub

Private Sub txtHeight_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        MoveWindow
        KeyCode = 0
    ElseIf KeyCode = 27 Then
        If Not oWin Is Nothing Then
            txtHeight = oWin.Height
        End If
    End If
End Sub

Private Sub MoveWindow()
    If Not oWin Is Nothing Then

        If IsNumeric(txtLeft) Then
            oWin.Left = txtLeft
        Else
            txtLeft = oWin.Left
        End If

    
        If IsNumeric(txtTop) Then
            oWin.Top = txtTop
        Else
            txtTop = oWin.Top
        End If

    
        If IsNumeric(txtWidth) Then
            oWin.Width = txtWidth
        Else
            txtWidth = oWin.Width
        End If
        
        If IsNumeric(txtHeight) Then
            oWin.Height = txtHeight
        Else
            txtHeight = oWin.Height
        End If
    End If
End Sub


 

Private Sub txtZoom_Change()

End Sub

Private Sub UserForm_Initialize()

    Me.Show
    Set Timer1 = New xTimer
    hCursor = LoadCursor(ByVal 0&, IDC_ARROW)
    Set oMouseIcon = Image1.Picture
    StyleName = Array("WS_BORDER", _
                        "WS_CAPTION", _
                        "WS_CHILD(WS_CHILDWINDOW)", _
                        "WS_CLIPCHILDREN", _
                        "WS_CLIPSIBLINGS", _
                        "WS_DISABLED", _
                        "WS_DLGFRAME", _
                        "WS_GROUP", _
                        "WS_HSCROLL", _
                        "WS_MAXIMIZE", _
                        "WS_MAXIMIZEBOX(WS_TABSTOP)", _
                        "WS_MINIMIZE(WS_ICONIC)", _
                        "WS_MINIMIZEBOX", _
                        "WS_OVERLAPPED(WS_TILED)", _
                        "WS_POPUP", _
                        "WS_SYSMENU", _
                        "WS_THICKFRAME(WS_SIZEBOX)", _
                        "WS_VISIBLE", _
                        "WS_VSCROLL", _
                        "未知样式")
    StyleLong = Array(&H800000, _
                        &HC00000, _
                        &H40000000, _
                        &H2000000, _
                        &H4000000, _
                        &H8000000, _
                        &H400000, _
                        &H20000, _
                        &H100000, _
                        &H1000000, _
                        &H10000, _
                        &H20000000, _
                        &H20000, _
                        &H0, _
                        &H80000000, _
                        &H80000, _
                        &H40000, _
                        &H10000000, _
                        &H200000, _
                        &H0)
    Dim i As Integer
    For i = LBound(StyleName) To UBound(StyleName)
        lstStyle.AddItem StyleName(i)
        lstStyle.List(i, 1) = "&H" & Hex(StyleLong(i))
    Next
    txtZoom.Text = "100%"
    spZoom.Value = 100
    spZoom.Min = 10
    spZoom.Max = 400
 
    Dim cBtn As CommandBarButton
    Set cBtn = Application.CommandBars.FindControl(, 3)
    If Not cBtn Is Nothing Then
        cmdSave.Picture = cBtn.Picture
    End If
    spZoom.ZOrder 0
End Sub


Private Sub UserForm_Terminate()
    Set Timer1 = Nothing
    DestroyCursor hCursor
    ClearTreeView TreeView1
End Sub


