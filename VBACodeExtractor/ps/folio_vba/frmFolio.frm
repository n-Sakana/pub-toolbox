Attribute VB_Name = "frmFolio"
Attribute VB_Base = "0{29432CC7-C816-421A-A1BF-938B2AAFA171}{BFD3D00F-EA29-4EFF-A978-7DA8E8D12387}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' ============================================================================
' Controls (WithEvents for event handling)
' ============================================================================
Private WithEvents m_cmbSource As MSForms.ComboBox
Attribute m_cmbSource.VB_VarHelpID = -1
Private WithEvents m_txtFilter As MSForms.TextBox
Attribute m_txtFilter.VB_VarHelpID = -1
Private WithEvents m_lstRecords As MSForms.ListBox
Attribute m_lstRecords.VB_VarHelpID = -1
Private WithEvents m_mpgTabs As MSForms.MultiPage
Attribute m_mpgTabs.VB_VarHelpID = -1
Private WithEvents m_cmdSettings As MSForms.CommandButton
Attribute m_cmdSettings.VB_VarHelpID = -1
Private WithEvents m_cmdLogClear As MSForms.CommandButton
Attribute m_cmdLogClear.VB_VarHelpID = -1
Private WithEvents m_cmdToggleLeft As MSForms.CommandButton
Attribute m_cmdToggleLeft.VB_VarHelpID = -1
Private WithEvents m_cmdToggleRight As MSForms.CommandButton
Attribute m_cmdToggleRight.VB_VarHelpID = -1
Private WithEvents m_resizeHandle As MSForms.label
Attribute m_resizeHandle.VB_VarHelpID = -1
Private WithEvents m_splitterLeft As MSForms.label
Attribute m_splitterLeft.VB_VarHelpID = -1
Private WithEvents m_splitterRight As MSForms.label
Attribute m_splitterRight.VB_VarHelpID = -1
Private WithEvents m_lstMail As MSForms.ListBox
Attribute m_lstMail.VB_VarHelpID = -1
Private WithEvents m_lstAttach As MSForms.ListBox
Attribute m_lstAttach.VB_VarHelpID = -1
Private WithEvents m_lstFiles As MSForms.ListBox
Attribute m_lstFiles.VB_VarHelpID = -1

' Non-event controls
Private m_lblStatus As MSForms.label
Private m_lblCount As MSForms.label
Private m_lstLog As MSForms.ListBox
Private m_lblSubject As MSForms.label
Private m_lblFrom As MSForms.label
Private m_lblDate As MSForms.label
Private m_txtMailBody As MSForms.TextBox

' ============================================================================
' State
' ============================================================================
Private m_currentSource As String
Private m_currentTable As ListObject
Private m_filteredRows As Collection
Private m_currentRecIdx As Long
Private m_fieldEditors As Collection
Private m_matchedMails As Object    ' Dict from FindMailRecords
Private m_matchedMailArr As Variant ' = m_matchedMails.Items (for indexed access)
Private m_watcher As SheetWatcher
Private m_workerPending As Boolean
Private m_pendingMailFolder As String
Private m_pendingCaseRoot As String
Private m_pendingMatchField As String
Private m_pendingMatchMode As String
Private m_workerLastVersion As Long
Private m_workerReady As Boolean
Private m_fileTreeItems As Collection
Private m_undoStack As Collection
Private m_loading As Boolean
Private m_initialLoadDone As Boolean
Private m_lastWidth As Single
Private m_lastHeight As Single
Private m_fieldGroupPageCount As Long
Private m_mailPageIdx As Long
Private m_filesPageIdx As Long

Private Const M As Long = 6
Private Const UNDO_MAX As Long = 50
Private m_leftW As Single
Private m_rightW As Single
Private m_leftVisible As Boolean
Private m_rightVisible As Boolean
Private m_splitterDragging As Boolean
Private m_splitterDragStartX As Single
Private m_splitterDragStartY As Single
Private m_splitterDragTarget As String  ' "left" or "right" or "resize"
Private m_splitterGuide As MSForms.label
Private m_origWidth As Single
Private m_origHeight As Single
Private m_fontSize As Long

' Window size range
Private Const SIZE_MIN_W As Long = 730
Private Const SIZE_MAX_W As Long = 1400
Private Const SIZE_MIN_H As Long = 400
Private Const SIZE_MAX_H As Long = 900

' ============================================================================
' Initialize
' ============================================================================

Private Sub UserForm_Initialize()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "Initialize"
    On Error GoTo ErrHandler
    Set m_filteredRows = New Collection
    Set m_fieldEditors = New Collection
    Set m_matchedMails = FolioLib.NewDict()
    Set m_fileTreeItems = New Collection
    Set m_undoStack = New Collection
    m_currentRecIdx = -1

    m_leftW = FolioLib.GetLng("left_width", 250)
    m_rightW = FolioLib.GetLng("right_width", 250)
    m_leftVisible = True
    m_rightVisible = True
    m_fontSize = FolioLib.GetLng("font_size", 10)
    Me.Width = FolioLib.GetLng("window_width", 870)
    Me.Height = FolioLib.GetLng("window_height", 540)

    m_loading = True
    eh.Trace "BuildLayout"
    BuildLayout

    eh.Trace "LoadSources"
    LoadSources

    Dim selSrc As String: selSrc = FolioLib.GetStr("selected_source")
    If Len(selSrc) > 0 Then
        Dim si As Long
        For si = 0 To m_cmbSource.ListCount - 1
            If m_cmbSource.List(si) = selSrc Then m_cmbSource.ListIndex = si: Exit For
        Next si
    End If
    m_txtFilter.Text = FolioLib.GetStr("search_text")

    m_lastWidth = Me.Width: m_lastHeight = Me.Height
    m_loading = False
    RepositionControls
    ' Deferred worker startup (after UI is visible)
    Application.OnTime Now, "FolioMain.DeferredStartup"
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Layout
' ============================================================================

Private Sub BuildLayout()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildLayout"
    On Error GoTo ErrHandler
    Me.Caption = "folio"
    Me.BackColor = &HFFFFFF

    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    Dim centerW As Single: centerW = cw - m_leftW - m_rightW - M * 4

    Set m_cmbSource = AddCombo(Me, "cmbSource", M, M, m_leftW, 18)
    m_cmbSource.Style = fmStyleDropDownList
    Set m_txtFilter = AddTextBox(Me, "txtFilter", M, M + 22, m_leftW, 18)
    Set m_lstRecords = AddListBox(Me, "lstRecords", M, M + 44, m_leftW, ch - 74)
    m_lstRecords.Font.Name = "Meiryo": m_lstRecords.Font.size = m_fontSize

    Dim cx As Single: cx = m_leftW + M * 2
    Set m_cmdSettings = AddButton(Me, "cmdSettings", cx, M, 60, 22, "Settings")
    Set m_cmdToggleLeft = AddButton(Me, "cmdToggleLeft", cx + 64, M, 22, 22, ChrW$(&H25C0))
    m_cmdToggleLeft.Font.size = 7
    Set m_cmdToggleRight = AddButton(Me, "cmdToggleRight", cx + 88, M, 22, 22, ChrW$(&H25B6))
    m_cmdToggleRight.Font.size = 7

    ' Left splitter
    Set m_splitterLeft = Me.Controls.Add("Forms.Label.1", "splitterLeft")
    m_splitterLeft.BackColor = &HD0D0D0: m_splitterLeft.Caption = ""
    m_splitterLeft.SpecialEffect = fmSpecialEffectFlat
    m_splitterLeft.BorderStyle = fmBorderStyleNone
    m_splitterLeft.MousePointer = 9

    ' Right splitter
    Set m_splitterRight = Me.Controls.Add("Forms.Label.1", "splitterRight")
    m_splitterRight.BackColor = &HD0D0D0: m_splitterRight.Caption = ""
    m_splitterRight.SpecialEffect = fmSpecialEffectFlat
    m_splitterRight.BorderStyle = fmBorderStyleNone
    m_splitterRight.MousePointer = 9

    ' Splitter guide (shared, shown during drag)
    Set m_splitterGuide = Me.Controls.Add("Forms.Label.1", "splitterGuide")
    m_splitterGuide.BackColor = &H808080: m_splitterGuide.Caption = ""
    m_splitterGuide.SpecialEffect = fmSpecialEffectFlat
    m_splitterGuide.BorderStyle = fmBorderStyleNone
    m_splitterGuide.Visible = False: m_splitterGuide.Enabled = False

    Set m_mpgTabs = Me.Controls.Add("Forms.MultiPage.1", "mpgTabs")
    With m_mpgTabs
        .Left = cx: .Top = M + 26: .Width = centerW: .Height = ch - 56
        .Font.Name = "Meiryo UI": .Font.size = 9
        .BackColor = &HFFFFFF
        .Style = fmTabStyleButtons
    End With
    m_mpgTabs.Pages(0).Caption = "Detail"
    Do While m_mpgTabs.Pages.count > 1: m_mpgTabs.Pages.Remove 1: Loop

    Dim rx As Single: rx = m_leftW + centerW + M * 3
    Set m_cmdLogClear = AddButton(Me, "cmdLogClear", rx, M, 40, 18, "Clear")
    m_cmdLogClear.Font.size = 8
    Set m_lstLog = AddListBox(Me, "lstLog", rx, M + 22, m_rightW, ch - 52)
    m_lstLog.Font.Name = "Meiryo": m_lstLog.Font.size = m_fontSize

    ' Status bar: count (left) + status (right)
    Dim sbTop As Single: sbTop = ch - 20
    Set m_lblCount = AddLabel(Me, "lblCount", M, sbTop, m_leftW, 16)
    m_lblCount.BackColor = &HF0F0F0
    m_lblCount.SpecialEffect = fmSpecialEffectFlat
    m_lblCount.BorderStyle = fmBorderStyleSingle
    m_lblCount.Caption = "  0 records"
    Set m_lblStatus = AddLabel(Me, "lblStatus", m_leftW + M * 2, sbTop, cw - m_leftW - M * 2, 16)
    m_lblStatus.BackColor = &HF0F0F0
    m_lblStatus.SpecialEffect = fmSpecialEffectFlat
    m_lblStatus.BorderStyle = fmBorderStyleSingle
    m_lblStatus.Caption = "  Ready"

    LoadChangeLog

    ' Resize handle ? created last so it's on top of other Labels
    Set m_resizeHandle = Me.Controls.Add("Forms.Label.1", "resizeHandle")
    m_resizeHandle.BackColor = &HD0D0D0: m_resizeHandle.Caption = ""
    m_resizeHandle.SpecialEffect = fmSpecialEffectFlat
    m_resizeHandle.BorderStyle = fmBorderStyleNone
    m_resizeHandle.MousePointer = 8

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' === Toggle left/right columns ===

Private Sub m_cmdToggleLeft_Click()
    m_leftVisible = Not m_leftVisible
    m_cmdToggleLeft.Caption = IIf(m_leftVisible, ChrW$(&H25C0), ChrW$(&H25B6))
    RepositionControls
End Sub

Private Sub m_cmdToggleRight_Click()
    m_rightVisible = Not m_rightVisible
    m_cmdToggleRight.Caption = IIf(m_rightVisible, ChrW$(&H25B6), ChrW$(&H25C0))
    RepositionControls
End Sub

' === Splitter drag (guide line follows mouse, apply on MouseUp) ===

Private Sub m_splitterLeft_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    If Button = 1 Then StartSplitterDrag "left", X
End Sub
Private Sub m_splitterLeft_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    MoveSplitterGuide "left", X
End Sub
Private Sub m_splitterLeft_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    EndSplitterDrag "left", X
End Sub

Private Sub m_splitterRight_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    If Button = 1 Then StartSplitterDrag "right", X
End Sub
Private Sub m_splitterRight_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    MoveSplitterGuide "right", X
End Sub
Private Sub m_splitterRight_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    EndSplitterDrag "right", X
End Sub

Private Sub StartSplitterDrag(Target As String, X As Single)
    m_splitterDragging = True
    m_splitterDragStartX = X
    m_splitterDragTarget = Target
    Dim src As MSForms.label
    If Target = "left" Then Set src = m_splitterLeft Else Set src = m_splitterRight
    m_splitterGuide.Left = src.Left
    m_splitterGuide.Top = 0
    m_splitterGuide.Width = 2
    m_splitterGuide.Height = Me.InsideHeight
    m_splitterGuide.Visible = True
End Sub

Private Sub MoveSplitterGuide(Target As String, X As Single)
    If Not m_splitterDragging Then Exit Sub
    If m_splitterDragTarget <> Target Then Exit Sub
    Dim src As MSForms.label
    If Target = "left" Then Set src = m_splitterLeft Else Set src = m_splitterRight
    Dim newLeft As Single: newLeft = src.Left + (X - m_splitterDragStartX)
    If newLeft < 80 Then newLeft = 80
    If newLeft > Me.InsideWidth - 80 Then newLeft = Me.InsideWidth - 80
    m_splitterGuide.Left = newLeft
End Sub

Private Sub EndSplitterDrag(Target As String, X As Single)
    If Not m_splitterDragging Then Exit Sub
    If m_splitterDragTarget <> Target Then Exit Sub
    m_splitterDragging = False
    m_splitterGuide.Visible = False
    Dim src As MSForms.label
    If Target = "left" Then Set src = m_splitterLeft Else Set src = m_splitterRight
    Dim delta As Single: delta = X - m_splitterDragStartX
    If Abs(delta) < 2 Then Exit Sub
    If Target = "left" Then
        m_leftW = m_leftW + delta
        If m_leftW < 80 Then m_leftW = 80
        If m_leftW > Me.InsideWidth - m_rightW - 100 Then m_leftW = Me.InsideWidth - m_rightW - 100
    Else
        m_rightW = m_rightW - delta
        If m_rightW < 80 Then m_rightW = 80
        If m_rightW > Me.InsideWidth - m_leftW - 100 Then m_rightW = Me.InsideWidth - m_leftW - 100
    End If
    RepositionControls
End Sub

' === Resize handle drag ===

Private Sub m_resizeHandle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    If Button = 1 Then
        m_splitterDragging = True
        m_splitterDragTarget = "resize"
        m_splitterDragStartX = X
        m_splitterDragStartY = y
        m_origWidth = Me.Width
        m_origHeight = Me.Height
    End If
End Sub

Private Sub m_resizeHandle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    If Not m_splitterDragging Then Exit Sub
    If m_splitterDragTarget <> "resize" Then Exit Sub
    Dim newW As Single: newW = m_origWidth + (X - m_splitterDragStartX)
    Dim newH As Single: newH = m_origHeight + (y - m_splitterDragStartY)
    If newW < 500 Then newW = 500
    If newH < 300 Then newH = 300
    Me.Width = newW
    Me.Height = newH
End Sub

Private Sub m_resizeHandle_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    If Not m_splitterDragging Then Exit Sub
    If m_splitterDragTarget <> "resize" Then Exit Sub
    m_splitterDragging = False
    m_lastWidth = Me.Width: m_lastHeight = Me.Height
    RepositionControls
End Sub

Private Sub RepositionControls()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "RepositionControls"
    On Error GoTo ErrHandler
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    Dim effLeftW As Single: effLeftW = IIf(m_leftVisible, m_leftW, 0)
    Dim effRightW As Single: effRightW = IIf(m_rightVisible, m_rightW, 0)
    Dim gap As Single: gap = 6
    Dim splitterW As Single: splitterW = 2
    Dim splitterArea As Single: splitterArea = gap * 2 + splitterW
    Dim handleSize As Single: handleSize = 14
    Dim centerW As Single: centerW = cw - effLeftW - effRightW - M * 2
    If m_leftVisible Then centerW = centerW - splitterArea
    If m_rightVisible Then centerW = centerW - splitterArea
    If centerW < 60 Then Exit Sub
    Dim cx As Single: cx = effLeftW + IIf(m_leftVisible, splitterArea, 0) + M
    Dim rx As Single: rx = cx + centerW + M + IIf(m_rightVisible, splitterArea, 0)

    ' Left column (no overlap with splitter)
    Dim leftCtlW As Single: leftCtlW = m_leftW - M
    m_cmbSource.Visible = m_leftVisible
    m_txtFilter.Visible = m_leftVisible
    m_lstRecords.Visible = m_leftVisible
    If m_leftVisible Then
        m_cmbSource.Left = M: m_cmbSource.Width = leftCtlW
        m_txtFilter.Left = M: m_txtFilter.Width = leftCtlW
        m_lstRecords.Left = M: m_lstRecords.Width = leftCtlW: m_lstRecords.Height = ch - 74
    End If

    ' Left splitter (centered in gap)
    m_splitterLeft.Visible = m_leftVisible
    If m_leftVisible Then
        m_splitterLeft.Left = effLeftW + gap: m_splitterLeft.Top = 0
        m_splitterLeft.Width = splitterW: m_splitterLeft.Height = ch
    End If

    ' Toolbar buttons
    m_cmdSettings.Left = cx
    m_cmdToggleLeft.Left = cx + 64
    m_cmdToggleRight.Left = cx + 88

    ' Center (tabs)
    m_mpgTabs.Left = cx: m_mpgTabs.Top = M + 26
    m_mpgTabs.Width = centerW: m_mpgTabs.Height = ch - 56

    ' Right splitter (centered in gap)
    m_splitterRight.Visible = m_rightVisible
    If m_rightVisible Then
        m_splitterRight.Left = rx - splitterArea + gap: m_splitterRight.Top = 0
        m_splitterRight.Width = splitterW: m_splitterRight.Height = ch
    End If

    ' Right column (log ? bottom trimmed so resize handle isn't covered)
    m_lstLog.Visible = m_rightVisible
    m_cmdLogClear.Visible = m_rightVisible
    If m_rightVisible Then
        m_lstLog.Left = rx: m_lstLog.Width = m_rightW - M
        m_lstLog.Height = ch - 40 - handleSize
        m_cmdLogClear.Left = rx
    End If

    ' Status bar
    Dim sbTop As Single: sbTop = ch - 20
    m_lblCount.Visible = m_leftVisible
    If m_leftVisible Then
        m_lblCount.Left = M: m_lblCount.Top = sbTop: m_lblCount.Width = leftCtlW
    End If
    m_lblStatus.Left = cx: m_lblStatus.Top = sbTop
    m_lblStatus.Width = centerW

    ' Resize handle (bottom-right, below log ListBox)
    m_resizeHandle.Left = cw - handleSize: m_resizeHandle.Top = ch - handleSize
    m_resizeHandle.Width = handleSize: m_resizeHandle.Height = handleSize

    ResizeTabContents
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub ResizeTabContents()
    On Error Resume Next
    Dim pw As Single: pw = m_mpgTabs.Width - 16
    Dim ph As Single: ph = m_mpgTabs.Height - 36

    ' Mail page: special layout
    If m_mailPageIdx >= 0 Then
        Dim mailListH As Single: mailListH = 70
        Dim hdrH As Single: hdrH = 54
        Dim attH As Single: attH = 60
        Dim hdrTop As Single: hdrTop = mailListH + 4
        Dim bodyTop As Single: bodyTop = hdrTop + hdrH
        Dim bodyH As Single: bodyH = ph - bodyTop - attH - 20

        If Not m_lstMail Is Nothing Then
            m_lstMail.Width = pw: m_lstMail.Height = mailListH
        End If
        If Not m_lblSubject Is Nothing Then m_lblSubject.Width = pw - 16: m_lblSubject.Top = hdrTop
        If Not m_lblFrom Is Nothing Then m_lblFrom.Width = pw - 16: m_lblFrom.Top = hdrTop + 20
        If Not m_lblDate Is Nothing Then m_lblDate.Width = pw - 16: m_lblDate.Top = hdrTop + 34
        If Not m_txtMailBody Is Nothing Then
            m_txtMailBody.Width = pw: m_txtMailBody.Top = bodyTop: m_txtMailBody.Height = bodyH
        End If
        ' Attachments label + list at bottom
        Dim pgM As MSForms.Page: Set pgM = m_mpgTabs.Pages(m_mailPageIdx)
        Dim ci As Long
        For ci = 0 To pgM.Controls.count - 1
            If TypeName(pgM.Controls(ci)) = "Label" Then
                If pgM.Controls(ci).Caption Like "*Attachment*" Then
                    pgM.Controls(ci).Top = ph - attH - 16: pgM.Controls(ci).Width = pw
                End If
            End If
        Next ci
        If Not m_lstAttach Is Nothing Then
            m_lstAttach.Width = pw: m_lstAttach.Top = ph - attH: m_lstAttach.Height = attH
        End If
    End If

    ' Files page
    If m_filesPageIdx >= 0 Then
        If Not m_lstFiles Is Nothing Then
            m_lstFiles.Width = pw: m_lstFiles.Height = ph
        End If
    End If

    ' Other pages (field editors)
    Dim pi As Long
    For pi = 0 To m_mpgTabs.Pages.count - 1
        If pi = m_mailPageIdx Or pi = m_filesPageIdx Then GoTo NextPage
        Dim pg As MSForms.Page: Set pg = m_mpgTabs.Pages(pi)
        Dim fi As Long
        For fi = 0 To pg.Controls.count - 1
            Dim ctl As MSForms.Control: Set ctl = pg.Controls(fi)
            If TypeName(ctl) = "Frame" Then
                ctl.Width = pw: ctl.Height = ph
                ResizeFrameEditors ctl, pw
            End If
        Next fi
NextPage:
    Next pi
    On Error GoTo 0
End Sub

Private Sub ResizeFrameEditors(fra As MSForms.Frame, frameW As Single)
    On Error Resume Next
    Dim sbW As Single: sbW = 18
    Dim editorW As Single: editorW = frameW - 8 - sbW - 8
    Dim ci As Long
    For ci = 0 To fra.Controls.count - 1
        Dim ctl As MSForms.Control: Set ctl = fra.Controls(ci)
        If TypeName(ctl) = "TextBox" Then
            ctl.Width = editorW
        End If
    Next ci
    On Error GoTo 0
End Sub

' ============================================================================
' Control Factory Helpers
' ============================================================================

Private Function AddLabel(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.label
    Set AddLabel = container.Controls.Add("Forms.Label.1", nm)
    With AddLabel
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = "Meiryo UI": .Font.size = 9
    End With
End Function

Private Function AddTextBox(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.TextBox
    Set AddTextBox = container.Controls.Add("Forms.TextBox.1", nm)
    With AddTextBox
        .Left = l: .Top = t: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HD0D0D0
        .Font.Name = "Meiryo": .Font.size = 9
    End With
End Function

Private Function AddListBox(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.ListBox
    Set AddListBox = container.Controls.Add("Forms.ListBox.1", nm)
    With AddListBox
        .Left = l: .Top = t: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HD0D0D0
    End With
End Function

Private Function AddCombo(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.ComboBox
    Set AddCombo = container.Controls.Add("Forms.ComboBox.1", nm)
    With AddCombo
        .Left = l: .Top = t: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = "Meiryo UI": .Font.size = 9
    End With
End Function

Private Function AddButton(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set AddButton = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddButton
        .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.size = 9
    End With
End Function

' ============================================================================
' Sources
' ============================================================================

Private Sub LoadSources()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "LoadSources"
    On Error GoTo ErrHandler
    m_cmbSource.Clear
    Dim wb As Workbook: Set wb = GetDataWorkbook()
    If Not wb Is Nothing Then
        Dim names As Collection: Set names = FolioData.GetWorkbookTableNames(wb)
        Dim n As Variant
        For Each n In names
            m_cmbSource.AddItem CStr(n)
        Next n
    End If
    If m_cmbSource.ListCount > 0 Then m_cmbSource.ListIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function GetDataWorkbook() As Workbook
    Dim excelPath As String: excelPath = FolioLib.GetStr("excel_path")
    If Len(excelPath) = 0 Then Exit Function
    If Dir$(excelPath) = "" Then Exit Function

    ' Check if already open
    Dim fileName As String: fileName = Dir$(excelPath)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(fileName) Then
            Set GetDataWorkbook = wb: Exit Function
        End If
    Next wb

    ' Open it
    On Error Resume Next
    Set GetDataWorkbook = Application.Workbooks.Open(excelPath, ReadOnly:=False, UpdateLinks:=0)
    On Error GoTo 0
End Function

Private Sub SwitchSource(sourceName As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "SwitchSource"
    On Error GoTo ErrHandler
    m_loading = True

    m_currentSource = sourceName
    Set m_currentTable = FolioData.FindTable(GetDataWorkbook(), sourceName)
    If m_currentTable Is Nothing Then m_loading = False: Exit Sub

    ' Watch table sheet for immediate change detection
    If Not m_watcher Is Nothing Then m_watcher.StopWatching
    Set m_watcher = New SheetWatcher
    m_watcher.Watch m_currentTable.parent, sourceName, Me

    FolioLib.EnsureSource sourceName
    FolioLib.InitFieldSettingsFromTable sourceName, m_currentTable

    ' Mail / Case config
    Dim mailFolder As String: mailFolder = FolioLib.GetStr("mail_folder")
    Dim caseRoot As String: caseRoot = FolioLib.GetStr("case_folder_root")
    Dim mailMatchField As String: mailMatchField = FolioLib.GetSourceStr(sourceName, "mail_match_field")
    If Len(mailMatchField) = 0 Then mailMatchField = "sender_email"
    Dim mailMatchMode As String: mailMatchMode = FolioLib.GetSourceStr(sourceName, "mail_match_mode", "exact")

    ' Worker startup deferred ? starts after UI is visible
    m_workerPending = True
    m_workerLastVersion = 0
    m_workerReady = False
    m_pendingMailFolder = mailFolder
    m_pendingCaseRoot = caseRoot
    m_pendingMatchField = mailMatchField
    m_pendingMatchMode = mailMatchMode

    BuildFieldEditors
    BuildJoinedTabs
    m_loading = False
    UpdateRecordList
    LoadChangeLog
    m_initialLoadDone = True
    eh.OK: Exit Sub
ErrHandler:
    m_loading = False
    eh.Catch
End Sub

' ============================================================================
' Field Editors
' ============================================================================

Private Sub BuildFieldEditors()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildFieldEditors"
    On Error GoTo ErrHandler
    Set m_fieldEditors = New Collection
    Dim pg As MSForms.Page: Set pg = m_mpgTabs.Pages(0)
    ClearPageControls pg

    Do While m_mpgTabs.Pages.count > 1
        m_mpgTabs.Pages.Remove m_mpgTabs.Pages.count - 1
    Loop
    m_fieldGroupPageCount = 0

    If m_currentTable Is Nothing Then Exit Sub
    Dim fields As Collection: Set fields = FolioLib.GetFieldNames(m_currentSource)
    If fields.count = 0 Then Exit Sub

    Dim keyCol As String: keyCol = FolioLib.GetSourceStr(m_currentSource, "key_column")
    Dim hasGroups As Boolean: hasGroups = (CountFieldGroups(fields) >= 2)

    If Not hasGroups Then
        pg.Caption = "Detail"
        AddFieldEditorsToPage pg, fields, keyCol
    Else
        Dim groups As Object: Set groups = NewDict()
        Dim groupOrder As New Collection
        Dim i As Long
        For i = 1 To fields.count
            Dim fn As String: fn = CStr(fields(i))
            Dim g As String: g = GetFieldGroup(fn)
            If Len(g) = 0 Then g = "_other"
            If Not groups.Exists(g) Then
                groups.Add g, New Collection
                groupOrder.Add g
            End If
            Dim gc As Collection: Set gc = groups(g)
            gc.Add fn
        Next i

        Dim gi As Long
        For gi = 1 To groupOrder.count
            Dim gName As String: gName = CStr(groupOrder(gi))
            If gi = 1 Then
                pg.Caption = IIf(gName = "_other", "Other", gName)
                Set gc = groups(gName)
                AddFieldEditorsToPage pg, gc, keyCol
            Else
                m_mpgTabs.Pages.Add
                Dim newPg As MSForms.Page: Set newPg = m_mpgTabs.Pages(m_mpgTabs.Pages.count - 1)
                newPg.Caption = IIf(gName = "_other", "Other", gName)
                Set gc = groups(gName)
                AddFieldEditorsToPage newPg, gc, keyCol
            End If
        Next gi
        m_fieldGroupPageCount = groupOrder.count - 1
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub AddFieldEditorsToPage(pg As MSForms.Page, fields As Collection, keyCol As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "AddFieldEditorsToPage"
    On Error GoTo ErrHandler
    Dim pw As Single: pw = m_mpgTabs.Width - 16
    Dim ph As Single: ph = m_mpgTabs.Height - 36
    Dim fraScroll As MSForms.Frame
    Set fraScroll = pg.Controls.Add("Forms.Frame.1", "fra_" & pg.Caption)
    fraScroll.Left = 0: fraScroll.Top = 0
    fraScroll.Width = pw: fraScroll.Height = ph
    fraScroll.BorderStyle = fmBorderStyleNone
    fraScroll.ScrollBars = fmScrollBarsVertical
    fraScroll.KeepScrollBarsVisible = fmScrollBarsNone
    fraScroll.Caption = ""
    fraScroll.BackColor = &HFFFFFF

    Dim yPos As Single: yPos = 8
    Dim labelW As Single: labelW = 80
    Dim sbW As Single: sbW = 18
    Dim editorLeft As Single: editorLeft = 8
    Dim editorW As Single: editorW = pw - editorLeft - sbW - 8

    Dim i As Long
    Dim HIDE_SUFFIX As String: HIDE_SUFFIX = "_" & ChrW$(38750) & ChrW$(34920) & ChrW$(31034)
    For i = 1 To fields.count
        Dim fn As String: fn = CStr(fields(i))
        ' Skip fields ending with "_非表示"
        If Right$(fn, Len(HIDE_SUFFIX)) = HIDE_SUFFIX Then GoTo NextField
        Dim isMultiline As Boolean: isMultiline = FolioLib.GetFieldBool(m_currentSource, fn, "multiline")
        Dim isEditable As Boolean: isEditable = FolioLib.GetFieldBool(m_currentSource, fn, "editable", True)
        If fn = keyCol Then isEditable = False

        Dim lbl As MSForms.label
        Set lbl = fraScroll.Controls.Add("Forms.Label.1", "lbl_" & fn)
        lbl.Left = editorLeft: lbl.Top = yPos: lbl.Width = labelW: lbl.Height = 14
        lbl.Caption = GetFieldShortName(fn)
        lbl.ControlTipText = fn
        lbl.Font.Name = "Meiryo UI": lbl.Font.size = 8
        lbl.ForeColor = RGB(100, 100, 100)

        Dim fType As String: fType = FolioLib.GetFieldStr(m_currentSource, fn, "type", "text")
        Dim isNumber As Boolean: isNumber = (fType = "number")
        Dim txtW As Single: txtW = IIf(isNumber, 120, editorW)
        Dim rowH As Single: rowH = IIf(isMultiline, 54, 22)
        Dim txt As MSForms.TextBox
        Set txt = fraScroll.Controls.Add("Forms.TextBox.1", "txt_" & fn)
        txt.Left = editorLeft: txt.Top = yPos + 14: txt.Width = txtW: txt.Height = rowH
        txt.Font.Name = "Meiryo": txt.Font.size = m_fontSize
        txt.SpecialEffect = fmSpecialEffectFlat
        txt.BorderStyle = fmBorderStyleSingle
        txt.BorderColor = &HD0D0D0
        txt.Locked = Not isEditable
        If Not isEditable Then txt.BackColor = &HF8F8F8
        If isMultiline Then txt.Multiline = True: txt.ScrollBars = fmScrollBarsVertical: txt.WordWrap = True
        If isNumber Then txt.TextAlign = fmTextAlignRight
        ' IME mode based on field type
        Select Case fType
            Case "number", "date", "currency": txt.IMEMode = fmIMEModeDisable
            Case Else: txt.IMEMode = fmIMEModeHiragana
        End Select

        Dim editor As FieldEditor
        Set editor = New FieldEditor
        editor.Init txt, fn, Me, Not isEditable, fType
        m_fieldEditors.Add editor

        yPos = yPos + rowH + 20
NextField:
    Next i

    fraScroll.ScrollHeight = yPos + 4
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub ClearPageControls(pg As MSForms.Page)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "ClearPageControls"
    On Error GoTo ErrHandler
    Do While pg.Controls.count > 0
        pg.Controls.Remove 0
    Loop
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Joined Tabs (Mail, Files)
' ============================================================================

Private Sub BuildJoinedTabs()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildJoinedTabs"
    On Error GoTo ErrHandler
    m_mailPageIdx = -1: m_filesPageIdx = -1

    ' DEBUG: show conditions in status bar
    If Len(FolioLib.GetSourceStr(m_currentSource, "mail_link_column")) > 0 And FolioData.GetMailCount() > 0 Then
        m_mpgTabs.Pages.Add
        m_mailPageIdx = m_mpgTabs.Pages.count - 1
        Dim pgMail As MSForms.Page: Set pgMail = m_mpgTabs.Pages(m_mailPageIdx)
        pgMail.Caption = "Mail (0)"
        BuildMailPage pgMail
    End If

    If Len(FolioLib.GetSourceStr(m_currentSource, "folder_link_column")) > 0 And FolioData.GetCaseCount() > 0 Then
        m_mpgTabs.Pages.Add
        m_filesPageIdx = m_mpgTabs.Pages.count - 1
        Dim pgFiles As MSForms.Page: Set pgFiles = m_mpgTabs.Pages(m_filesPageIdx)
        pgFiles.Caption = "Files (0)"
        BuildFilesPage pgFiles
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub BuildMailPage(pg As MSForms.Page)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildMailPage"
    On Error GoTo ErrHandler
    Dim pw As Single: pw = m_mpgTabs.Width - 12
    Dim ph As Single: ph = m_mpgTabs.Height - 30
    Dim attH As Single: attH = 60
    Dim hdrH As Single: hdrH = 54
    Dim mailListH As Single: mailListH = 70

    ' Mail list (top)
    Set m_lstMail = pg.Controls.Add("Forms.ListBox.1", "lstMail")
    With m_lstMail
        .Left = 0: .Top = 0: .Width = pw: .Height = mailListH
        .Font.Name = "Meiryo": .Font.size = m_fontSize
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle: .BorderColor = &HD0D0D0
    End With

    ' Header area (subject, from, date)
    Dim hdrTop As Single: hdrTop = mailListH + 4
    Set m_lblSubject = AddLabel(pg, "lblSubject", 8, hdrTop, pw - 16, 18)
    m_lblSubject.Font.Name = "Meiryo UI": m_lblSubject.Font.size = 10: m_lblSubject.Font.Bold = True

    Set m_lblFrom = AddLabel(pg, "lblFrom", 8, hdrTop + 20, pw - 16, 14)
    m_lblFrom.Font.Name = "Meiryo UI": m_lblFrom.Font.size = 8
    m_lblFrom.ForeColor = RGB(100, 100, 100)

    Set m_lblDate = AddLabel(pg, "lblDate", 8, hdrTop + 34, pw - 16, 14)
    m_lblDate.Font.Name = "Meiryo UI": m_lblDate.Font.size = 8
    m_lblDate.ForeColor = RGB(100, 100, 100)

    ' Body (center, takes remaining space)
    Dim bodyTop As Single: bodyTop = hdrTop + hdrH
    Dim bodyH As Single: bodyH = ph - bodyTop - attH - 20
    Set m_txtMailBody = pg.Controls.Add("Forms.TextBox.1", "txtMailBody")
    With m_txtMailBody
        .Left = 0: .Top = bodyTop: .Width = pw: .Height = bodyH
        .Multiline = True: .ScrollBars = fmScrollBarsVertical
        .Locked = True: .BackColor = &HFFFFFF
        .Font.Name = "Meiryo": .Font.size = m_fontSize
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle: .BorderColor = &HE0E0E0
    End With

    ' Attachments label
    Dim attLblTop As Single: attLblTop = ph - attH - 16
    Dim lblAtt As MSForms.label
    Set lblAtt = AddLabel(pg, "lblAtt", 4, attLblTop, pw, 14)
    lblAtt.Caption = "Attachments": lblAtt.ForeColor = RGB(100, 100, 100)
    lblAtt.Font.Name = "Meiryo UI": lblAtt.Font.size = 8

    ' Attachment list (bottom)
    Set m_lstAttach = pg.Controls.Add("Forms.ListBox.1", "lstAttach")
    With m_lstAttach
        .Left = 0: .Top = ph - attH: .Width = pw: .Height = attH
        .Font.Name = "Meiryo": .Font.size = m_fontSize
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle: .BorderColor = &HD0D0D0
    End With
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub BuildFilesPage(pg As MSForms.Page)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildFilesPage"
    On Error GoTo ErrHandler
    Dim pw As Single: pw = m_mpgTabs.Width - 12
    Dim ph As Single: ph = m_mpgTabs.Height - 30
    Set m_lstFiles = pg.Controls.Add("Forms.ListBox.1", "lstFiles")
    m_lstFiles.Left = 0: m_lstFiles.Top = 0: m_lstFiles.Width = pw: m_lstFiles.Height = ph
    m_lstFiles.Font.Name = "Meiryo": m_lstFiles.Font.size = m_fontSize
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Table Direct Access
' ============================================================================

Private Function TableRowCount() As Long
    If m_currentTable Is Nothing Then Exit Function
    If m_currentTable.DataBodyRange Is Nothing Then Exit Function
    TableRowCount = m_currentTable.DataBodyRange.Rows.count
End Function

Private Function TableCellValue(rowIdx As Long, colName As String) As Variant
    On Error Resume Next
    TableCellValue = m_currentTable.DataBodyRange.Cells(rowIdx, m_currentTable.ListColumns(colName).Index).value
End Function

' ============================================================================
' Record List
' ============================================================================

Private Sub UpdateRecordList()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "UpdateRecordList"
    On Error GoTo ErrHandler
    m_lstRecords.Clear
    Set m_filteredRows = New Collection
    Dim rowCount As Long: rowCount = TableRowCount()
    If rowCount = 0 Then Exit Sub

    Dim dispCols As New Collection
    Dim keyCol As String: keyCol = FolioLib.GetSourceStr(m_currentSource, "key_column")
    Dim nameCol As String: nameCol = FolioLib.GetSourceStr(m_currentSource, "display_name_column")
    If Len(keyCol) > 0 Then dispCols.Add keyCol
    If Len(nameCol) > 0 And nameCol <> keyCol Then dispCols.Add nameCol

    Dim filterText As String: filterText = Trim$(m_txtFilter.Text)
    Dim r As Long
    For r = 1 To rowCount
        ' Text filter (simple full-text search)
        If Len(filterText) > 0 Then
            Dim allText As String: allText = ""
            Dim col As ListColumn
            For Each col In m_currentTable.ListColumns
                If Not col.Name Like "_*" Then
                    Dim v As Variant: v = m_currentTable.DataBodyRange.Cells(r, col.Index).value
                    If Not IsNull(v) And Not IsEmpty(v) Then allText = allText & " " & CStr(v)
                End If
            Next col
            If InStr(1, allText, filterText, vbTextCompare) = 0 Then GoTo NextRec
        End If
        Dim label As String: label = ""
        Dim ci As Long
        For ci = 1 To dispCols.count
            Dim cn As String: cn = CStr(dispCols(ci))
            Dim fType As String: fType = FolioLib.GetFieldStr(m_currentSource, cn, "type", "text")
            Dim cv As Variant: cv = TableCellValue(r, cn)
            If Len(label) > 0 Then label = label & " | "
            label = label & FormatFieldValue(cv, fType)
        Next ci
        m_filteredRows.Add r
        m_lstRecords.AddItem label
NextRec:
    Next r

    m_lblCount.Caption = "  " & m_filteredRows.count & IIf(Len(filterText) > 0, " / " & rowCount, "") & " records"
    If m_lstRecords.ListCount > 0 Then m_lstRecords.ListIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Detail Update
' ============================================================================

Private Sub CommitPendingEdits()
    If m_fieldEditors Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To m_fieldEditors.count
        Dim editor As FieldEditor: Set editor = m_fieldEditors(i)
        editor.CommitEdit
    Next i
End Sub

Private Sub UpdateDetail()
    CommitPendingEdits
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "UpdateDetail"
    On Error GoTo ErrHandler
    If m_currentSource = "" Then Exit Sub
    Dim idx As Long: idx = m_lstRecords.ListIndex
    If idx < 0 Or idx >= m_filteredRows.count Then
        m_currentRecIdx = -1
        ClearFieldEditors
        Exit Sub
    End If
    m_currentRecIdx = CLng(m_filteredRows(idx + 1))
    On Error Resume Next
    m_mpgTabs.value = 0
    On Error GoTo 0
    FillFieldEditors
    UpdateMailTab
    UpdateFilesTab
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub FillFieldEditors()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "FillFieldEditors"
    On Error GoTo ErrHandler
    m_loading = True
    Dim i As Long
    For i = 1 To m_fieldEditors.count
        Dim editor As FieldEditor: Set editor = m_fieldEditors(i)
        Dim val As Variant: val = TableCellValue(m_currentRecIdx, editor.FieldName)
        If IsNull(val) Or IsEmpty(val) Then val = ""
        editor.SetValue CStr(val)
    Next i
    m_loading = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub ClearFieldEditors()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "ClearFieldEditors"
    On Error GoTo ErrHandler
    m_loading = True
    Dim i As Long
    For i = 1 To m_fieldEditors.count
        Dim editor As FieldEditor: Set editor = m_fieldEditors(i)
        editor.SetValue ""
    Next i
    m_loading = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Mail Tab
' ============================================================================

Private Sub UpdateMailTab()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "UpdateMailTab"
    On Error GoTo ErrHandler
    If m_mailPageIdx < 0 Then Exit Sub
    If m_lstMail Is Nothing Then Exit Sub
    m_lstMail.Clear
    m_lblSubject.Caption = "": m_lblFrom.Caption = "": m_lblDate.Caption = ""
    m_txtMailBody.Text = "": m_lstAttach.Clear

    If m_currentRecIdx < 1 Then Exit Sub
    Dim linkCol As String: linkCol = FolioLib.GetSourceStr(m_currentSource, "mail_link_column")
    If Len(linkCol) = 0 Then Exit Sub

    Dim linkVar As Variant: linkVar = TableCellValue(m_currentRecIdx, linkCol)
    Dim linkVal As String
    If Not IsNull(linkVar) And Not IsEmpty(linkVar) Then linkVal = CStr(linkVar)
    If Len(linkVal) = 0 Then Exit Sub

    Dim mailMatchField As String: mailMatchField = FolioLib.GetSourceStr(m_currentSource, "mail_match_field")
    If Len(mailMatchField) = 0 Then mailMatchField = "sender_email"
    Dim mailMatchMode As String: mailMatchMode = FolioLib.GetSourceStr(m_currentSource, "mail_match_mode", "exact")
    Set m_matchedMails = FolioData.FindMailRecords(linkVal, mailMatchField, mailMatchMode)
    If m_matchedMails.count > 0 Then m_matchedMailArr = m_matchedMails.items

    Dim i As Long
    For i = 0 To m_matchedMails.count - 1
        Dim mr As Object: Set mr = m_matchedMailArr(i)
        Dim line As String
        line = DictStr(mr, "subject") & "  -  " & DictStr(mr, "sender_email") & "  " & DictStr(mr, "received_at")
        m_lstMail.AddItem line
    Next i

    m_mpgTabs.Pages(m_mailPageIdx).Caption = "Mail (" & m_matchedMails.count & ")"
    If m_lstMail.ListCount > 0 Then m_lstMail.ListIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Files Tab
' ============================================================================

Private Sub UpdateFilesTab()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "UpdateFilesTab"
    On Error GoTo ErrHandler
    If m_filesPageIdx < 0 Then Exit Sub
    If m_lstFiles Is Nothing Then Exit Sub
    m_lstFiles.Clear
    Set m_fileTreeItems = New Collection

    If m_currentRecIdx < 1 Then m_mpgTabs.Pages(m_filesPageIdx).Caption = "Files (0)": Exit Sub
    Dim linkCol As String: linkCol = FolioLib.GetSourceStr(m_currentSource, "folder_link_column")
    If Len(linkCol) = 0 Then Exit Sub

    Dim linkVar As Variant: linkVar = TableCellValue(m_currentRecIdx, linkCol)
    Dim linkVal As String
    If Not IsNull(linkVar) And Not IsEmpty(linkVar) Then linkVal = CStr(linkVar)
    If Len(linkVal) = 0 Then Exit Sub

    ' Read from FE-side cache (populated by BE's background scan loop)
    Dim caseRoot As String: caseRoot = FolioLib.GetStr("case_folder_root")
    Dim matched As Object: Set matched = FolioData.FindCaseFiles(linkVal)

    ' Build tree: group by folder_path, show folder nodes then files
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folders As Object: Set folders = NewDict()
    Dim folderOrder As Object: Set folderOrder = NewDict()
    If matched.count > 0 Then
        Dim mItems As Variant: mItems = matched.items
        Dim i As Long
        For i = 0 To matched.count - 1
            Dim fr As Object: Set fr = mItems(i)
            Dim fp As String: fp = DictStr(fr, "folder_path")
            If Not folders.Exists(fp) Then
                folders.Add fp, NewDict()
                folderOrder.Add fp, True
            End If
            Dim fc As Object: Set fc = folders(fp)
            Set fc(DictStr(fr, "file_path")) = fr
        Next i
    End If

    ' Find case root path (e.g. C:\...\cases\R06-001)
    Dim rootPath As String
    If matched.count > 0 Then
        Dim firstItem As Object: Set firstItem = mItems(0)
        Dim cid As String: cid = DictStr(firstItem, "case_id")
        rootPath = caseRoot & "\" & cid
    End If

    ' Collect all tree nodes (folder + files) in order
    Dim nodes As New Collection
    Dim foKeys As Variant
    If folderOrder.count > 0 Then foKeys = folderOrder.keys
    Dim fi As Long
    For fi = 0 To folderOrder.count - 1
        Dim folderPath As String: folderPath = CStr(foKeys(fi))
        Dim folderName As String: folderName = fso.GetFileName(folderPath)
        Dim depth As Long: depth = 0
        If Len(rootPath) > 0 And Len(folderPath) > Len(rootPath) Then
            Dim relFolder As String: relFolder = Mid$(folderPath, Len(rootPath) + 2)
            depth = 1 + Len(relFolder) - Len(Replace(relFolder, "\", ""))
        End If

        Dim nd As Object
        Set nd = NewDict()
        nd.Add "depth", CLng(depth)
        nd.Add "type", "folder"
        nd.Add "name", folderName
        nd.Add "path", folderPath
        nodes.Add nd

        Set fc = folders(folderPath)
        Dim fcItems As Variant
        If fc.count > 0 Then fcItems = fc.items
        Dim j As Long
        For j = 0 To fc.count - 1
            Set fr = fcItems(j)
            Set nd = NewDict()
            nd.Add "depth", CLng(depth + 1)
            nd.Add "type", "file"
            nd.Add "name", DictStr(fr, "file_name")
            nd.Add "path", DictStr(fr, "file_path")
            nodes.Add nd
        Next j
    Next fi

    ' Render tree with box-drawing characters
    Dim fileCount As Long: fileCount = 0
    Dim k As Long
    For k = 1 To nodes.count
        Set nd = nodes(k)
        Dim d As Long: d = CLng(nd("depth"))
        Dim nm As String: nm = CStr(nd("name"))
        Dim tp As String: tp = CStr(nd("type"))
        Dim prefix As String: prefix = TreePrefix(nodes, k, d)

        If tp = "folder" Then
            m_lstFiles.AddItem prefix & "[" & nm & "]"
        Else
            m_lstFiles.AddItem prefix & nm
            fileCount = fileCount + 1
        End If

        Dim treeItem As Object: Set treeItem = NewDict()
        treeItem.Add "type", tp
        treeItem.Add "path", CStr(nd("path"))
        m_fileTreeItems.Add treeItem
    Next k

    m_mpgTabs.Pages(m_filesPageIdx).Caption = "Files (" & fileCount & ")"
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function TreePrefix(nodes As Collection, idx As Long, depth As Long) As String
    If depth = 0 Then TreePrefix = "": Exit Function

    ' Is this the last node at its depth?
    Dim isLast As Boolean: isLast = True
    Dim k As Long
    For k = idx + 1 To nodes.count
        Dim d2 As Long: d2 = CLng(nodes(k)("depth"))
        If d2 < depth Then Exit For
        If d2 = depth Then isLast = False: Exit For
    Next k

    ' Build prefix for ancestor levels
    Dim result As String: result = ""
    Dim level As Long
    For level = 1 To depth - 1
        Dim hasMore As Boolean: hasMore = False
        For k = idx + 1 To nodes.count
            d2 = CLng(nodes(k)("depth"))
            If d2 < level Then Exit For
            If d2 = level Then hasMore = True: Exit For
        Next k
        If hasMore Then
            result = result & "|   "
        Else
            result = result & "    "
        End If
    Next level

    ' Connector
    If isLast Then
        result = result & "+-- "
    Else
        result = result & "+-- "
    End If
    TreePrefix = result
End Function


' ============================================================================
' Save / Undo
' ============================================================================

Public Sub OnFieldEdited(FieldName As String, newVal As String)
    ' Called on every keystroke ? write to table only, no logging
    If m_loading Then Exit Sub
    If m_currentRecIdx > 0 Then
        FolioData.WriteTableCell m_currentTable, m_currentRecIdx, FieldName, newVal
    End If
End Sub

Public Sub OnFieldChanged(FieldName As String, oldVal As String, newVal As String, origin As String)
    ' Called once per edit session (on blur for local, on refresh for external)
    If m_loading Then Exit Sub
    Dim keyCol As String: keyCol = FolioLib.GetSourceStr(m_currentSource, "key_column")
    Dim keyVal As String
    Dim kv As Variant: kv = TableCellValue(m_currentRecIdx, keyCol)
    If Not IsNull(kv) And Not IsEmpty(kv) Then keyVal = CStr(kv)
    FolioLib.AddLogEntry m_currentSource, keyVal, FieldName, oldVal, newVal, origin
    AddLogLine m_currentSource, keyVal, FieldName, oldVal, newVal, origin
    If origin = "local" Then PushUndo m_currentSource, keyVal, FieldName, oldVal, newVal
    m_lblStatus.Caption = "  " & origin & ": " & FieldName & " @ " & Format$(Now, "hh:nn:ss")
End Sub

Private Sub PushUndo(src As String, key As String, field As String, oldVal As String, newVal As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "PushUndo"
    On Error GoTo ErrHandler
    Dim entry As Object: Set entry = NewDict()
    entry.Add "source", src
    entry.Add "key", key
    entry.Add "field", field
    entry.Add "old_value", oldVal
    entry.Add "new_value", newVal
    m_undoStack.Add entry
    Do While m_undoStack.count > UNDO_MAX: m_undoStack.Remove 1: Loop
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub InvokeUndo()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "InvokeUndo"
    On Error GoTo ErrHandler
    If m_undoStack.count = 0 Then m_lblStatus.Caption = "  Nothing to undo.": Exit Sub
    Dim entry As Object: Set entry = m_undoStack(m_undoStack.count)
    m_undoStack.Remove m_undoStack.count

    Dim src As String: src = DictStr(entry, "source")
    Dim field As String: field = DictStr(entry, "field")
    Dim oldVal As String: oldVal = DictStr(entry, "old_value")
    Dim newVal As String: newVal = DictStr(entry, "new_value")
    Dim key As String: key = DictStr(entry, "key")

    If src <> m_currentSource Then Exit Sub
    If m_currentTable Is Nothing Then Exit Sub

    ' ControlSource binds TextBox to cell, so writing to cell auto-updates TextBox
    If m_currentRecIdx > 0 Then FolioData.WriteTableCell m_currentTable, m_currentRecIdx, field, oldVal
    ' ChangeEvent in FieldEditor will log this automatically
    m_lblStatus.Caption = "  Undone: " & field
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Change Log
' ============================================================================

Private Sub LoadChangeLog()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "LoadChangeLog"
    On Error GoTo ErrHandler
    m_lstLog.Clear
    Dim entries As Collection: Set entries = FolioLib.GetRecentEntries(200)
    Dim i As Long
    For i = 1 To entries.count
        m_lstLog.AddItem FolioLib.FormatLogLine(entries(i))
    Next i
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub AddLogLine(src As String, key As String, field As String, oldVal As String, newVal As String, origin As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "AddLogLine"
    On Error GoTo ErrHandler
    Dim recName As String
    If m_currentRecIdx > 0 Then
        Dim nameCol As String: nameCol = FolioLib.GetSourceStr(m_currentSource, "display_name_column")
        If Len(nameCol) > 0 Then
            Dim nv As Variant: nv = TableCellValue(m_currentRecIdx, nameCol)
            If Not IsNull(nv) And Not IsEmpty(nv) Then recName = CStr(nv)
        End If
    End If
    Dim entry As Object: Set entry = NewDict()
    entry.Add "ts", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    entry.Add "src", src
    entry.Add "key", key
    entry.Add "name", recName
    entry.Add "field", field
    entry.Add "old", oldVal
    entry.Add "new", newVal
    entry.Add "origin", origin
    Dim line As String: line = FolioLib.FormatLogLine(entry)
    If m_lstLog.ListCount > 0 Then
        m_lstLog.AddItem line, 0
    Else
        m_lstLog.AddItem line
    End If
    m_lstLog.TopIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Table Change (via SheetWatcher)
' ============================================================================

Public Sub OnTableChanged()
    If m_loading Then Exit Sub
    On Error Resume Next
    RefreshCurrentRecord
    On Error GoTo 0
End Sub

' ============================================================================
' Worksheet_Change handler ? called when BE writes to FE's hidden sheets
' All reads are LOCAL (no cross-process). Returns fast.
' ============================================================================

Public Sub OnFolioSheetChange(sheetName As String)
    On Error Resume Next
    Select Case sheetName
        Case "_folio_signal"
            ' Version from local sheet
            Dim sigSh As Worksheet: Set sigSh = ThisWorkbook.Worksheets("_folio_signal")
            Dim ver As Long: ver = 0
            On Error Resume Next: ver = CLng(sigSh.Range("B1").value): On Error GoTo 0
            If ver > 0 And ver <> m_workerLastVersion Then
                m_workerLastVersion = ver
                LoadDataFromLocalSheets
            End If

        Case "_folio_diff"
            ' Diff log from local sheet
            LogDiffsFromSheet

        Case "_folio_files"
            ' Now handled synchronously in UpdateFilesTab; ignore async SheetChange
    End Select
    On Error GoTo 0
End Sub

Private Sub LoadDataFromLocalSheets()
    On Error Resume Next
    FolioData.LoadFromLocalSheets ThisWorkbook

    If Not m_workerReady Then
        m_workerReady = True
        If m_mailPageIdx < 0 Or m_filesPageIdx < 0 Then BuildJoinedTabs
    End If

    If Not m_lblStatus Is Nothing Then
        m_lblStatus.Caption = "  Active (v" & m_workerLastVersion & _
            " mail:" & FolioData.GetMailCount() & " cases:" & FolioData.GetCaseCount() & ")"
    End If

    If m_currentRecIdx > 0 Then
        UpdateMailTab
        UpdateFilesTab
    End If
    On Error GoTo 0
End Sub

Private Sub LogDiffsFromSheet()
    On Error GoTo DiffExit
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("_folio_diff")
    If ws.Range("A1").value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.value
    If IsEmpty(data) Then Exit Sub

    Dim diffs As New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 1))) = 0 Then GoTo NextDiff
        Dim d As Object: Set d = FolioLib.NewDict()
        d.Add "action", CStr(data(i, 1))
        d.Add "type", CStr(data(i, 2))
        d.Add "id", CStr(data(i, 3))
        d.Add "description", CStr(data(i, 4))
        diffs.Add d
NextDiff:
    Next i

    If diffs.count = 0 Then Exit Sub
    FolioLib.AddLogEntries diffs

    For i = 1 To diffs.count
        Dim de As Object: Set de = diffs(i)
        Dim action As String: action = FolioLib.DictStr(de, "action")
        Dim dtype As String: dtype = FolioLib.DictStr(de, "type")
        Dim desc As String: desc = FolioLib.DictStr(de, "description")
        Dim prefix As String
        If action = "added" Then prefix = "+" Else prefix = "-"
        Dim line As String
        line = Format$(Now, "hh:nn:ss") & "  " & prefix & dtype & "  " & desc
        If m_lstLog.ListCount > 0 Then
            m_lstLog.AddItem line, 0
        Else
            m_lstLog.AddItem line
        End If
    Next i
    m_lstLog.TopIndex = 0
DiffExit:
End Sub

' ============================================================================
' Deferred Worker Startup
' ============================================================================

Public Sub DoPollCycle()
    If Not m_workerPending Then Exit Sub
    m_workerPending = False
    On Error Resume Next
    If FolioMain.g_workerApp Is Nothing Then
        If Not m_lblStatus Is Nothing Then m_lblStatus.Caption = "  Starting worker..."
        DoEvents
        FolioMain.StartWorker m_pendingMailFolder, m_pendingCaseRoot, m_pendingMatchField, m_pendingMatchMode
        If Not m_lblStatus Is Nothing Then m_lblStatus.Caption = "  Scanning..."
    Else
        If Not m_lblStatus Is Nothing Then m_lblStatus.Caption = "  Updating config..."
        DoEvents
        FolioMain.g_workerApp.Run "FolioWorker.UpdateConfig", _
            m_pendingMailFolder, m_pendingCaseRoot, m_pendingMatchField, m_pendingMatchMode
    End If
    On Error GoTo 0
End Sub

Private Sub RefreshCurrentRecord()
    On Error Resume Next
    If m_currentTable Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To m_fieldEditors.count
        Dim editor As FieldEditor: Set editor = m_fieldEditors(i)
        Dim val As Variant: val = TableCellValue(m_currentRecIdx, editor.FieldName)
        If IsNull(val) Or IsEmpty(val) Then val = ""
        editor.RefreshIfChanged CStr(val)
    Next i
    On Error GoTo 0
End Sub


' ============================================================================
' Event Handlers
' ============================================================================

Private Sub m_cmbSource_Change()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmbSource_Change"
    On Error GoTo ErrHandler
    If m_cmbSource.ListIndex >= 0 Then SwitchSource m_cmbSource.Text
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_txtFilter_Change()
    If m_loading Then Exit Sub
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "txtFilter_Change"
    On Error GoTo ErrHandler
    UpdateRecordList
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstRecords_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstRecords_Click"
    On Error GoTo ErrHandler

    UpdateDetail
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstRecords_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstRecords_DblClick"
    On Error GoTo ErrHandler
    If m_currentTable Is Nothing Then Exit Sub
    m_currentTable.parent.Activate
    m_currentTable.Range.Select
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdSettings_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmdSettings_Click"
    On Error GoTo ErrHandler

    frmSettings.Show vbModal
    Me.Caption = "folio"
    LoadSources
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdLogClear_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmdLogClear_Click"
    On Error GoTo ErrHandler
    FolioLib.ClearLog
    m_lstLog.Clear
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstMail_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstMail_Click"
    On Error GoTo ErrHandler
    If m_lstMail Is Nothing Then Exit Sub
    Dim idx As Long: idx = m_lstMail.ListIndex
    If idx < 0 Or idx >= m_matchedMails.count Then Exit Sub
    Dim mr As Object: Set mr = m_matchedMailArr(idx)
    m_lblSubject.Caption = DictStr(mr, "subject")
    m_lblFrom.Caption = DictStr(mr, "sender_email")
    m_lblDate.Caption = DictStr(mr, "received_at")
    Dim bp As String: bp = DictStr(mr, "body_path")
    If Len(bp) > 0 And FileExists(bp) Then
        m_txtMailBody.Text = ReadTextFile(bp)
    Else
        m_txtMailBody.Text = ""
    End If
    m_lstAttach.Clear
    Dim aps As Object: Set aps = DictObj(mr, "attachment_paths")
    If Not aps Is Nothing Then
        If TypeName(aps) = "Dictionary" And aps.count > 0 Then
            Dim apKeys As Variant: apKeys = aps.keys
            Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
            Dim ai As Long
            For ai = 0 To aps.count - 1
                m_lstAttach.AddItem fso.GetFileName(CStr(apKeys(ai)))
            Next ai
        End If
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstMail_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstMail_DblClick"
    On Error GoTo ErrHandler
    If m_lstMail Is Nothing Then Exit Sub
    Dim idx As Long: idx = m_lstMail.ListIndex
    If idx < 0 Or idx >= m_matchedMails.count Then Exit Sub
    Dim mr As Object: Set mr = m_matchedMailArr(idx)
    Dim msgPath As String: msgPath = DictStr(mr, "msg_path")
    If Len(msgPath) > 0 And FileExists(msgPath) Then ThisWorkbook.FollowHyperlink msgPath
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstAttach_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstAttach_DblClick"
    On Error GoTo ErrHandler
    If m_lstAttach Is Nothing Or m_lstMail Is Nothing Then Exit Sub
    Dim mi As Long: mi = m_lstMail.ListIndex
    Dim ai As Long: ai = m_lstAttach.ListIndex
    If mi < 0 Or ai < 0 Then Exit Sub
    Dim mr As Object: Set mr = m_matchedMailArr(mi)
    Dim aps As Object: Set aps = DictObj(mr, "attachment_paths")
    If aps Is Nothing Then Exit Sub
    If TypeName(aps) = "Dictionary" And ai < aps.count Then
        Dim attKeys As Variant: attKeys = aps.keys
        ThisWorkbook.FollowHyperlink CStr(attKeys(ai))
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstFiles_DblClick"
    On Error GoTo ErrHandler
    If m_lstFiles Is Nothing Then Exit Sub
    Dim idx As Long: idx = m_lstFiles.ListIndex
    If idx < 0 Or idx + 1 > m_fileTreeItems.count Then Exit Sub
    Dim item As Object: Set item = m_fileTreeItems(idx + 1)
    Dim itemPath As String: itemPath = DictStr(item, "path")
    If Len(itemPath) = 0 Then Exit Sub
    ThisWorkbook.FollowHyperlink itemPath
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub UserForm_Layout()
    If m_loading Then Exit Sub
    If m_splitterDragging Then Exit Sub
    On Error Resume Next
    If Me.Width <> m_lastWidth Or Me.Height <> m_lastHeight Then
        m_lastWidth = Me.Width: m_lastHeight = Me.Height
        RepositionControls
    End If
    On Error GoTo 0
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    If Shift = 2 Then
        Select Case KeyCode
            Case vbKeyF: m_txtFilter.SetFocus: KeyCode = 0
            Case vbKeyZ: InvokeUndo: KeyCode = 0
        End Select
    End If
    On Error GoTo 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "QueryClose"
    On Error GoTo ErrHandler
    FolioMain.g_formLoaded = False

    If FolioMain.g_forceClose Then
        CleanupRefs
        eh.OK: Exit Sub
    End If

    FolioLib.SetLng "window_width", CLng(Me.Width)
    FolioLib.SetLng "window_height", CLng(Me.Height)
    FolioLib.SetLng "left_width", CLng(m_leftW)
    FolioLib.SetLng "right_width", CLng(m_rightW)
    FolioLib.SetLng "font_size", m_fontSize
    FolioLib.SetStr "selected_source", m_currentSource
    FolioLib.SetStr "search_text", m_txtFilter.Text
    CleanupRefs
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub CleanupRefs()
    On Error Resume Next
    CommitPendingEdits
    If Not m_watcher Is Nothing Then m_watcher.StopWatching
    Set m_watcher = Nothing
    FolioMain.StopWorker
    Set m_currentTable = Nothing
    Set m_filteredRows = Nothing
    Set m_fieldEditors = Nothing
    Set m_matchedMails = Nothing
    Set m_fileTreeItems = Nothing
    Set m_undoStack = Nothing
    On Error GoTo 0
End Sub

