VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTextBox_ContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CTextBox_ContextMenu
'* Created    : 11-04-2021 11:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private Const mEDIT_CONTEXTMENU_NAME = "ajpiEditContextMenu"
Private Const mCUT_TAG = "CUT"
Private Const mCOPY_TAG = "COPY"
Private Const mPASTE_TAG = "PASTE"

Private m_cbrContextMenu As CommandBar
Private WithEvents m_txtTBox As msforms.TextBox
Attribute m_txtTBox.VB_VarHelpID = -1
Private WithEvents m_cbtCut As CommandBarButton
Attribute m_cbtCut.VB_VarHelpID = -1
Private WithEvents m_cbtCopy As CommandBarButton
Attribute m_cbtCopy.VB_VarHelpID = -1
Private WithEvents m_cbtPaste As CommandBarButton
Attribute m_cbtPaste.VB_VarHelpID = -1
Private m_objDataObject As DataObject
Private m_objParent As Object



Private Function m_CreateEditContextMenu() As CommandBar
    '
    ' Build Context menu controls.
    '
    Dim cbrTemp     As CommandBar
    Const CUT_MENUID = 21
    Const COPY_MENUID = 19
    Const PASTE_MENUID = 22

    Set cbrTemp = Application.CommandBars.Add(mEDIT_CONTEXTMENU_NAME, Position:=msoBarPopup)
    With cbrTemp
        With .Controls.Add(msoControlButton)
            .Caption = "Cu&t"
            .FaceId = CUT_MENUID
            .Tag = mCUT_TAG
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "&Copy"
            .FaceId = COPY_MENUID
            .Tag = mCOPY_TAG
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "&Paste"
            .FaceId = PASTE_MENUID
            .Tag = mPASTE_TAG
        End With
    End With

    Set m_CreateEditContextMenu = cbrTemp

End Function
Private Sub m_DestroyEditContextMenu()
    On Error Resume Next
    Application.CommandBars(mEDIT_CONTEXTMENU_NAME).Delete
    Exit Sub
End Sub
Private Function m_GetEditContextMenu() As CommandBar

    On Error Resume Next

    Set m_GetEditContextMenu = Application.CommandBars(mEDIT_CONTEXTMENU_NAME)
    If m_GetEditContextMenu Is Nothing Then
        Set m_GetEditContextMenu = m_CreateEditContextMenu
    End If

    Exit Function

End Function

Private Function m_ActiveTextbox() As Boolean
    '
    ' Make sure this instance is connected to active control
    ' May need to drill down through container controls to
    ' reach ActiveControl object
    '
    Dim objCtl      As Object

    Set objCtl = m_objParent.ActiveControl
    Do While UCase(TypeName(objCtl)) <> "TEXTBOX"
        If UCase(TypeName(objCtl)) = "MULTIPAGE" Then
            Set objCtl = objCtl.Pages(objCtl.Value).ActiveControl
        Else
            Set objCtl = objCtl.ActiveControl
        End If
    Loop
    m_ActiveTextbox = (StrComp(objCtl.Name, m_txtTBox.Name, vbTextCompare) = 0)

ErrActivetextbox:
    Exit Function

End Function

Public Property Set Parent(RHS As Object)
    Set m_objParent = RHS
End Property

Private Sub m_UseMenu()

    Dim lngIndex    As Long

    For lngIndex = 1 To m_cbrContextMenu.Controls.Count
        Select Case m_cbrContextMenu.Controls(lngIndex).Tag
            Case mCUT_TAG
                Set m_cbtCut = m_cbrContextMenu.Controls(lngIndex)
            Case mCOPY_TAG
                Set m_cbtCopy = m_cbrContextMenu.Controls(lngIndex)
            Case mPASTE_TAG
                Set m_cbtPaste = m_cbrContextMenu.Controls(lngIndex)
        End Select
    Next

End Sub
Public Property Set TBox(RHS As msforms.TextBox)
    Set m_txtTBox = RHS
End Property


Private Sub Class_Initialize()

    Set m_objDataObject = New DataObject
    Set m_cbrContextMenu = m_GetEditContextMenu

    If Not m_cbrContextMenu Is Nothing Then
        m_UseMenu
    End If

End Sub

Private Sub Class_Terminate()

    Set m_objDataObject = Nothing
    m_DestroyEditContextMenu

End Sub


Private Sub m_cbtCopy_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    ' check active textbox is this instance of CTextBox_ContextMenu
    If m_ActiveTextbox() Then
        With m_objDataObject
            .Clear
            .SetText m_txtTBox.SelText
            .PutInClipboard
        End With
    End If

End Sub

Private Sub m_cbtCut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    ' check active textbox is this instance of CTextBox_ContextMenu
    If m_ActiveTextbox() Then
        With m_objDataObject
            .Clear
            .SetText m_txtTBox.SelText
            .PutInClipboard
            m_txtTBox.SelText = vbNullString
        End With
    End If

End Sub


Private Sub m_cbtPaste_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    ' check active textbox is this instance of CTextBox_ContextMenu
    On Error GoTo ErrPaste

    If m_ActiveTextbox() Then
        With m_objDataObject
            .GetFromClipboard
            m_txtTBox.SelText = .GetText
        End With
    End If

ErrPaste:
    Exit Sub
End Sub


Private Sub m_txtTBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 2 Then
        ' right click
        m_cbrContextMenu.ShowPopup
    End If

End Sub


