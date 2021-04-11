VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Textbox Contextual Menus"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5790
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : UserForm1
'* Created    : 11-04-2021 11:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private m_colContextMenus As Collection

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    Dim clsContextMenu As CTextBox_ContextMenu

    Set m_colContextMenus = New Collection

    Set clsContextMenu = New CTextBox_ContextMenu
    With clsContextMenu
        Set .TBox = UserForm1.TextBox1
        Set .Parent = Me
    End With
    m_colContextMenus.Add clsContextMenu, CStr(m_colContextMenus.Count + 1)

    Set clsContextMenu = New CTextBox_ContextMenu
    With clsContextMenu
        Set .TBox = UserForm1.TextBox2
        Set .Parent = Me
    End With
    m_colContextMenus.Add clsContextMenu, CStr(m_colContextMenus.Count + 1)

    Set clsContextMenu = New CTextBox_ContextMenu
    With clsContextMenu
        Set .TBox = UserForm1.TextBox3
        Set .Parent = Me
    End With
    m_colContextMenus.Add clsContextMenu, CStr(m_colContextMenus.Count + 1)

    Set clsContextMenu = New CTextBox_ContextMenu
    With clsContextMenu
        Set .TBox = UserForm1.TextBox4
        Set .Parent = Me
    End With
    m_colContextMenus.Add clsContextMenu, CStr(m_colContextMenus.Count + 1)

    Set clsContextMenu = New CTextBox_ContextMenu
    With clsContextMenu
        Set .TBox = UserForm1.TextBox5
        Set .Parent = Me
    End With
    m_colContextMenus.Add clsContextMenu, CStr(m_colContextMenus.Count + 1)

End Sub

Private Sub UserForm_Terminate()

    Do While m_colContextMenus.Count > 0
        m_colContextMenus.Remove m_colContextMenus.Count
    Loop
    Set m_colContextMenus = Nothing

End Sub


