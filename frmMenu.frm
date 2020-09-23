VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1590
      TabIndex        =   2
      Top             =   1785
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1605
      TabIndex        =   1
      Top             =   945
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1575
      TabIndex        =   0
      Top             =   330
      Width           =   1215
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

      'Menu item constants.
      Private Const SC_CLOSE       As Long = &HF060&

      'SetMenuItemInfo fMask constants.
      Private Const MIIM_STATE     As Long = &H1&
      Private Const MIIM_ID        As Long = &H2&

      'SetMenuItemInfo fState constants.
      Private Const MFS_GRAYED     As Long = &H3&
      Private Const MFS_CHECKED    As Long = &H8&

      'SendMessage constants.
      Private Const WM_NCACTIVATE  As Long = &H86

      'User-defined Types.
      Private Type MENUITEMINFO
          cbSize        As Long
          fMask         As Long
          fType         As Long
          fState        As Long
          wID           As Long
          hSubMenu      As Long
          hbmpChecked   As Long
          hbmpUnchecked As Long
          dwItemData    As Long
          dwTypeData    As String
          cch           As Long
      End Type

      'Declarations.
      Private Declare Function GetSystemMenu Lib "user32" ( _
          ByVal hwnd As Long, ByVal bRevert As Long) As Long

      Private Declare Function GetMenuItemInfo Lib "user32" Alias _
          "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
          ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

      Private Declare Function SetMenuItemInfo Lib "user32" Alias _
          "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
          ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

      Private Declare Function SendMessage Lib "user32" Alias _
          "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, lParam As Any) As Long

      'Application-specific constants and variables.
      Private Const xSC_CLOSE  As Long = -10
      Private Const SwapID     As Long = 1
      Private Const ResetID    As Long = 2

      Private hMenu  As Long
      Private MII    As MENUITEMINFO

      Private Sub Command1_Click()
          Dim Ret As Long

          Ret = SetId(SwapID)
          If Ret <> 0 Then

              If MII.fState = (MII.fState Or MFS_GRAYED) Then
                  MII.fState = MII.fState - MFS_GRAYED
              Else
                  MII.fState = (MII.fState Or MFS_GRAYED)
              End If

              MII.fMask = MIIM_STATE
              Ret = SetMenuItemInfo(hMenu, MII.wID, False, MII)
              If Ret = 0 Then
                  Ret = SetId(ResetID)
              End If

              Ret = SendMessage(Me.hwnd, WM_NCACTIVATE, True, 0)
              SetButtons
          End If
      End Sub

      Private Sub Command2_Click()
          Dim Ret As Long

          If MII.fState = (MII.fState Or MFS_CHECKED) Then
              MII.fState = MII.fState - MFS_CHECKED
          Else
              MII.fState = (MII.fState Or MFS_CHECKED)
          End If

          MII.fMask = MIIM_STATE
          Ret = SetMenuItemInfo(hMenu, MII.wID, False, MII)
          SetButtons
      End Sub

      Private Sub Command3_Click()
          Unload Me
      End Sub

      Private Function SetId(Action As Long) As Long
          Dim MenuID As Long
          Dim Ret As Long

          MenuID = MII.wID
          If MII.fState = (MII.fState Or MFS_GRAYED) Then
              If Action = SwapID Then
                  MII.wID = SC_CLOSE
              Else
                  MII.wID = xSC_CLOSE
              End If
          Else
              If Action = SwapID Then
                  MII.wID = xSC_CLOSE
              Else
                  MII.wID = SC_CLOSE
              End If
          End If

          MII.fMask = MIIM_ID
          Ret = SetMenuItemInfo(hMenu, MenuID, False, MII)
          If Ret = 0 Then
              MII.wID = MenuID
          End If
          SetId = Ret
      End Function

      Private Sub SetButtons()
          If MII.fState = (MII.fState Or MFS_GRAYED) Then
              Command1.Caption = "Enable"
          Else
              Command1.Caption = "Disable"
          End If
          If MII.fState = (MII.fState Or MFS_CHECKED) Then
              Command2.Caption = "Uncheck"
          Else
              Command2.Caption = "Check"
          End If
      End Sub

      Private Sub Form_Load()
          Dim Ret As Long

          hMenu = GetSystemMenu(Me.hwnd, 0)
          MII.cbSize = Len(MII)
          MII.dwTypeData = String(80, 0)
          MII.cch = Len(MII.dwTypeData)
          MII.fMask = MIIM_STATE
          MII.wID = SC_CLOSE
          Ret = GetMenuItemInfo(hMenu, MII.wID, False, MII)
          SetButtons
          Command3.Caption = "Exit"
      End Sub
 

