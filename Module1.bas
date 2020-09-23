Attribute VB_Name = "Module1"
Option Explicit

'**********************************************************************************************************************************
' Description:
' SetParent moves a window from having one parent window to another. _
     If needed, the window itself moves so it can be "inside" its new parent. _
     The child window can also become independent by making it a child of the desktop.

' Return Value
' If an error occured, the function returns 0 (use GetLastError to get the error code). _
     If successful, the function returns a handle to the child window's former parent window.
     
' Platforms:
'    Windows 95: Supported.
'    Windows 98: Supported.
'    Windows NT: Requires Windows NT 3.1 or later.
'    Windows 2000: Supported.
'    Windows CE: Requires Windows CE 1.0 or later.
     
' Parameters:
'        hWndChild      -- The handle of the window to change the parent of.
'        hWndNewParent  -- The handle of the window to become the new parent _
                           of the child window. To make the desktop the parent, _
                           pass 0 for this parameter.
'**********************************************************************************************************************************
Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


'**********************************************************************************************************************************
' GetWindowLong & SetWindowLong Constant Definitions
'**********************************************************************************************************************************
Public Const GWL_EXSTYLE = -20
Public Const GWL_HINSTANCE = -6
Public Const GWL_HWNDPARENT = -8
Public Const GWL_ID = -12
Public Const GWL_STYLE = -16
Public Const GWL_USERDATA = -21
Public Const GWL_WNDPROC = -4
Public Const DWL_DLGPROC = 4
Public Const DWL_MSGRESULT = 0
Public Const DWL_USER = 8

'**********************************************************************************************************************************
' Description:
' GetWindowLong retrieves a 32-bit value from the information about a window. _
     This function can also read a 32-bit value from the block of extra memory given to the window, if one exists

' Return Value
' If an error occured, the function returns 0 (use GetLastError to get the error code). _
  If successful, the function returns the 32-bit value which was retrieved.

' Platforms:
'    Windows 95: Supported.
'    Windows 98: Supported.
'    Windows NT: Requires Windows NT 3.1 or later.
'    Windows 2000: Supported.
'    Windows CE: Requires Windows CE 1.0 or later.

' Parameters:
'       hwnd            - A handle to the window to retrieve a 32-bit value from.
'       nIndex          - To get a 32-bit value from the window's extra memory block, _
                          this is the zero-based offset of the byte to begin reading from. _
                          Valid values range from 0 to the size of the extra memory block in bytes minus four. _
                          To get a 32-bit value from the properties of the window, this is one of the following flags _
                          specifying which piece of information to retieve:

'                            GWL_EXSTYLE     - Retrieve the extended window styles of the window.
'                            GWL_HINSTANCE   - Retrieve a handle to the owning application's instance.
'                            GWL_HWNDPARENT  - Retrieve a handle to the parent window, if any.
'                            GWL_ID          - Retrieve the identifier of the window.
'                            GWL_USERDATA    - Retrieve the application-defined 32-bit value associated with the window.
'                            GWL_STYLE       - Retrieve the window styles of the window.
'                            GWL_WNDPROC     - Retrieve a pointer to the WindowProc hook function acting as the window's procedure.
'
'                            If the window happens to be a dialog box, this could also be one of the following flags:
'
'                            DWL_DLGPROC     - Retrieve a handle to the WindowProc hook function acting as the dialog box procedure.
'                            DWL_MSGRESULT   - Retrieve the return value of the last message processed by the dialog box.
'                            DWL_USER        - Retrieve the application-defined 32-bit value associated with the dialog box.
     
'**********************************************************************************************************************************
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


'**********************************************************************************************************************************
' Description:
' SetWindowLong sets a 32-bit value constituting the information about a window. _
       This function can also set a 32-bit value within the block of extra memory given to the window, if such a block exists.

' Return Value
' If an error occured, the function returns 0 (use GetLastError to get the error code). _
  If successful, the function returns the previous setting of whatever 32-bit value was replaced.

' Platforms:
'    Windows 95: Supported.
'    Windows 98: Supported.
'    Windows NT: Requires Windows NT 3.1 or later.
'    Windows 2000: Supported.
'    Windows CE: Requires Windows CE 1.0 or later.

' Parameters:
'        hwnd        - A handle to the window to set a 32-bit value in.
'        nIndex      - To set a 32-bit value within the window's extra memory block, _
                      this is the zero-based offset of the byte to begin writing to. _
                      Valid values range from 0 to the size of the extra memory block in bytes minus four. _
                      To set a 32-bit value of one of the properties of the window, this is one of the following _
                      flags specifying which piece of information to set:
                      
'                           GWL_EXSTYLE     - Set the extended window styles of the window. dwNewLong is the new setting.
'                           GWL_HINSTANCE   - Set which application instance is considered to own the window. dwNewLong is the new setting.
'                           GWL_HWNDPARENT  - Retrieve a handle to the parent window, if any. _
                                              dwNewLong is a handle to the instance to set as the owner.
'                           GWL_ID          - Set the identifier of the window. dwNewLong is the new identifier.
'                           GWL_USERDATA    - Set the application-defined 32-bit value associated with the window. _
                                              dwNewLong is the new value.
'                           GWL_STYLE       - Retrieve the window styles of the window.
'                           GWL_WNDPROC     - Set the WindowProc hook function to use as the window's procedure. _
                                              dwNewLong is a pointer to the hook function to set as the window procedure.
                                              
'                           If the window happens to be a dialog box, this could also be one of the following flags:
                                   
'                               DWL_DLGPROC     - Set the WindowProc hook function to use as the dialog box procedure. _
                                                  dwNewLong is a pointer to the hook function to set as the window procedure.
                                                  
'                               DWL_MSGRESULT   - Set the return value of the last message processed by the dialog box. _
                                                  dwNewLong is the new value.

'                               DWL_USER        - Set the application-defined 32-bit value associated with the dialog box. _
                                                  dwNewLong is the new value.
                                                  
'        dwNewLong   -  The 32-bit value to set as specified by nIndex.

'**********************************************************************************************************************************
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


'**********************************************************************************************************************************
' Base Window Styles
' The following window styles are shared by all windows, regardless of their class. They generally describe the window's general _
' appearance, although many of the styles apply best to non-control windows (particularly overlapped windows).
'**********************************************************************************************************************************
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_ICONIC = &H20000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0
Public Const WS_OVERLAPPEDWINDOW = &HCF0000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = &H80880000
Public Const WS_SIZEBOX = &H40000
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_TILED = &H0
Public Const WS_TILEDWINDOW = &HCF0000
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000

Public Sub Main()
    Dim lStyle   As Long

    Load DlgLoginParent
    Load DlgChildLogin
    
    With DlgChildLogin
        '*** Get the style for the DlgChildLogin form
        lStyle = GetWindowLong(.hwnd, GWL_STYLE)
        
        '*** BitWise apply the WS_CHILD and WS_POPUP Style Masks to the DlgChildLogin form
        '    This will allow the user to freely move the child form on the parent
        '    form without it jumping around, and also be enabled.
        '
        ' Note: I've seen a lot of simple examples of doing this, but only using the SetParent API call.
        ' if you only use SetParent, it will make your Child form, jump around when attempting to move
        ' it.
        lStyle = lStyle Or WS_CHILD Or WS_POPUP
        
        '*** Set the Style Attributes to the DlgChildLogin form
        Call SetWindowLong(.hwnd, GWL_STYLE, lStyle)
        
        '*** Make the DlgChildLogin form a Child of the DlgLoginParent form
        Call SetParent(.hwnd, DlgLoginParent.hwnd)
        
        '*** Set the Parent Style Attribute to the DlgLognParent form
        Call SetWindowLong(.hwnd, GWL_HWNDPARENT, DlgLoginParent.hwnd)
        
    End With

    '*** Show the Child and Parent Forms
    DlgChildLogin.Show
    DlgLoginParent.Show
    
    '*** Set the focus on the Child Form
    DlgChildLogin.SetFocus
    
End Sub
