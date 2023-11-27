VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfMain
' Author    : beededea
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public WithEvents aboutForm As cWidgetForm
Attribute aboutForm.VB_VarHelpID = -1
Public WithEvents sovereignForm As cWidgetForm
Attribute sovereignForm.VB_VarHelpID = -1


'---------------------------------------------------------------------------------------
' Procedure : InitAndShowAsFreeForm
' Author    :
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub InitAndShowAsFreeForm(ByVal X As Long, ByVal Y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo InitAndShowAsFreeForm_Error
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and location
    Set sovereignForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 1200, 1200)
    
    ' make the form transparent
    sovereignForm.WidgetRoot.BackColor = -1

    ' create a new sovereign widget with a name, location and width, runs Class_Initialize
    Set sovereignWidget = sovereignForm.Widgets.Add(New cwSovereign, "sovereign widget", 0, 0, 1000, 1000)
        
    ' I do not 'get' the size/position relationship between the form and the widget on the form. Seems peculiar to me.
      
    'NOTE that when you move a widget by dragging, you are moving the invisible form it is drawn upon.
    sovereignForm.Move X, Y ' position the form & display it
    sovereignForm.Load
    
    ' set the z-ordering of the main form
    Call setWindowZordering
    
    ' create a second invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 350, 675)
    
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
         
    ' create a new sovereign widget with a name and location, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 0, 0, 470, 670)

   On Error GoTo 0
   Exit Sub

InitAndShowAsFreeForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitAndShowAsFreeForm of Class Module cfMain"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : sovereignForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub sovereignForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo sovereignForm_DblClick_Error

    If LTrim$(gblPlDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gblPlDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If SHIFT_1 = True Then
        SHIFT_1 = False
        Call ShellExecute(fMain.sovereignForm.hwnd, "Open", gblPlOpenFile, vbNullString, App.Path, 1)
    Else
        Call ShellExecute(fMain.sovereignForm.hwnd, "runas", thisCommand, vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

sovereignForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sovereignForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

 
'Private Sub sovereignForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'        'If pinClicked = True Then MsgBox "X = " & x & " Y = " & y & " pinClicked = " & pinClicked
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : sovereignForm_MouseMove
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sovereignForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo sovereignForm_MouseMove_Error

    If sovereignWidget.Locked = True Then Exit Sub
    If gblPlIgnoreMouse = "1" Then Exit Sub
    
    Static x0 As Single
    Static y0 As Single
    If Button Then sovereignForm.Move sovereignForm.Left + X - x0, sovereignForm.Top + Y - y0 Else x0 = X: y0 = Y

    On Error GoTo 0
    Exit Sub

sovereignForm_MouseMove_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sovereignForm_MouseMove of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : sovereignForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub sovereignForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo sovereignForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift)

    On Error GoTo 0
    Exit Sub

sovereignForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sovereignForm_KeyDown of Class Module cfMain"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : sovereignForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the sovereign/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub sovereignForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo sovereignForm_MouseUp_Error
    
    If sovereignWidget.Locked = True Then Exit Sub
    
    ' no point in saving when prefs are visible.
    If sovereignPrefs.IsVisible = False Then ' checks without activating the form - important!.
       gblPlMaximiseFormX = Str$(sovereignForm.Left) ' saving in pixels
       gblPlMaximiseFormY = Str$(sovereignForm.Top)
       
       sPutINISetting sovereignSoftwareLocation, "maximiseFormX", gblPlMaximiseFormX, gblPlSettingsFile
       sPutINISetting sovereignSoftwareLocation, "maximiseFormY", gblPlMaximiseFormY, gblPlSettingsFile
    End If
    
    'Static x0, y0: If Button Then sovereignForm.Move sovereignForm.Left + X - x0, sovereignForm.Top + Y - y0 Else x0 = X: y0 = Y

    On Error GoTo 0
    Exit Sub

sovereignForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sovereignForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub
