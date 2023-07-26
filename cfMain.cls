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
Public WithEvents BodyForm As cWidgetForm
Attribute BodyForm.VB_VarHelpID = -1


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
    Set BodyForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 1200, 1200)
    
    ' make the form transparent
    BodyForm.WidgetRoot.BackColor = -1

    ' create a new Body widget with a name, location and width, runs Class_Initialize
    Set BodyWidget = BodyForm.Widgets.Add(New cwBody, "Body", 0, 0, 1000, 1000)
        
    ' I do not 'get' the size/position relationship between the form and the widget on the form. Seems peculiar to me.
      
    'NOTE that when you move a widget by dragging, you are moving the invisible form it is drawn upon.
    BodyForm.Move X, Y ' position the form & display it
    BodyForm.Load
    
    ' set the z-ordering of the main form
    Call setWindowZordering
    
    ' create a second invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 750, 675)
    
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
         
    ' create a new Body widget with a name and location, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 0, 0, 870, 870)

   On Error GoTo 0
   Exit Sub

InitAndShowAsFreeForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitAndShowAsFreeForm of Class Module cfMain"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : BodyForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub BodyForm_DblClick()
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    On Error GoTo BodyForm_DblClick_Error

    If LTrim$(PzEDblClickCommand) = vbNullString Then Exit Sub
    
    fileToPlay = "ting.wav"
    If PzEEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    If SHIFT_1 = True Then
        SHIFT_1 = False
        Call ShellExecute(fMain.BodyForm.hwnd, "Open", PzEOpenFile, vbNullString, App.Path, 1)
    Else
        answer = MsgBox("Are you sure you want to kill this process? " & vbCr & "    " & PzEDblClickCommand & vbCr & " Do you wish to proceed? ", vbExclamation + vbYesNo)
        If answer = vbYes Then
             checkAndKill PzEDblClickCommand, False, False
            'Call ShellExecute(fMain.BodyForm.hwnd, "Open", PzEDblClickCommand, vbNullString, App.Path, 1)
        End If
    End If

    On Error GoTo 0
    Exit Sub

BodyForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BodyForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

 
'Private Sub BodyForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'        'If pinClicked = True Then MsgBox "X = " & x & " Y = " & y & " pinClicked = " & pinClicked
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : BodyForm_MouseMove
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub BodyForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo BodyForm_MouseMove_Error

    If BodyWidget.Locked = True Then Exit Sub
    If PzEIgnoreMouse = "1" Then Exit Sub
    
    Static x0 As Single
    Static y0 As Single
    If Button Then BodyForm.Move BodyForm.Left + X - x0, BodyForm.Top + Y - y0 Else x0 = X: y0 = Y

    On Error GoTo 0
    Exit Sub

BodyForm_MouseMove_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BodyForm_MouseMove of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : BodyForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub BodyForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo BodyForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift)

    On Error GoTo 0
    Exit Sub

BodyForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BodyForm_KeyDown of Class Module cfMain"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : BodyForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the Body/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub BodyForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo BodyForm_MouseUp_Error
    
    If BodyWidget.Locked = True Then Exit Sub
    
    ' no point in saving when prefs are visible.
    If panzerPrefs.IsVisible = False Then ' checks without activating the form - important!.
       PzEMaximiseFormX = Str$(BodyForm.Left) ' saving in pixels
       PzEMaximiseFormY = Str$(BodyForm.Top)
       
       sPutINISetting "Software\PzKill", "maximiseFormX", PzEMaximiseFormX, PzESettingsFile
       sPutINISetting "Software\PzKill", "maximiseFormY", PzEMaximiseFormY, PzESettingsFile
    End If
    
    'Static x0, y0: If Button Then BodyForm.Move BodyForm.Left + X - x0, BodyForm.Top + Y - y0 Else x0 = X: y0 = Y

    On Error GoTo 0
    Exit Sub

BodyForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BodyForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub
