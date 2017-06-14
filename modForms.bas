Attribute VB_Name = "modForms"
Option Compare Database
Option Explicit


' Copyright (c) FMS, Inc.        www.fmsinc.com
' Licensed to owners of Total Visual SourceBook
'
' Module      : modForms
' Description : General routines for Microsoft Access forms (Jet and ADP). Includes VBA support for 32 and 64 bit API calls.
' Source      : Total Visual SourceBook

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Public Sub ComboBoxOpen(cbo As ComboBox)
    ' Comments: Open the passed combo box control
    ' Params  : cbo       The combo box control
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    cbo.SetFocus
    cbo.Dropdown
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.ComboBoxOpen"
    Resume PROC_EXIT
End Sub

Public Sub ComboSetFirst(cbo As ComboBox)
    ' Comments: Set the first value in a combo box as the selected value
    ' Params  : cbo       The combo box control
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    ' Set the value to combo's first ItemData value
    cbo.Value = cbo.ItemData(0)
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.ComboSetFirst"
    Resume PROC_EXIT
End Sub

Public Function CurrentFormView(frm As Form) As String
    ' Comments: Retrieve the current view of the specified form, if it's open
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Returns : Open mode of the form ("" if closed)
    ' Source  : Total Visual SourceBook
    
    Dim strMode As String
    Dim intCurrentView As Integer
    
    strMode = ""
    
    On Error Resume Next
    
    intCurrentView = frm.CurrentView
    
    If Err.Number = 0 Then
        Select Case intCurrentView
            Case acCurViewDesign
                strMode = "Design View"
            Case acCurViewDatasheet
                strMode = "Datasheet View"
            Case acCurViewFormBrowse
                strMode = "Form View"
            Case acCurViewLayout
                strMode = "Layout View"
            Case acCurViewPivotChart
                strMode = "Pivot Chart View"
            Case acCurViewPivotTable
                strMode = "Pivot Table View"
        End Select
    End If
    
    On Error GoTo 0
    
    CurrentFormView = strMode
End Function

Public Sub DeSelectControl(ctl As Control)
    ' Comments: De-highlights the text in the specified control by placing the cursor at the beginning of the control.
    '           This is useful to handle situations where a user tabs into a control and Access highlights the entire control.
    '           Call this function after arriving at the control and it will de-highlight the control.
    '           Note that this function works only with textbox and combo box controls.
    ' Params  : ctl     Handle to the control (must be a text box or combobox)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    ctl.SetFocus
    ctl.SelStart = 0
    ctl.SelLength = 0
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.DeSelectControl"
    Resume PROC_EXIT
End Sub

Public Sub EnableDisableControls(frm As Form, intSection As Integer, fEnable As Boolean)
    ' Comments: Enables or disables all controls in the specified section of the specified form.
    '           If disabling, make sure the control doesn't have focus. Controls that have focus cannot be disabled.
    '           If you call this procedure for a section that contains the control with focus, all controls except that control is disabled.
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    '           intSection    Number of the section to enable/disable controls in
    '           fEnable       True to enable controls, False to disable
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim ctl As Control
    
    On Error Resume Next
    
    For Each ctl In frm.Controls
        If ctl.Section = intSection Then
            ctl.Enabled = fEnable
        End If
    Next ctl
    
    On Error GoTo PROC_ERR
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.EnableDisableControls"
    Resume PROC_EXIT
End Sub

Public Sub Form_SetCommandButtons(frm As Form, ByVal strFont As String, ByVal dblSize As Double, ByVal fTransparent As Boolean)
    ' Comments: Set command button properties on the form, including the use of the hyperlink hand while hovering
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    '           strFont       Name of font to use
    '           dblSize       Font size (if set to zero, do not modify the font size)
    '           fTransparent  True to use transparent back style, False to use normal style
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim ctl As Control
    
    For Each ctl In frm.Controls
        If ctl.ControlType = acCommandButton Then
            ' For Access 2007 or later, this sets the cursor to change to hyperlink hand when it's over the button
            ctl.CursorOnHover = acCursorOnHoverHyperlinkHand
            
            ctl.Properties("FontName") = strFont
            If dblSize <> 0 Then
                ctl.Properties("FontSize") = dblSize
            End If
            If fTransparent Then
                ctl.Properties("BackStyle") = 0
            Else
                ctl.Properties("BackStyle") = 1
            End If
        End If
    Next ctl
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.Form_SetCommandButtons"
    Resume PROC_EXIT
End Sub

Public Sub Form_SetFonts(ByRef frm As Form, ByVal strFont As String, ByVal dblSize As Double, Optional ByVal dblSizeLimit As Boolean = 0)
    ' Comments: Set the font name and size for label, text box, and tab controls
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    '           strFont       Name of font to use
    '           dblSize       Font size (if set to zero, do not modify the font size)
    '           dblSizeLimit  Ignore controls with fonts equal to or larger than this size to avoid modifying large titles (if zero, all controls are modified)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim ctl As Control
    Dim fChangeFont As Boolean
    
    ' Go through all the controls on the form
    For Each ctl In frm.Controls
        ' See if the control is a type with a font to change
        Select Case ctl.ControlType
            Case acTextBox, acLabel, acTabCtl
                fChangeFont = True
            Case Else
                fChangeFont = False
        End Select
        
        If fChangeFont Then
            If dblSizeLimit <> 0 Then
                ' Only change the control if it's smaller than the specified size limit
                fChangeFont = (ctl.Properties("FontSize") < dblSizeLimit)
            End If
            
            If fChangeFont Then
                ' Change the font name and size
                ctl.Properties("FontName") = strFont
                If dblSize <> 0 Then
                    ctl.Properties("FontSize") = dblSize
                End If
            End If
        End If
    Next ctl
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.Form_SetFonts"
    Resume PROC_EXIT
End Sub

Public Sub Form_SetSectionColors(ByRef frm As Form, Optional ByVal lngColor As Long = -2147483643)
    ' Comments: Set the form's section colors to one consistent color (defaults to the Windows System color)
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    '           lngColor      Color to apply, defaults to System Window which automatically changes with the user's Windows style selection
    ' Source  : Total Visual SourceBook
    
    Dim intSection As Integer
    Dim sec As Section
    
    ' Need to trap for error because not every section number exists
    On Error Resume Next
    
    For intSection = 0 To 4
        Set sec = frm.Section(intSection)
        If Err.Number = 0 Then
            ' Set detail and header/footer sections to the specified color
            sec.BackColor = lngColor
        End If
    Next intSection
    
    On Error GoTo PROC_ERR
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.Form_SetSectionColors"
    Resume PROC_EXIT
End Sub

Public Sub Form_SetSectionColorsIndividually(ByRef frm As Form, ByVal lngColorDetail As Long, ByVal lngColorHeader As Long, ByVal lngColorFooter As Long, _
                                                                                        ByVal lngColorPageHeader As Long, ByVal lngColorPageFooter As Long)
    ' Comments: Set the each of the form's section colors.  Use -2147483643 if you want the System Windows color
    ' Params  : frm                 Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    '           lngColorDetail      Color to apply to form detail section
    '           lngColorHeader      Color to apply to form header section
    '           lngColorFooter      Color to apply to form footer section
    '           lngColorPageHeader  Color to apply to form page header section
    '           lngColorPageFooter  Color to apply to form page footer section
    ' Source  : Total Visual SourceBook
    
    Dim intSection As Integer
    Dim sec As Section
    
    ' Need to trap for error because not every section number exists
    On Error Resume Next
    
    For intSection = 0 To 4
        Set sec = frm.Section(intSection)
        If Err.Number = 0 Then
            ' Set each section's specified color
            Select Case intSection
                Case acDetail
                    sec.BackColor = lngColorDetail
                Case acHeader
                    sec.BackColor = lngColorHeader
                Case acFooter
                    sec.BackColor = lngColorFooter
                Case acPageHeader
                    sec.BackColor = lngColorPageHeader
                Case acPageFooter
                    sec.BackColor = lngColorPageFooter
            End Select
        End If
    Next intSection
    
    On Error GoTo PROC_ERR
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.Form_SetSectionColorsIndividually"
    Resume PROC_EXIT
End Sub

Public Sub FormAddRecord(frm As Form)
    ' Comments: Adds a new record in the specified form.
    '           Many editing forms in Access applications have the need for the user to easily add records.
    '           You can accomplish this by using the navigation buttons of the form, or by using the keyboard to move to the last record and opening a new record.
    '           However, your application may need to add additional logic to the operation of adding a new record.
    '           Use this function to add a new record, handling your special needs before or after calling it.
    ' Params  : frm     Pointer to the form
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    DoCmd.GoToRecord acForm, frm.Name, acNewRec
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormAddRecord"
    Resume PROC_EXIT
End Sub

Public Function FormClose(ByVal strForm As String, Optional ByVal fSave As Boolean = True) As Boolean
    ' Comments: Close the named form without errors.
    '           This function is useful when closing a form and that form may not be open. It first checks to see if the object is open, and if it is, closes it.
    ' Params  : strForm         Form name to close
    '           fSave           True to save any changes, False to close without saving
    ' Returns : True if form isn't open or was open and successfully closed, False if form remains open
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fOK As Boolean
    
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) = 0 Then
        fOK = True
    Else
        fOK = False
        If fSave Then
            DoCmd.Close acForm, strForm, acSaveYes
        Else
            DoCmd.Close acForm, strForm, acSaveNo
        End If
        fOK = True
    End If
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormClose"
    Resume PROC_EXIT
End Function

Public Function FormCloseAll(fSave As Boolean) As Boolean
    ' Comments: Close all open forms that can be closed
    ' Params  : fSave         True to save any changes, False to close without saving
    ' Returns : True if successful, False if one or more forms could not be closed
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fOK As Boolean
    Dim astrForms() As String
    Dim intCount As Integer
    
    fOK = True
    If Forms.Count > 0 Then
        ' Get array of form names.  Must be done before closing forms because that would change the Forms collection.
        ReDim astrForms(Forms.Count - 1)
        For intCount = 0 To Forms.Count - 1
            astrForms(intCount) = Forms(intCount).Name
        Next intCount
        
        ' Close them one by one
        For intCount = 0 To UBound(astrForms)
            If SysCmd(acSysCmdGetObjectState, acForm, astrForms(intCount)) <> 0 Then
                On Error Resume Next
                If fSave Then
                    DoCmd.Close acForm, astrForms(intCount), acSaveYes
                Else
                    DoCmd.Close acForm, astrForms(intCount), acSaveNo
                End If
                If Err.Number <> 0 Then
                    ' Note if a form could not be closed
                    fOK = False
                End If
                On Error GoTo PROC_ERR
            End If
        Next intCount
    End If
    
    FormCloseAll = fOK
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormCloseAll"
    Resume PROC_EXIT
End Function

Public Function FormControlPropertiesChange(ByVal strForm As String, ByVal strProperty As String, ByVal varValue As Variant, ByVal fSave As Boolean) As Integer
    ' Comments: Open a form in design view, sets the specified  property to the specified value on all controls, and optionally saves the form.
    '           You can use this function to make mass changes to a form's properties. The fSave parameter indicates whether or not the changes should be changed.
    ' Params  : strForm             Form name to modify
    '           strProperty         Property name to set
    '           varValue            Value to set the property to
    '           fSave               True to save and close the object, or False to leave it unsaved and open in design mode
    ' Returns : Number of controls changed
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim ctl As Control
    Dim intChanged As Integer
    
    intChanged = 0
    
    ' If the form isn't already open, open it in design view
    If FormOpenDesign(strForm, True, acSaveYes) Then
        
        ' Disable error trapping in case control doesn't support requested property
        On Error Resume Next
        
        For Each ctl In Forms(strForm).Controls
            ctl.Properties(strProperty).Value = varValue
            
            ' Keep track of the number of controls changed
            If Err = 0 Then
                intChanged = intChanged + 1
            End If
        Next ctl
        
        On Error GoTo PROC_ERR
        
        If fSave Then
            ' Save the changes and close the form
            DoCmd.Close acForm, strForm, acSaveYes
        End If
    End If
    
    FormControlPropertiesChange = intChanged
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormControlPropertiesChange"
    Resume PROC_EXIT
End Function

Public Function FormControlsToArray(strForm As String, astrFormControls() As String) As Integer
    ' Comments: Populate the passed array with control names of a form. The passed array must be 0-based.
    '           The procedure will expand the array as needed, so you do not need to pre-allocate array storage before calling the procedure.
    '           If the form specified is not open, the procedure opens it in design view, fills the array and then closes the form.
    '           If the form specified is already open, this procedure does not re-open or close the form.
    ' Params  : strForm             Form name
    '           astrFormControls()  Array to controls in the form (0-based)
    ' Returns : Number of controls
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim frm As Form
    Dim intCount As Integer
    Dim intCounter As Integer
    Dim fNeedsClosing As Boolean
    
    ' If the Form isn't already open, open it in design view
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) = 0 Then
        DoCmd.OpenForm strForm, acDesign, , , , acHidden
        fNeedsClosing = True
    End If
    
    ' Set an object variable and get the control count
    Set frm = Forms(strForm)
    intCount = frm.Controls.Count
    
    ReDim astrFormControls(0 To intCount - 1)
    For intCounter = 0 To intCount - 1
        astrFormControls(intCounter) = frm(intCounter).Name
    Next intCounter
    
    If fNeedsClosing Then
        DoCmd.Close acForm, strForm, acSaveNo
    End If
    
    FormControlsToArray = intCount
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormControlsToArray"
    Resume PROC_EXIT
End Function

Public Function FormControlsToString(ByVal strForm As String, ByVal chrDelimit As String, ByRef strControls As String) As Integer
    ' Comments: Populate the passed string with control names of a form.
    '           Use the chrDelimit parameter to specify the character or characters to use as the delimiter between control names.
    '           The procedure places the resulting string in the strIn parameter.
    ' Params  : strForm           Form name
    '           chrDelimit          Character to use as delimiter between control names
    ' Sets    : strControls         Updated list of controls
    ' Returns : Number of controls
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim frm As Form
    Dim intCount As Integer
    Dim intCounter As Integer
    Dim fNeedsClosing As Boolean
    
    ' If the Form isn't already open, open it in design view
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) = 0 Then
        DoCmd.OpenForm strForm, acDesign, , , , acHidden
        fNeedsClosing = True
    End If
    
    ' Set an object variable and get the control count
    Set frm = Forms(strForm)
    intCount = frm.Controls.Count
    
    For intCounter = 0 To intCount - 1
        strControls = strControls & frm(intCounter).Name
        If intCounter < intCount - 1 Then
            strControls = strControls & chrDelimit
        End If
    Next intCounter
    
    If fNeedsClosing Then
        DoCmd.Close acForm, strForm, acSaveNo
    End If
    
    FormControlsToString = intCount
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormControlsToString"
    Resume PROC_EXIT
End Function

Public Sub FormDeleteRecord(frm As Form, fNoWarning As Boolean)
    ' Comments: Delete the current record from the specified form
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    '           fNoWarning    True to skip confirmation dialog, False otherwise
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    ' Set focus to the form
    DoCmd.SelectObject acForm, frm.Name, False
    
    ' Set SetWarnings
    If fNoWarning Then
        DoCmd.SetWarnings False
    End If
    
    ' Delete the record using menu items
    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDeleteRecord
    
    ' Restore SetWarnings
    If fNoWarning Then
        DoCmd.SetWarnings True
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormDeleteRecord"
    Resume PROC_EXIT
End Sub

Public Sub FormMoveFirst(frm As Form)
    ' Comments: Move to the first record of a form.
    '           This is useful when you are implementing your own navigation buttons, and want to duplicate the behavior of the built-in Access navigation buttons.
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    DoCmd.GoToRecord acForm, frm.Name, acFirst
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormMoveFirst"
    Resume PROC_EXIT
End Sub

Public Sub FormMoveLast(frm As Form)
    ' Comments: Move to the last record of a form
    '           This is useful when you are implementing your own navigation buttons, and want to duplicate the behavior of the built-in Access navigation buttons.
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    DoCmd.GoToRecord acForm, frm.Name, acLast
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormMoveLast"
    Resume PROC_EXIT
End Sub

Public Sub FormMoveNext(frm As Form)
    ' Comments: Move to the next record of a form
    '           This is useful when you are implementing your own navigation buttons, and want to duplicate the behavior of the built-in Access navigation buttons.
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    DoCmd.GoToRecord acForm, frm.Name, acNext
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormMoveNext"
    Resume PROC_EXIT
End Sub

Public Sub FormMovePrevious(frm As Form)
    ' Comments: Move to the previous record of a form
    '           This is useful when you are implementing your own navigation buttons, and want to duplicate the behavior of the built-in Access navigation buttons.
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    DoCmd.GoToRecord acForm, frm.Name, acPrevious
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormMovePrevious"
    Resume PROC_EXIT
End Sub

Public Sub FormMoveSizeInches(ByVal strForm As String, ByVal dblTop As Double, ByVal dblLeft As Double, ByVal dblWidth As Double, ByVal dblHeight As Double)
    ' Comments: Move and/or resizes the named form in units of inches.
    '           The form named by the strForm parameter must be open for this procedure to work.
    '           To leave a particular sizing/position value unchanged, do not specify a value when calling the procedure.
    '           For example, to resize a form, leaving its position intact, specify values for the varWidth and varHeight parameters, and zero-length strings for the varTop and varLeft parameters.
    ' Params  : strForm         Form name
    '           dblTop          Top position of form
    '           dblLeft         Left position of form
    '           dblWidth        Width of form
    '           dblHeight       Height of form
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Const clngTwipsPerInch As Long = 1440
    
    DoCmd.SelectObject acForm, strForm, False
    DoCmd.MoveSize dblLeft * clngTwipsPerInch, dblTop * clngTwipsPerInch, dblWidth * clngTwipsPerInch, dblHeight * clngTwipsPerInch
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormMoveSizeInches"
    Resume PROC_EXIT
End Sub

Public Sub FormMoveSizeTwips(ByVal strForm As String, ByVal dblTop As Double, ByVal dblLeft As Double, ByVal dblWidth As Double, ByVal dblHeight As Double)
    ' Comments: Move and/or resizes the named form in units of Twips (1 twip = 1/1400 inch).
    '           This procedure gives a greater amount of accuracy in moving and sizing forms than the related FormMoveSizeInches procedure.
    '           The form named by the strForm parameter must be open for this procedure to work.
    '           To leave a particular sizing/position value unchanged, do not specify a value when calling the procedure.
    '           For example, to resize a form, leaving its position intact, specify values for the varWidth and varHeight parameters, and zero-length strings for the varTop and varLeft parameters.
    ' Params  : strForm         Form name
    '           dblTop          Top position of form
    '           dblLeft         Left position of form
    '           dblWidth        Width of form
    '           dblHeight       Height of form
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    DoCmd.SelectObject acForm, strForm, False
    DoCmd.MoveSize dblLeft, dblTop, dblWidth, dblHeight
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormMoveSizeTwips"
    Resume PROC_EXIT
End Sub

Public Function FormNamesToArray(ByRef astrForms() As String) As Integer
    ' Comments: Loads an array with the names of all the forms in the current database
    ' Params  : astrForms      Array of report names (0-based)
    ' Returns : Number of forms
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim intTotal As Integer
    Dim intForm As Integer
    
    intTotal = CurrentProject.AllForms.Count
    ReDim astrForms(0 To intTotal - 1)
    
    For intForm = 0 To intTotal - 1
        astrForms(intForm) = CurrentProject.AllForms(intForm).Name
    Next intForm
    
    FormNamesToArray = intTotal
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormNamesToArray"
    Resume PROC_EXIT
End Function

Public Function FormNamesToString(ByRef strForms As String, Optional ByVal chrDelimit As String = ";") As Integer
    ' Comments: Populate a string with names of all forms in the current database
    ' Sets    : strForms        String of form names to populate
    ' Params  : chrDelimit      Character to use as a delimiter, defaults to ; if not assigned
    ' Returns : Number of forms
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim intTotal As Integer
    Dim intForm As Integer
    
    intTotal = CurrentProject.AllForms.Count
    
    For intForm = 0 To intTotal - 1
        strForms = strForms & CurrentProject.AllForms(intForm).Name & chrDelimit
    Next intForm
    
    FormNamesToString = intTotal
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormNamesToString"
    Resume PROC_EXIT
End Function

Public Function FormOnNewRecord(frm As Form) As Boolean
    ' Comments: Determine if the specified form is on a new record by checking the value of the NewRecord property.
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Returns : True if the form is on a new unsaved record, False otherwise
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    FormOnNewRecord = frm.NewRecord
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormOnNewRecord"
    Resume PROC_EXIT
End Function

Public Function FormOpen(strForm As String, fDialog As Boolean) As Boolean
    ' Comments: Open a form in normal or dialog mode and trap for any errors. If the form is already open, just select it.
    ' Params  : strForm       Form name
    '           fDialog       True to open as a dialog form, False for regular mode
    ' Returns : True if form was opened successfully
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    ' Assume failure
    FormOpen = False
    
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0 Then
        ' Form is already open, set focus to it
        DoCmd.SelectObject acForm, strForm, False
    Else
        ' Form is not open, open it
        If fDialog Then
            DoCmd.OpenForm strForm, WindowMode:=acDialog
        Else
            DoCmd.OpenForm strForm, WindowMode:=acWindowNormal
        End If
    End If
    
    FormOpen = True
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormOpen"
    Resume PROC_EXIT
End Function

Public Function FormOpenDesign(strForm As String, fHidden As Boolean, Optional intSave As AcCloseSave = acSaveYes) As Boolean
    ' Comments: Open a form in design mode.  If it's already open, change the current view to design view
    ' Params  : strForm           Form name to open
    '           fHidden           True to open in hidden mode, False for visible
    '           intSave           If the form is already open, whether changes are discarded (acSaveNo), saved (acSaveYes), or the user prompted (acSavePrompt). By default, they are saved without prompting.
    ' Returns : True if form is in design mode successfully, False if it could not be opened in design mode
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fOK As Boolean
    Dim fOpen As Boolean
    
    fOK = False
    fOpen = True
    
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0 Then
        If Forms(strForm).CurrentView = acCurViewDesign Then
            ' Form is already open in design mode, so simply set focus to it
            DoCmd.SelectObject acForm, strForm, False
            fOpen = False
        Else
            ' Already open but not in design mode, so switch to it below.
            ' Before doing so, close the form and save any changes
            DoCmd.Close acForm, strForm, intSave
        End If
    End If
    
    If fOpen Then
        ' Form is not open, so open it
        If fHidden Then
            DoCmd.OpenForm strForm, acViewDesign, , , , acHidden
        Else
            DoCmd.OpenForm strForm, acViewDesign, , , , acWindowNormal
        End If
    End If
    
    fOK = True
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormOpenDesign"
    Resume PROC_EXIT
End Function

Public Function FormOpenDesignAll(Optional intWindowMode As AcWindowMode = acWindowNormal) As Integer
    ' Comments: Open all the forms in design view and if visible, minimizes them.  Skips forms that are already open
    ' Params  : intWindowMode         Open the forms normally (acWindowNormal) or hidden (acHidden)
    ' Returns : Number of forms opened by this routine
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim intCounter As Integer
    Dim obj As Object
    Dim strName As String
    
    intCounter = 0
    For Each obj In CurrentProject.AllForms
        strName = obj.Name
        
        ' Don't try to open it if it's already open
        If SysCmd(acSysCmdGetObjectState, acForm, strName) = 0 Then
            ' Try to open in design mode.
            ' May fail if it's a subform is in another form that's already open, but that's okay (error 7784)
            
            On Error Resume Next
            
            DoCmd.OpenForm strName, acDesign, , , , intWindowMode
            If Err.Number = 0 Then
                DoCmd.SelectObject acForm, strName, False
                If intWindowMode = acWindowNormal Then
                    ' This makes no difference if the workspace displays tabbed documents
                    DoCmd.Minimize
                End If
                intCounter = intCounter + 1
            End If
            
            On Error GoTo PROC_ERR
            
        End If
    Next obj
    
    FormOpenDesignAll = intCounter
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormOpenDesignAll"
    Resume PROC_EXIT
End Function

Public Function FormOpenReplace(ByVal strForm As String, Optional ByVal intView As AcFormView = acNormal, Optional ByVal strFilter As String = "", _
                                                                Optional ByVal strWhere As String, Optional ByVal intDataMode As AcFormOpenDataMode = acFormPropertySettings, _
                                                                Optional ByVal intWindowMode As AcWindowMode = acWindowNormal, Optional ByVal varOpenArgs As Variant) As Boolean
    ' Comments: Open a form with all the options and trap for any errors.
    '           If the form is open, close it and open it again (useful if an argument is passed and/or running Open and Load events)
    ' Params  : strForm           Name of form to open
    '           pintView          The view in which the form will open (default is acNormal)
    '           pstrFilter        Filter query, if any
    '           pstrWhere         Where clause, if any
    '           pintDataMode      Data mode (usually acFormPropertySettings)
    '           pintWindowMode    how it's opened (usually acWindowNormal)
    ' Returns : True if form was opened successfully
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    ' Assume failure
    FormOpenReplace = False
    
    If (SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0) Then
        ' Close the form if it's already open
        DoCmd.Close acForm, strForm
    End If
    
    DoCmd.OpenForm strForm, intView, strFilter, strWhere, intDataMode, intWindowMode, varOpenArgs
    
    FormOpenReplace = True
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormOpenReplace"
    Resume PROC_EXIT
End Function

Public Function FormOpenWait(strForm As String, fDialog As Boolean, fWait As Boolean, Optional strFormHide As String = "") As Boolean
    ' Comments: Open a form in normal or dialog mode, and wait until that form is closed with an option to hide the current form.
    '           This is useful when a form opens another form and wants to be invisible until the user closes that form.
    '           If the form is already open, just select it.
    ' Params  : strForm       Form name
    '           fDialog       True to open as a dialog form, False for regular mode
    '           fWait         True to wait until the form closes
    '           strFormHide   Form name to hide, if any. Usually the name of the calling form.
    ' Returns : True if form was opened successfully
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fOK As Boolean
    
    ' Open the form
    fOK = FormOpen(strForm, fDialog)
    If fOK Then
        ' Hide the calling form, if any
        If strFormHide <> "" Then
            Forms(strFormHide).Visible = False
        End If
    
        ' Wait until the form is closed
        Call WaitUntilFormClose(strForm)
    
        If strFormHide <> "" Then
            ' Make the calling form visible again.
            Forms(strFormHide).Visible = True
        End If
        fOK = True
    End If
    
    FormOpenWait = fOK
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormOpenWait"
    Resume PROC_EXIT
End Function

Public Function FormPropertyChange(ByVal strForm As String, ByVal strProperty As String, ByVal varValue As Variant, ByVal fSave As Boolean) As Boolean
    ' Comments: Changes the value of the named property on a form.
    ' Params  : strForm           Form name
    '           strProperty       Property name
    '           varValue          New value of the property
    '           fSave             True to save the change and close the form, False to leave it open in design mode so you can make more changes
    ' Returns : True if the property is changed
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fChanged As Boolean
    
    fChanged = False
    
    ' If the Form isn't already open, open it in design view
    If FormOpenDesign(strForm, True, acSaveYes) Then
        Forms(strForm).Properties(strProperty).Value = varValue
        fChanged = True
        
        If fSave Then
            ' Save and close
            DoCmd.Close acForm, strForm, acSaveYes
        End If
    End If
    
    FormPropertyChange = fChanged
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormPropertyChange"
    Resume PROC_EXIT
End Function

Public Sub FormSaveRecord(frm As Form)
    ' Comments: Saves the current record on the specified form.
    '           This is useful if you want to add a [Save] button to your form, and call this procedure from that button to save the current record.
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    ' Verify form is in an unsaved state first, otherwise an error may be triggered
    If frm.Dirty Then
        frm.Dirty = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormSaveRecord"
    Resume PROC_EXIT
End Sub

Public Function FormsOpen() As Boolean
    ' Comments: Determine if any forms are open.
    '           This is useful when your application must perform an operation (such as closing) that requires all forms and reports to be closed.
    '           See the CloseObjectsOfType() and CloseAllOpenObjects() functions for ways to close open objects.
    ' Returns : True if one or more forms are open, false otherwise
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    FormsOpen = (Forms.Count > 0)
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormsOpen"
    Resume PROC_EXIT
End Function

Public Sub FormUndoRecord(frm As Form)
    ' Comments: Undo the changes to the current record on the specified form.
    '            This is useful if you want to put an 'Undo' command button on your form to undo edits to the current record or field.
    ' Params  : frm           Handle to the form to modify.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    frm.Undo
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormUndoRecord"
    Resume PROC_EXIT
End Sub

Public Function GetActiveForm() As String
    ' Comments: Get the name of the currently active form
    ' Returns : Name of form, or a blank string if no form is active
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    GetActiveForm = Application.Screen.ActiveForm.Name
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.GetActiveForm"
    Resume PROC_EXIT
End Function

Public Function GetControlType(ctl As Control) As String
    ' Comments: Get the control type (English name) of a form or report control
    ' Params  : ctl         Control to check
    ' Returns : Control type
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim strType As String
    
    Select Case ctl.ControlType
        Case acAttachment
            strType = "Attachment"
        Case acBoundObjectFrame
            strType = "Bound Object Frame"
        Case acCheckBox
            strType = "Check Box"
        Case acComboBox
            strType = "Combo Box"
        Case acCommandButton
            strType = "Command Button"
        Case acCustomControl
            strType = "ActiveX Control"
        Case acImage
            strType = "Image"
        Case acLabel
            strType = "Label"
        Case acLine
            strType = "Line"
        Case acListBox
            strType = "List Box"
        Case acObjectFrame
            strType = "Unbound Object Frame"
        Case acOptionButton
            strType = "Option Button"
        Case acOptionGroup
            strType = "Option Group"
        Case acPage
            strType = "Page"
        Case acPageBreak
            strType = "Page Break"
        Case acRectangle
            strType = "Rectangle"
        Case acSubform
            strType = "SubForm/SubReport"
        Case acTabCtl
            strType = "Tab"
        Case acTextBox
            strType = "Text Box"
        Case acToggleButton
            strType = "Toggle Button"
        Case 113                        ' No id for graph object
            strType = "Chart"
        Case Else
            strType = "unknown"
    End Select
    
    GetControlType = strType
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.GetControlType"
    Resume PROC_EXIT
End Function

Public Function IsFormOpen(strForm As String) As Boolean
    ' Comments: Determine if a form is open. Typically, before opening a form, you check to see if that form is open.
    ' Params  : strForm     Form name to check
    ' Returns : True if form is open, False if not
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    IsFormOpen = (SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0)
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.IsFormOpen"
    Resume PROC_EXIT
End Function

Public Function IsFormOpenInDatasheetView(strForm As String) As Boolean
    ' Comments: Determine if a form is open in datasheet view
    ' Params  : strForm     Form name to check
    ' Returns : True if the form is open in datasheet view, False if not
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fOpen As Boolean
    
    fOpen = False
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0 Then
        fOpen = (Forms(strForm).CurrentView = acCurViewDatasheet)
    End If
    
    IsFormOpenInDatasheetView = fOpen
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.IsFormOpenInDatasheetView"
    Resume PROC_EXIT
End Function

Public Function IsFormOpenInDesignView(strForm As String) As Boolean
    ' Comments: Determine if a form is open in design view.
    '           Access supports certain operations on forms only when they are open in design view. Use this procedure to verify how a form is opened before attempting such operations.
    ' Params  : strForm     Form name to check
    ' Returns : True if the form is open in design view, False if not
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fOpen As Boolean
    
    fOpen = False
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0 Then
        fOpen = (Forms(strForm).CurrentView = acCurViewDesign)
    End If
    
    IsFormOpenInDesignView = fOpen
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.IsFormOpenInDesignView"
    Resume PROC_EXIT
End Function

Public Function IsFormOpenInFormView(strForm As String) As Boolean
    ' Comments: Determine if a form is open in form (normal) view
    ' Params  : strForm     Form name to check
    ' Returns : True if the form is open in form view, False if not
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fOpen As Boolean
    
    fOpen = False
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0 Then
        fOpen = (Forms(strForm).CurrentView = acCurViewFormBrowse)
    End If
    
    IsFormOpenInFormView = fOpen
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.IsFormOpenInFormView"
    Resume PROC_EXIT
End Function

Public Function IsFormOpenInLayoutView(strForm As String) As Boolean
    ' Comments: Determine if a form is open in layout view
    ' Params  : strForm     Form name to check
    ' Returns : True if the form is open in layout view, False if not
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim fOpen As Boolean
    
    fOpen = False
    If SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0 Then
        fOpen = (Forms(strForm).CurrentView = acCurViewLayout)
    End If
    
    IsFormOpenInLayoutView = fOpen
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.IsFormOpenInLayoutView"
    Resume PROC_EXIT
End Function

Public Function IsFormOpenInMode(strForm As String, intMode As AcFormView) As Boolean
    ' Comments: Determine if a form is open in a particular view
    ' Params  : strForm         Form to check
    '           intMode         Mode to check (acViewDesign, acViewLayout, etc.)
    ' Returns : True if the form is in the specified view, False if not open or open in another mode
    ' Source  : Total Visual SourceBook
    
    Dim intView As Integer
    Dim fOpen As Boolean
    
    fOpen = False
    
    intView = -1
    
    Select Case intMode
        Case acNormal
            intView = acCurViewFormBrowse
        Case acDesign
            intView = acCurViewDesign
        Case acFormDS
            intView = acCurViewDatasheet
        Case acFormPivotChart
            intView = acCurViewPivotChart
        Case acFormPivotTable
            intView = acCurViewPivotTable
        Case acLayout
            intView = acCurViewLayout
        Case acPreview
            intView = acCurViewPreview
    End Select
    
    If intView <> -1 Then
        On Error Resume Next
        fOpen = (Forms(strForm).CurrentView = intView)
        On Error GoTo 0
    End If
    
    IsFormOpenInMode = fOpen
End Function

Public Function IsSubForm(frm As Form) As Boolean
    ' Comments: Determine if the passed form is open as a form or a subform.
    '           This can be useful when your database contains a form that is used in regular view mode as well as a subform.
    '           If your application has special logic that is affected by whether or not the form is opened as a subform, use this procedure to detect this case.
    ' Params  : frm           Handle to the form.  If called in the form's module, use Me, otherwise use Forms(FormName)
    ' Returns : True if form is a subform, false otherwise
    ' Source  : Total Visual SourceBook
    
    Const clngErrInvalidRefToProp As Long = 2452
    Dim strName As Variant
    
    On Error Resume Next
    
    ' This line triggers error 2452 (Invalid reference to Parent property) if the form is not a subform
    strName = frm.Parent.Name
    
    IsSubForm = (Err.Number <> clngErrInvalidRefToProp)
    
    On Error GoTo 0
    
End Function

Public Function ListFillForms(ctl As Control, lngID As Long, lngRow As Long, lngCol As Long, intCode As Integer) As Variant
    ' Comments: Provides a list fill function for a list/combo box for a list of forms.
    '           Use the name of this function as the "RowSourceType" property of a listbox or combo box. Do NOT use an "=" sign or parentheses
    ' Params  : Standard Access list fill function parameters. See Access online help for details.
    ' Returns : List fill function return values based on the call type
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Static intCount As Integer
    Static aObjects() As String
    Dim varReturn As Variant
    
    Select Case intCode
        Case acLBInitialize
            ' Initialize code (run once)
            intCount = ObjectNamesToArray(acForm, aObjects())
            varReturn = intCount
            
        Case acLBOpen
            ' Open code
            varReturn = Timer
            
        Case acLBGetRowCount
            ' Get the number of rows (records)
            varReturn = intCount
            
        Case acLBGetColumnCount
            ' Get the number of columns (fields)
            varReturn = 1
            
        Case acLBGetColumnWidth
            ' Set default for column widths
            varReturn = -1
            
        Case acLBGetValue
            ' Return the data
            varReturn = aObjects(lngRow)
            
        Case acLBEnd
            ' Final call
            Erase aObjects
            
    End Select
    
    ListFillForms = varReturn
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.ListFillForms"
    Resume PROC_EXIT
End Function

Public Function ListFillMacros(ctl As Control, lngID As Long, lngRow As Long, lngCol As Long, intCode As Integer) As Variant
    ' Comments: Provides a list fill function for a list/combo box for a list of macros.
    '           Use the name of this function as the "RowSourceType" property of a listbox or combo box. Do NOT use an "=" sign or parentheses
    ' Params  : Standard Access list fill function parameters. See Access online help for details.
    ' Returns : List fill function return values based on the call type
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Static intCount As Integer
    Static aObjects() As String
    Dim varReturn As Variant
    
    Select Case intCode
        Case acLBInitialize
            ' Initialize code (run once)
            intCount = ObjectNamesToArray(acMacro, aObjects())
            varReturn = intCount
            
        Case acLBOpen
            ' Open code
            varReturn = Timer
            
        Case acLBGetRowCount
            ' Get the number of rows (records)
            varReturn = intCount
            
        Case acLBGetColumnCount
            ' Get the number of columns (fields)
            varReturn = 1
            
        Case acLBGetColumnWidth
            ' Set default for column widths
            varReturn = -1
            
        Case acLBGetValue
            ' Return the data
            varReturn = aObjects(lngRow)
            
        Case acLBEnd
            ' Final call
            Erase aObjects
            
    End Select
    
    ListFillMacros = varReturn
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.ListFillMacros"
    Resume PROC_EXIT
End Function

Public Function ListFillModules(ctl As Control, lngID As Long, lngRow As Long, lngCol As Long, intCode As Integer) As Variant
    ' Comments: Provides a list fill function for a list/combo box for a list of modules.
    '           Use the name of this function as the "RowSourceType" property of a listbox or combo box. Do NOT use an "=" sign or parentheses
    ' Params  : Standard Access list fill function parameters. See Access online help for details.
    ' Returns : List fill function return values based on the call type
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Static intCount As Integer
    Static aObjects() As String
    Dim varReturn As Variant
    
    Select Case intCode
        Case acLBInitialize
            ' Initialize code (run once)
            intCount = ObjectNamesToArray(acModule, aObjects())
            varReturn = intCount
            
        Case acLBOpen
            ' Open code
            varReturn = Timer
            
        Case acLBGetRowCount
            ' Get the number of rows (records)
            varReturn = intCount
            
        Case acLBGetColumnCount
            ' Get the number of columns (fields)
            varReturn = 1
            
        Case acLBGetColumnWidth
            ' Set default for column widths
            varReturn = -1
            
        Case acLBGetValue
            ' Return the data
            varReturn = aObjects(lngRow)
            
        Case acLBEnd
            ' Final call
            Erase aObjects
            
    End Select
    
    ListFillModules = varReturn
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.ListFillModules"
    Resume PROC_EXIT
    
End Function

Public Function ListFillReports(ctl As Control, lngID As Long, lngRow As Long, lngCol As Long, intCode As Integer) As Variant
    ' Comments: Provides a list fill function for a list/combo box for a list of reports.
    '           Use the name of this function as the "RowSourceType" property of a listbox or combo box. Do NOT use an "=" sign or parentheses
    ' Params  : Standard Access list fill function parameters. See Access online help for details.
    ' Returns : List fill function return values based on the call type
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Static intCount As Integer
    Static aObjects() As String
    Dim varReturn As Variant
    
    Select Case intCode
        Case acLBInitialize
            ' Initialize code (run once)
            intCount = ObjectNamesToArray(acForm, aObjects())
            varReturn = intCount
            
        Case acLBOpen
            ' Open code
            varReturn = Timer
            
        Case acLBGetRowCount
            ' Get the number of rows (records)
            varReturn = intCount
            
        Case acLBGetColumnCount
            ' Get the number of columns (fields)
            varReturn = 1
            
        Case acLBGetColumnWidth
            ' Set default for column widths
            varReturn = -1
            
        Case acLBGetValue
            ' Return the data
            varReturn = aObjects(lngRow)
            
        Case acLBEnd
            ' Final call
            Erase aObjects
            
    End Select
    
    ListFillReports = varReturn
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.ListFillReports"
    Resume PROC_EXIT
End Function

Private Function ObjectNamesToArray(intObjectType As AcObjectType, astrObjects() As String) As Integer
    ' Comments: Loads an array with names of all of an object type
    ' Params  : intObjectType       Access object type (acForm, acReport, acMacro, and acModule)
    '           astrObjects         Array updated with list of object names (0-based)
    ' Returns : Number of objects (may be zero)
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Dim intObjects As Integer
    Dim intCounter As Integer
    
    ' Get the count first
    Select Case intObjectType
        Case acTable
            intObjects = CurrentData.AllTables.Count
        Case acQuery
            intObjects = CurrentData.AllQueries.Count
        Case acServerView
            intObjects = CurrentData.AllViews.Count
        Case acStoredProcedure
            intObjects = CurrentData.AllStoredProcedures.Count
        Case acDiagram
            intObjects = CurrentData.AllDatabaseDiagrams.Count
        Case acFunction
            intObjects = CurrentData.AllFunctions.Count
        Case acForm
            intObjects = CurrentProject.AllForms.Count
        Case acReport
            intObjects = CurrentProject.AllReports.Count
        Case acMacro
            intObjects = CurrentProject.AllMacros.Count
        Case acModule
            intObjects = CurrentProject.AllModules.Count
    End Select
    
    If intObjects > 0 Then
        ReDim astrObjects(0 To intObjects - 1)
        Select Case intObjectType
            Case acTable
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentData.AllTables(intCounter).Name
                Next intCounter
                
            Case acQuery
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentData.AllQueries(intCounter).Name
                Next intCounter
                
            Case acServerView
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentData.AllViews(intCounter).Name
                Next intCounter
                
            Case acStoredProcedure
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentData.AllStoredProcedures(intCounter).Name
                Next intCounter
                
            Case acDiagram
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentData.AllDatabaseDiagrams(intCounter).Name
                Next intCounter
                
            Case acFunction
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentData.AllFunctions(intCounter).Name
                Next intCounter
                
            Case acForm
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentProject.AllForms(intCounter).Name
                Next intCounter
                
            Case acReport
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentProject.AllReports(intCounter).Name
                Next intCounter
                
            Case acMacro
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentProject.AllMacros(intCounter).Name
                Next intCounter
                
            Case acModule
                For intCounter = 0 To intObjects - 1
                    astrObjects(intCounter) = CurrentProject.AllModules(intCounter).Name
                Next intCounter
                
        End Select
    End If
    
    ObjectNamesToArray = intObjects
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.ObjectNamesToArray"
    Resume PROC_EXIT
End Function



Public Sub SelectControl(ctl As Control, fNone As Boolean)
    ' Comments: Selects/de-selects the contents of the specified control.
    '           This can be useful when you want your code to select all the text in a particular control.
    '           By default Access highlights an entire control's contents when you move into that control.
    '           You can use this function to remove the highlight by setting the fNone parameter to true.
    ' Params  : ctl         Handle to the control (must be a text box or combobox)
    '           fNone       True to de-select the contents, False to select the contents
    ' Returns : True if successful, False otherwise
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    ctl.SetFocus
    ctl.SelStart = 0
    
    If fNone Then
        ctl.SelLength = 0
    Else
        ctl.SelLength = Len(ctl)
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.SelectControl"
    Resume PROC_EXIT
End Sub

Public Sub WaitUntilFormClose(strForm As String)
    ' Comments: Wait for the form to close, using the Sleep command rather than DoEvents (which eats up CPU processing resources).
    '           This is useful to suspend the execution of code until the user closes a form and the form cannot be opened modally.
    ' Params  : strForm       Form name
    ' Source  : Total Visual SourceBook
    
    On Error GoTo PROC_ERR
    
    Do While SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0
        ' Pause for 100 milliseconds (0.1 seconds)
        DoEvents
        Sleep 100
    Loop
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.WaitUntilFormClose"
    Resume PROC_EXIT
End Sub



