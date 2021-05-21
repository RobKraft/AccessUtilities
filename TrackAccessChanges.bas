Dim debuggin As Boolean
Dim filepath As String
Dim stringOptionsTab(40) As String
Dim stringOptionsSection(40) As String
Dim stringOptions(40) As String
Dim stringOptionsFullName(40) As String
Dim stringOptionsSource(40) As String
'Run this code to export Form Properties, Report Properties, code from Forms, modules, and reports, along with query sources to external files.
'Then you can check the files into a source control repo.  Do this daily and it gives you a good way to keep track of all the changes
'you made to your access database over time.
'Note that this does not include ALL form and report properties, but you can change the code below if you want ALL of them.

Sub GatherInfo()
    debuggin = False
    filepath = CurrentProject.Path & "\"
    exportOptions

    ExportAllCode
    robListAllFormProps
    robListAllReportProps
    robListAllQuerySQL
    robListAllTableSQL
End Sub
Sub exportOptions()
    AddOption 0, "Auto Compact", "Compact on Close", "Current Database", "Application Options"
    AddOption 1, "Remove Personal Information", "Remove personal information from file properties on save", "Current Database", "Application Options"
    AddOption 2, "Themed Form Controls", "Use Windows-themed Controls on Forms", "Current Database", "Application Options"
    AddOption 3, "DesignWithData", "Enable Layout View for this database", "Current Database", "Application Options"
    AddOption 4, "CheckTruncatedNumFields", "Check for truncated number fields", "Current Database", "Application Options"
    AddOption 5, "Picture Property Storage Format", "Picture Property Storage Format", "Current Database", "Application Options"
    
    AddOption 6, "Track Name AutoCorrect Info", "Track Name AutoCorrect Info", "Current Database", "Name AutoCorrect Options"
    AddOption 7, "Perform Name AutoCorrect", "Perform Name AutoCorrect", "Current Database", "Name AutoCorrect Options"
    AddOption 8, "Log Name AutoCorrect Changes", "Log Name AutoCorrect Changes", "Current Database", "Name AutoCorrect Options"
    
    
    AddOption 9, "Show Values in Indexed", "Show list of values in, Local indexed fields", "Current Database", "Filter lookup options"
    AddOption 10, "Show Values in Non-Indexed", "Show list of values in, Local nonindexed fields", "Current Database", "Filter lookup options"
    AddOption 11, "Show Values in Remote", "Show list of values in, ODBC fields", "Current Database", "Filter lookup options"
    AddOption 12, "Show Values in Snapshot", "Show list of values in, Records in local snapshot", "Current Database", "Filter lookup options"
    AddOption 13, "Show Values in Server", "Show list of values in, Records at server", "Current Database", "Filter lookup options"
    AddOption 14, "Show Values Limit", "Don't display lists where more than this number of records read", "Current Database", "Filter lookup options"
    
    AddOption 15, "Default Font Color", "Font color", "Datasheet", "Default colors"
    AddOption 16, "Default Background Color", "Background color", "Datasheet", "Default colors"
    AddOption 17, "_64", "Alternate background color", "Datasheet", "Default colors"
    AddOption 18, "Default Gridlines Color", "Gridlines color", "Datasheet", "Default colors"
    
    AddOption 19, "Default Gridlines Horizontal", "Default gridlines showing, Horizontal", "Datasheet", "Gridlines and cell effects"
    AddOption 20, "Default Gridlines Vertical", "efault gridlines showing, Vertical", "Datasheet", "Gridlines and cell effects"
    AddOption 21, "Default Cell Effect", "Default cell effect", "Datasheet", "Gridlines and cell effects"
    AddOption 22, "Default Column Width", "Default Column Width", "Datasheet", "Gridlines and cell effects"
    
    AddOption 23, "Default Font Name", "Font", "Datasheet", "Default font"
    AddOption 24, "Default Font Size", "Size", "Datasheet", "Default font"
    AddOption 25, "Default Font Weight", "Weight", "Datasheet", "Default font"
    AddOption 26, "Default Font Underline", "Underline", "Datasheet", "Default font"
    AddOption 27, "Default Font Italic", "Italic", "Datasheet", "Default font"
    
    AddOption 28, "AllowFullMenus", "Allow Full Menus", "Current Database", "Ribbon and Toolbar options", "C"
    Open filepath & "OPTIONS.txt" For Output As #1
    
    Dim i As Integer
    For i = 0 To 28
        x = GetOptionValue(i)
    Next i
End Sub
Private Function GetOptionValue(index As Integer) As String
    If stringOptionsSource(index) = "A" Then
        GetOptionValue = Application.GetOption(stringOptions(index))
    End If
    If stringOptionsSource(index) = "C" Then
        GetOptionValue = CurrentDb.Properties(stringOptions(index))
    End If
    OutputWrite "Tab: " & stringOptionsTab(index) & ", Section: " & stringOptionsSection(index) & ", Option: " & stringOptionsFullName(index) & ", Value: " & GetOptionValue
End Function
Private Sub AddOption(index As Integer, optionName As String, optionFullName As String, tabName As String, sectionName As String, Optional source As String = "A")
    stringOptionsTab(index) = tabName
    stringOptionsSection(index) = sectionName
    stringOptions(index) = optionName
    stringOptionsFullName(index) = optionFullName
    stringOptionsSource(index) = source
End Sub
Sub robListAllReportProps()
    Dim rpt As Report
    Dim reportIsLoaded As Boolean
    Dim outputThisProp As Boolean

    On Error Resume Next

    For Each rptHolder In Application.CurrentProject.AllReports
        reportIsLoaded = False
        For Each aLoadedReport In Application.Reports
            If aLoadedReport.Name = rptHolder.Name Then
                reportIsLoaded = True
            End If
        Next aLoadedReport

        If reportIsLoaded = False Then
            DoCmd.OpenReport rptHolder.Name, acViewDesign, , , acHidden
            If Err.Number <> 0 Then
                If debuggin Then
                    Debug.Print "Unable to analyze report: " & rptHolder.Name & " probably because of needing a specific printer. " & Err.Description
                Else
                    Print #1, "Unable to analyze report: " & rptHolder.Name & " probably because of needing a specific printer. " & Err.Description
                End If
            End If
        End If
        
        Set rpt = Application.Reports(rptHolder.Name)
        If debuggin Then
            Debug.Print rpt.Name
            ExportPropertiesOfThisFormOrReport rpt.Properties
            Debug.Print ""
        Else
            Open filepath & "PROPSforRPT_" & rpt.Name & ".txt" For Output As #1
            ExportPropertiesOfThisFormOrReport rpt.Properties
            Print #1, ""
        End If

        ExportPropertiesOfEachControlOnObject rpt.controls

        DoCmd.Close acReport, rpt.Name, acSaveNo

        If debuggin Then
        Else
            Close #1
        End If
    Next rptHolder
End Sub

Sub robListAllFormProps()
    'https://docs.microsoft.com/en-us/office/vba/api/access.accontroltype
    Dim frm As Form
    Dim formIsLoaded As Boolean
    Dim outputThisProp As Boolean

    For Each frmholder In Application.CurrentProject.AllForms
        formIsLoaded = False
        For Each aLoadedForm In Application.Forms
            If aLoadedForm.Name = frmholder.Name Then
                formIsLoaded = True
            End If
        Next aLoadedForm

        If formIsLoaded = False Then
            DoCmd.OpenForm frmholder.Name, acDesign, , , acFormReadOnly, acHidden
        End If
        
        Set frm = Application.Forms(frmholder.Name)
        
        If debuggin Then
            Debug.Print frm.Name
            ExportPropertiesOfThisFormOrReport frm.Properties
            Debug.Print ""
        Else
            Dim safeFormName As String
            safeFormName = Replace(frm.Name, "/", "slash")
            Open filepath & "PROPSforFRM_" & safeFormName & ".txt" For Output As #1
            ExportPropertiesOfThisFormOrReport frm.Properties
            Print #1, ""
        End If

        ExportPropertiesOfEachControlOnObject frm.controls


        DoCmd.Close acForm, frm.Name, acSaveNo

        If debuggin Then
        Else
            Close #1
        End If
    Next frmholder



End Sub
Private Sub OutputWrite(output As String)
    If debuggin Then
        Debug.Print output
    Else
        Print #1, output
    End If
End Sub
Private Sub robListAllTableSQL()


    Dim outputThisProp As Boolean

    For Each qryd In Application.CurrentDb.TableDefs
        Open filepath & "TABLE_" & qryd.Name & ".tbl" For Output As #1
        OutputWrite "TABLE: " & qryd.Name

        If Left(qryd.Name, 1) <> "~" Then
            For Each prp In qryd.Properties
                outputThisProp = True
                If prp.Name = "ConflictTable" Or prp.Name = "ReplicaFilter" Or prp.Name = "NameMap" Or prp.Name = "GUID" Then
                    outputThisProp = False
                End If

                If outputThisProp = True Then
                    OutputWrite vbTab & prp.Name & " " & Trim(prp.Value)
                End If
            Next
            For Each fld In qryd.Fields
                OutputWrite vbTab & "FIELD: " & qryd.Name & ".[" & fld.Name & "]"

                For Each prp In fld.Properties
                    outputThisProp = True

                    If prp.Name = "Value" Or prp.Name = "ValidateOnSet" Or prp.Name = "ForeignName" Or prp.Name = "FieldSize" Or prp.Name = "OriginalValue" _
                     Or prp.Name = "VisibleValue" Or prp.Name = "GUID" Then
                        outputThisProp = False
                    End If

                    If outputThisProp = True Then
                        OutputWrite vbTab & vbTab & prp.Name & " " & Trim(prp.Value)
                    End If
                Next
            Next
            On Error Resume Next
            For Each fld In qryd.Indexes
                If Err.Number <> 0 Then
                    OutputWrite "Error processing table indexes: " & Err.Number & " - " & Err.Description
                    Exit For
                End If
                OutputWrite vbTab & "INDEX: " & qryd.Name & ".[" & fld.Name & "]"

                For Each prp In fld.Properties
                    outputThisProp = True
                    If outputThisProp = True Then
                        OutputWrite vbTab & vbTab & prp.Name & " " & Trim(prp.Value)
                    End If
                Next
            Next
            Close #1
        End If
    Next qryd
    Exit Sub

    
End Sub

Private Sub robListAllQuerySQL()

    For Each qryd In Application.CurrentDb.QueryDefs
        If Left(qryd.Name, 1) <> "~" Then
            Open filepath & "QUERY_" & qryd.Name & ".qry" For Output As #1
            Print #1, Trim(qryd.SQL)
            Close #1
        End If
    Next qryd
End Sub
Private Sub ExportPropertiesOfThisFormOrReport(ctl As Properties)
    For Each prp In ctl
        outputThisProp = True
        If Left(prp.Name, 3) = "Sel" Or Left(prp.Name, 7) = "Current" Or prp.Name = "Picture" Or prp.Name = "ImageData" Or LCase(Left(prp.Name, 3)) = "prt" Or prp.Name = "PictureData" Or Left(prp.Name, 7) = "Palette" Then
            'We get errors trying to export a picture or PaletteSource, or Current..., or SelectionChanged events
            outputThisProp = False
        End If
        If prp.Name = "Hwnd" Or prp.Name = "WindowWidth" Or prp.Name = "InsideWidth" Then
            'These values are constantly changing - don't want to try to track them
            outputThisProp = False
        End If

        If Left(prp.Name, 2) = "On" Then
            'Methods are properties, but I only want to export ones that have some code linked to them
            If Trim(prp.Value) <> "" Then
                outputThisProp = True
            Else
                outputThisProp = False
            End If
        End If
        If (Left(prp.Name, 6) = "Before" Or Left(prp.Name, 5) = "After") Then
            'Methods are properties, but I only want to export ones that have some code linked to them
            If Trim(prp.Value) <> "" Then
                outputThisProp = True
            Else
                outputThisProp = False
            End If
        End If
        If Right(prp.Name, 5) = "Macro" Then
            'Macros export within their forms, but I only export ones that have code linked to them
            If Trim(prp.Value) <> "" Then
                outputThisProp = True
            Else
                outputThisProp = False
            End If
        End If
        If outputThisProp = True Then
            If debuggin Then
                Debug.Print prp.Name & " " & Trim(prp.Value)
            Else
                Print #1, prp.Name & " " & Trim(prp.Value)
            End If
        End If
    Next prp
End Sub
Private Sub ExportPropertiesOfEachControlOnObject(controls As controls)
    For Each ctl In controls
        If ctl.ControlType <> acObjectFrame Then 'And ctl.ControlType <> acRectangle And ctl.ControlType <> acPage And ctl.ControlType <> acLine _
            '           And ctl.ControlType <> acObjectLabel And ctl.ControlType <> acPageBreak And ctl.ControlType <> acTabCtl _
            '           And ctl.ControlType <> acImage And ctl.ControlType <> acCommandButton Then
            If debuggin Then
                Debug.Print TypeName(ctl) & " - Name = " & ctl.Properties("Name")
            Else
                Print #1, TypeName(ctl) & " - " & ctl.Properties("Name")
            End If

            For Each prp In ctl.Properties
                outputThisProp = False
                If prp.Name = "LabelName" Or prp.Name = "LpOleObject" Or prp.Name = "InSelection" Or prp.Name = "Text" Or prp.Name = "SelText" Or prp.Name = "SelStart" Or prp.Name = "SelLength" Or prp.Name = "ListCount" Or prp.Name = "ListIndex" Or prp.Name = "PictureData" Or prp.Name = "ImageData" Or Left(prp.Name, 7) = "Palette" Or prp.Name = "ObjectPalette" Then
                Else
                    outputThisProp = True
                    If ctl.ControlType = acTextBox Then
                        If Left(prp.Name, 16) = "ConditionalFormat" Then
                            If Trim(prp.Value) <> "" Then
                                outputThisProp = False
                                If debuggin Then
                                    Debug.Print vbTab & prp.Name & " " & " HAS A VALUE!!!"
                                    Else
                                    Print #1, vbTab & prp.Name & " " & " HAS A VALUE!!!"
                                    End If
                                'https://stackoverflow.com/questions/63839201/access-application-saveastext-read-hex-values
                            End If
                        End If
                    End If
                    If Left(prp.Name, 2) = "On" Then
                        'Methods are properties, but I only want to export ones that have some code linked to them
                        If Trim(prp.Value) <> "" Then
                            outputThisProp = True
                        Else
                            outputThisProp = False
                        End If
                    End If
                    If Right(prp.Name, 5) = "Macro" Then
                        'Macros export within their forms, but I only export ones that have code linked to them
                        If Trim(prp.Value) <> "" Then
                            outputThisProp = True
                        Else
                            outputThisProp = False
                        End If
                    End If

                    If (prp.Name = "BeforeUpdate" Or prp.Name = "AfterUpdate") Then
                        'Methods are properties, but I only want to export ones that have some code linked to them
                        If Trim(prp.Value) <> "" Then
                            outputThisProp = True
                        Else
                            outputThisProp = False
                        End If
                    End If
                    If outputThisProp = True Then
                        If debuggin Then
                            Debug.Print vbTab & prp.Name & " " & Trim(prp.Value)
                        Else
                            Print #1, vbTab & prp.Name & " " & Trim(prp.Value)
                        End If
                    End If
                End If
            Next prp
        End If
    Next ctl

End Sub


'This method exports all the Visual Basic code - it will only create files for objects that have code behind them
Public Sub ExportAllCode()

    Dim vbComponent As Variant
    Dim suffix As String
    Dim filen As String
    Dim prefix As String 'Prefixes to group all Modules and Classes together

    'I chose the .bas extension for all files because I view them in Visual Studio and .bas format nicely
    For Each vbComponent In Application.VBE.VBProjects(1).VBComponents
        prefix = ""
        Select Case vbComponent.Type
            Case 2 'vbext_ct_ClassModule, vbext_ct_Document
                suffix = ".bas"
                prefix = "Class_"
            Case 100 'vbext_ct_MSForm
                If Left(vbComponent.Name, 6) = "Report" Then
                    suffix = ".bas"
                Else
                    suffix = ".bas"
                End If
            Case 1 'vbext_ct_StdModule
                suffix = ".bas"
                prefix = "Module_"
            Case Else
                suffix = ""
        End Select

        filen = vbComponent.Name
        If suffix <> "" Then
            vbComponent.Export _
                FileName:=CurrentProject.Path & "\" & prefix & filen & suffix
        End If
    Next vbComponent

End Sub
