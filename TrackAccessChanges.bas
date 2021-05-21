Dim debuggin As Boolean
Dim filepath As String
Dim stringOptionsTab(100) As String
Dim stringOptionsSection(100) As String
Dim stringOptions(100) As String
Dim stringOptionsFullName(100) As String
Dim stringOptionsSource(100) As String
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
'source: https://docs.microsoft.com/en-us/office/vba/access/concepts/settings/set-options-from-visual-basic
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
    
    AddOption 28, "Default Text Field Size", "Default Text Field Size", "Object Designers", "Table design"
    AddOption 29, "Default Number Field Size", "Default Number Field Size", "Object Designers", "Table design"
    AddOption 30, "Default Field Type", "Default Field Type", "Object Designers", "Table design"
    AddOption 31, "AutoIndex on Import/Create", "AutoIndex on Import/Create", "Object Designers", "Table design"
    AddOption 32, "Show Property Update Options Buttons", "Show Property Update Options Buttons", "Object Designers", "Table design"
    
    AddOption 33, "Show Table Names", "Show Table Names", "Object Designers", "Query design"
    AddOption 34, "Output All Fields", "Output All Fields", "Object Designers", "Query design"
    AddOption 35, "Enable AutoJoin", "Enable AutoJoin", "Object Designers", "Query design"
    AddOption 36, "ANSI Query Mode", "SQL Server Compatible Syntax (ANSI 92), This database", "Object Designers", "Query design"
    AddOption 37, "ANSI Query Mode Default", "SQL Server Compatible Syntax (ANSI 92), Default for new databases", "Object Designers", "Query design"
    AddOption 38, "Query Design Font Name", "Query design font, Font", "Object Designers", "Query design"
    AddOption 39, "Query Design Font Size", "Query design font, Size", "Object Designers", "Query design"
    
    AddOption 40, "Selection Behavior", "Selection Behavior", "Object Designers", "Forms/Reports design"
    AddOption 41, "Form Template", "Form Template", "Object Designers", "Forms/Reports design"
    AddOption 42, "Report Template", "Report Template", "Object Designers", "Forms/Reports design"
    AddOption 43, "Always Use Event Procedures", "Always Use Event Procedures", "Object Designers", "Forms/Reports design"
    
    AddOption 44, "Enable Error Checking", "Enable Error Checking", "Object Designers", "Error checking"
    AddOption 45, "Error Checking Indicator Color", "Error indicator color", "Object Designers", "Error checking"
    AddOption 46, "Unassociated Label and Control Error Checking", "Check for unassociated label and control", "Object Designers", "Error checking"
    AddOption 47, "New Unassociated Labels Error Checking", "Check for new unassociated labels", "Object Designers", "Error checking"
    AddOption 48, "Keyboard Shortcut Errors Error Checking", "Check for keyboard shortcut errors", "Object Designers", "Error checking"
    AddOption 49, "Invalid Control Properties Error Checking", "Check for invalid control properties", "Object Designers", "Error checking"
    AddOption 50, "Common Report Errors Error Checking", "Check for common report errors", "Object Designers", "Error checking"
    
    AddOption 51, "Spelling ignore words in UPPERCASE", "Ignore words in UPPERCASE", "Proofing", "Correct Spelling"
    AddOption 52, "Spelling ignore words with number", "Ignore words that contain numbers", "Proofing", "Correct Spelling"
    AddOption 53, "Spelling ignore Internet and file addresses", "Ignore Internet and file addresses", "Proofing", "Correct Spelling"
    AddOption 54, "Spelling suggest from main dictionary only", "Suggest from main dictionary only", "Proofing", "Correct Spelling"
    AddOption 55, "Spelling dictionary language", "Dictionary Language", "Proofing", "Correct Spelling"
    
    AddOption 56, "Move After Enter", "Move After Enter", "Advanced", "Editing"
    AddOption 57, "Behavior Entering Field", "Behavior Entering Field", "Advanced", "Editing"
    AddOption 58, "Arrow Key Behavior", "Arrow Key Behavior", "Advanced", "Editing"
    AddOption 59, "Cursor Stops at First/Last Field", "Cursor Stops at First/Last Field", "Advanced", "Editing"
    AddOption 60, "Default Find/Replace Behavior", "Default Find/Replace Behavior", "Advanced", "Editing"
    AddOption 61, "Confirm Record Changes", "Confirm Record Changes", "Advanced", "Editing"
    AddOption 62, "Confirm Document Deletions", "Confirm Document Deletions", "Advanced", "Editing"
    AddOption 63, "Confirm Action Queries", "Confirm Action Queries", "Advanced", "Editing"
    AddOption 64, "Default Direction", "Default Direction", "Advanced", "Editing"
    AddOption 65, "General Alignment", "General Alignment", "Advanced", "Editing"
    AddOption 66, "Cursor Movement", "Cursor Movement", "Advanced", "Editing"
    AddOption 67, "Datasheet Ime Control", "Datasheet Ime Control", "Advanced", "Editing"
    AddOption 68, "Use Hijri Calendar", "Use Hijri Calendar", "Advanced", "Editing"
    
    AddOption 69, "Size of MRU File List", "Show this number of Recent Documents", "Advanced", "Display"
    AddOption 70, "Show Status Bar", "Status bar", "Advanced", "Display"
    AddOption 71, "Show Animations", "Show Animations", "Advanced", "Display"
    AddOption 72, "Show Smart Tags on Datasheets", "Show Smart Tags on Datasheets", "Advanced", "Display"
    AddOption 73, "Show Smart Tags on Forms and Reports", "Show Smart Tags on Forms and Reports", "Advanced", "Display"
    AddOption 74, "Show Macro Names Column", "Show in Macro Design, Names column", "Advanced", "Display"
    AddOption 75, "Show Conditions Column", "Show in Macro Design, Conditions column", "Advanced", "Display"
    
    AddOption 76, "Left Margin", "Left Margin", "Advanced", "Printing"
    AddOption 77, "Right Margin", "Right Margin", "Advanced", "Printing"
    AddOption 78, "Top Margin", "Top Margin", "Advanced", "Printing"
    AddOption 79, "Bottom Margin", "Bottom Margin", "Advanced", "Printing"
    
    AddOption 80, "Provide Feedback with Sound", "Provide Feedback with Sound", "Advanced", "General"
    AddOption 81, "Four-Digit Year Formatting", "Use four-year digit year formatting, This database", "Advanced", "General"
    AddOption 82, "Four-Digit Year Formatting All Databases", "Use four-year digit year formatting, All databases", "Advanced", "General"
    
    AddOption 83, "Open Last Used Database When Access Starts", "Open Last Used Database When Access Starts", "Advanced", "Advanced"
    AddOption 84, "Default Open Mode for Databases", "Default open mode", "Advanced", "Advanced"
    AddOption 85, "Default Record Locking", "Default Record Locking", "Advanced", "Advanced"
    AddOption 86, "Use Row Level Locking", "Open databases by using record-level locking", "Advanced", "Advanced"
    AddOption 87, "OLE/DDE Timeout (sec)", "OLE/DDE Timeout (sec)", "Advanced", "Advanced"
    AddOption 88, "Refresh Interval (sec)", "Refresh Interval (sec)", "Advanced", "Advanced"
    AddOption 89, "Number of Update Retries", "Number of Update Retries", "Advanced", "Advanced"
    AddOption 90, "ODBC Refresh Interval (sec)", "ODBC Refresh Interval (sec)", "Advanced", "Advanced"
    AddOption 91, "Update Retry Interval (msec)", "Update Retry Interval (msec)", "Advanced", "Advanced"
    AddOption 92, "Ignore DDE Requests", "DDE operations, Ignore DDE requests", "Advanced", "Advanced"
    AddOption 93, "Enable DDE Refresh", "DDE operations, Enable DDE refresh", "Advanced", "Advanced"
    AddOption 94, "Command-Line Arguments", "Command-Line Arguments", "Advanced", "Advanced"
    
    AddOption 95, "AllowFullMenus", "Allow Full Menus", "Current Database", "Ribbon and Toolbar options", "C"
    Open filepath & "OPTIONS.txt" For Output As #1
    
    Dim i As Integer
    For i = 0 To 100
        x = GetOptionValue(i)
    Next i
End Sub
Private Function GetOptionValue(index As Integer) As String
    If stringOptionsSource(index) = "A" Then
        GetOptionValue = Application.GetOption(stringOptions(index))
        OutputWrite "Tab: " & stringOptionsTab(index) & ", Section: " & stringOptionsSection(index) & ", Option: " & stringOptionsFullName(index) & ", Value: " & GetOptionValue
    End If
    If stringOptionsSource(index) = "C" Then
        'GetOptionValue = CurrentDb.Properties(stringOptions(index))
        Dim i As Integer
        For i = 0 To CurrentDb.Properties.Count - 1
            If CurrentDb.Properties(i).Name <> "Connection" Then
                OutputWrite "CurrentDb.Properties: " & CurrentDb.Properties(i).Name & ", Value: " & CStr(CurrentDb.Properties(i).Value)
            End If
        Next i
    End If
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








Public Sub LinkSQLServerTables()
'This method relinks the SQL Server tables without using an ODBC DSN
    Dim strCnnStr As String
    Dim strCnnStrGED As String
    Dim db As Database
    Dim tblDef As TableDef
    Dim qryDef As QueryDef
    Dim intI As Integer
    
    Set db = CurrentDb
    
    SetNameOfSQLServer
    
    strCnnStr = GetSQLServerConnString
    strCnnStrGED = "ODBC;DSN=OrderTrackingGP;DATABASE=" & GBL_GEDSQLDatabaseName & ";AutoTranslate=No;AnsiNPW=No;" _
        & "Uid=" & GBL_GEDUserName & ";Pwd=" & GBL_GEDPassword & ";"
    
    
    'strCnnStr = "ODBC;DRIVER=SQL Server;SERVER=" & GBL_SQLServerName & "\" & GBL_SQLServerInstanceName & ";DATABASE=" & GBL_SQLDatabaseName & ";Trusted_Connection=Yes;AutoTranslate=No;AnsiNPW=No;"
    'strCnnStrGED = "ODBC;DRIVER=SQL Server;SERVER=" & GBL_GEDSQLServerName & ";DATABASE=" & GBL_GEDSQLDatabaseName & ";Uid=" _
    '& GBL_GEDUsername & ";Pwd=" & GBL_GEDPassword & ";AutoTranslate=No;AnsiNPW=No;"

    For intI = 0 To db.TableDefs.Count - 1
        Set tblDef = db.TableDefs(intI)
        If Left(tblDef.Name, 4) <> "Msys" And Left(tblDef.Name, 4) <> "~TMP" Then
            If InStr(tblDef.Connect, "SQL Server") > 0 Or InStr(tblDef.Connect, "ODBC;") > 0 Then ' Skip tables not SQL Server
                'Debug.Print tblDef.Name
                If Left(tblDef.Name, 4) = "dbo_" Then
                    tblDef.Connect = strCnnStrGED
                    tblDef.RefreshLink
                Else
                    'On Error Resume Next
                    tblDef.Connect = strCnnStr
                    tblDef.RefreshLink
                End If
            End If
         End If
    Next intI
    For intI = 0 To db.QueryDefs.Count - 1
        Set qryDef = db.QueryDefs(intI)
        If Left(qryDef.Name, 4) <> "Msys" And Left(qryDef.Name, 4) <> "~TMP" Then
            If InStr(qryDef.Connect, "SQL Server") > 0 Or InStr(qryDef.Connect, "ODBC;") > 0 Then ' Skip tables not SQL Server
            'On Error GoTo 0
                qryDef.Connect = strCnnStr
            End If
         End If
    Next intI

End Sub

