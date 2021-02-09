Attribute VB_Name = "Module1"
Option Compare Database
    Dim debuggin As Boolean
    Dim filepath As String

'Run this code to export Form Properties, Report Properties, code from Forms, modules, and reports, along with query sources to external files.
'Then you can check the files into a source control repo.  Do this daily and it gives you a good way to keep track of all the changes
'you made to your access database over time.
'Note that this does not include ALL form and report properties, but you can change the code below if you want ALL of them.

Sub GatherInfo()
    debuggin = False
    filepath = CurrentProject.Path & "\"
    
    ExportAllCode
    robListAllFormProps
    robListAllReportProps
    robListAllQuerySQL
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
            Debug.Print "RecordSource = " & Trim(rpt.RecordSource)
            Debug.Print "Filter = " & Trim(rpt.Filter)
            ProcessFormOrReportMethods rpt.Properties
            Debug.Print ""
        Else
            Open filepath & "REPORTPROPSfor_" & rpt.Name & ".txt" For Output As #1
            Print #1, "RecordSource = " & Trim(rpt.RecordSource)
            Print #1, "Filter = " & Trim(rpt.Filter)
            ProcessFormOrReportMethods rpt.Properties
            Print #1, ""
        End If
        
        ProcessControls rpt.controls
        
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
            Debug.Print "RecordSource = " & Trim(frm.RecordSource)
            Debug.Print "Filter = " & Trim(frm.Filter)
            ProcessFormOrReportMethods frm.Properties
            Debug.Print ""
        Else
            Open filepath & "FORMPROPSfor_" & frm.Name & ".txt" For Output As #1
            Print #1, "RecordSource = " & Trim(frm.RecordSource)
            Print #1, "Filter = " & Trim(frm.Filter)
            ProcessFormOrReportMethods frm.Properties
            Print #1, ""
        End If

        ProcessControls frm.controls
        
        
        DoCmd.Close acForm, frm.Name, acSaveNo
    
        If debuggin Then
        Else
            Close #1
        End If
    Next frmholder
    
    

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
Private Sub ProcessFormOrReportMethods(ctl As Properties)
    For Each prp In ctl
        outputThisProp = False
        If Left(prp.Name, 2) = "On" Then
                If Trim(prp.Value) <> "" Then
                    outputThisProp = True
                End If
            End If
            If (prp.Name = "BeforeUpdate" Or prp.Name = "AfterUpdate") Then
                If Trim(prp.Value) <> "" Then
                    outputThisProp = True
                End If
            End If
            If outputThisProp = True Then
                If debuggin Then
                    Debug.Print prp.Name & " " & Trim(prp.Value)
                Else
                    Print #1, prp.Name & " " & Trim(prp.Value)
                End If
            End If
        'End If
    Next prp
End Sub
Private Sub ProcessControls(controls As controls)
        For Each ctl In controls
            If ctl.ControlType <> acLabel And ctl.ControlType <> acRectangle And ctl.ControlType <> acPage And ctl.ControlType <> acLine _
                And ctl.ControlType <> acObjectFrame And ctl.ControlType <> acPageBreak And ctl.ControlType <> acTabCtl _
                And ctl.ControlType <> acCommandButton Then
                If debuggin Then
                    Debug.Print TypeName(ctl) & " - Name = " & ctl.Properties("Name")
                Else
                    Print #1, TypeName(ctl) & " - " & ctl.Properties("Name")
                End If
        
                For Each prp In ctl.Properties
                    outputThisProp = False
                    If prp.Name = "LabelName" Or prp.Name = "Text" Or prp.Name = "SelText" Or prp.Name = "SelStart" Or prp.Name = "SelLength" Or prp.Name = "InSelection" Or prp.Name = "ListCount" Or prp.Name = "ListIndex" Then
                    Else
                        If ctl.ControlType = acTextBox Then
                            If prp.Name = "ControlSource" Or prp.Name = "DefaultValue" Then
                                outputThisProp = True
                            End If
                        ElseIf ctl.ControlType = acCheckBox Then
                            If prp.Name = "ControlSource" Or prp.Name = "DefaultValue" Then
                                outputThisProp = True
                            End If
                        ElseIf ctl.ControlType = acListBox Then
                            If prp.Name = "ControlSource" Or prp.Name = "ColumnCount" Or prp.Name = "RowSource" Or prp.Name = "RowSourceType" Or prp.Name = "BoundColumn" Then
                                outputThisProp = True
                            End If
                        ElseIf ctl.ControlType = acComboBox Then
                            If prp.Name = "ControlSource" Or prp.Name = "ColumnCount" Or prp.Name = "RowSource" Or prp.Name = "RowSourceType" Or prp.Name = "BoundColumn" Then
                                outputThisProp = True
                            End If
                        ElseIf ctl.ControlType = acOptionGroup Or ctl.ControlType = acOptionButton Then
                            If prp.Name = "ControlSource" Then
                                outputThisProp = True
                            End If
                        ElseIf ctl.ControlType = acSubform Or ctl.ControlType = acToggleButton Then
                            If prp.Name = "SourceObject" Or Left(prp.Name, 4) = "Link" Then
                                outputThisProp = True
                            End If
                        Else
                            If ctl.ControlType = acRectangle Or ctl.ControlType = acPage Or ctl.ControlType = acLine Or ctl.ControlType = acObjectFrame Or ctl.ControlType = acPageBreak Or ctl.ControlType = acTabCtl Then
                            Else
                                outputThisProp = True
                            End If
                        End If
                        If Left(prp.Name, 2) = "On" Then
                            If Trim(prp.Value) <> "" Then
                                outputThisProp = True
                            End If
                        End If
                        If (prp.Name = "BeforeUpdate" Or prp.Name = "AfterUpdate") Then
                            If Trim(prp.Value) <> "" Then
                                outputThisProp = True
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



Public Sub ExportAllCode()

    Dim c As Variant
    Dim Sfx As String
    Dim filen As String

    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case 2 'vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case 100 'vbext_ct_MSForm
                Sfx = ".frm"
            Case 1 'vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select
        
        filen = c.Name
        If (Left(filen, 18) = "report_rpt_company") Then
            If (Right(filen, 7) = "summary") Then
                filen = "report_rpt_company_YearlySales_Summary"
            End If
        End If
        If Sfx <> "" Then
            'If Left(c.Name, 16) = "form_frm_product" Then
            c.Export _
                FileName:=CurrentProject.Path & "\" & _
                filen & Sfx
            'End If
        End If
    Next c

End Sub
