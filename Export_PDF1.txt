Public Sub Main()
	On Error Resume Next
	Call ExportPDF1()
End Sub
Public Sub PublishPDF(ByVal SheetName As String, ByVal Index As Integer)
    On Error Resume Next
    ' Get the PDF translator Add-In.
    Dim PDFAddIn As TranslatorAddIn
        PDFAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
    'Set a reference to the active document (the document to be published).
    Dim oDocument As Document
        oDocument = ThisApplication.ActiveDocument
    Dim oContext As TranslationContext
        oContext = ThisApplication.TransientObjects.CreateTranslationContext
    oContext.Type = kFileBrowseIOMechanism
    ' Create a NameValueMap object
    Dim oOptions As NameValueMap
        oOptions = ThisApplication.TransientObjects.CreateNameValueMap
    ' Create a DataMedium object
    Dim oDataMedium As DataMedium
        oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
    ' Check whether the translator has 'SaveCopyAs' options
    If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
        ' Options for drawings...
        oOptions.Value("All_Color_AS_Black") = 0
        'oOptions.Value("Remove_Line_Weights") = 0
        'oOptions.Value("Vector_Resolution") = 400
        'oOptions.Value("Sheet_Range") = kPrintAllSheets
        oOptions.Value("Custom_Begin_Sheet") = Index
        oOptions.Value("Custom_End_Sheet") = Index
    End If
    
    oCustomPropertySet = ThisDoc.Document.PropertySets.Item("Inventor User Defined Properties")
    RegisterNO = iProperties.Value("Custom", "10. REGISTER NO.")
	Unit = iProperties.Value("Custom", "16.Unit")
    
    'Set the destination file name
    Dim FileNameT As String
	If SheetName = "DWG First Page" Then
		If Unit = "" Then
			FileNameT = RegisterNO & "-MEDWG" & ".PDF"
		Else
			FileNameT = RegisterNO & "-MEDWG" & Unit & ".PDF"
		End If
	Else
    	FileNameT = RegisterNO & "-MEDWG-" & SheetName & ".PDF"
	End If
		
    oDataMedium.FileName = "c:\\test\\" & FileNameT
    'Publish document.
    Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
End Sub

Public Sub ExportPDF1()
	On Error Resume Next
    Dim oDoc As DrawingDocument
        oDoc = ThisApplication.ActiveDocument
    
    Dim oTitleBlock As Inventor.TitleBlock
    Dim oTextBox As TextBox
    Dim oSheet As Sheet
    
    Dim lPos As Long
    Dim sSheetName As String
    
    Dim SheetIndex As Integer
    SheetIndex = 1
    
    For Each oSheet In oDoc.Sheets
        'If SheetIndex = 0 Then
            '123
        'Else
            'Set oTitleBlock = oSheet.TitleBlock
            lPos = InStr(oSheet.Name, ":")
            sSheetName = Left(oSheet.Name, lPos - 1)
            Call PublishPDF(sSheetName, SheetIndex)
            'For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
                'If oTextBox.Text = "<PART NO.>" Then
                    'Call oTitleBlock.SetPromptResultText(oTextBox, sSheetName)
                    'Call oPromptEntry  =  oTitleBlock.GetResultText(oTextBox)
                'End If
            'Next
        'End If
        SheetIndex = SheetIndex + 1
    Next
    
    'Set Change_Border_Part_No = New Border_Part_No
    'Call Change_Border_Part_No.Show
End Sub
