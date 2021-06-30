VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDocExportBatch 
   Caption         =   "DocExport: BATCH MODE"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   OleObjectBlob   =   "FormDocExportBatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDocExportBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Binary
' --------------------------------------------------------------------03/09/2007
' DocExport.swp                     Written by Leonard Kikstra,
'                                   Copyright 2003-2007, Leonard Kikstra
'                                   Downloaded from Lenny's SolidWorks Resources
'                                        at http://www.lennyworks.com/solidworks
' ------------------------------------------------------------------------------
' Version   1.00 01/27/2004 Initial released
'                           Combined ConfigExport and DrawingExport macros
'           1.10 05/11/2005 Added selections for document types to show
'           1.20 01/11/2007 Added single doc export for multi sheet drawings
'                               This is BATCH EXPORT only functionality
'                           Disable type select when document type not selected
'------------------------------------------------------------------------------
Dim longstatus As Long
Dim WorkDir As String
Dim ModType As Integer
Dim DocSelect As Integer
Dim SelDocCount As Integer
Dim SelDocStat As Integer
Dim ThisDocCount As Integer
Dim ThisDocStat As Integer
Dim ExpCount As Integer

Private Sub CheckBoxShowParts_Click()
  ModelCheck                                                    ' v 1.20
  PopulateFileList
End Sub

Private Sub CheckBoxShowAssem_Click()
  ModelCheck                                                    ' v 1.20
  PopulateFileList
End Sub

Private Sub CheckBoxShowDraw_Click()
  DrawCheck                                                     ' v 1.20
  PopulateFileList
End Sub

Private Sub CommandClearAll_Click()
  LabelStatus.Caption = ""                                     ' Clear status
  For x = 0 To ListBoxSelect.ListCount - 1                     ' Each config
    ListBoxSelect.Selected(x) = False                          ' Clear select
  Next
  CountSelected
End Sub

Private Sub CommandSelectAll_Click()
  LabelStatus.Caption = ""                                     ' Clear status
  For x = 0 To ListBoxSelect.ListCount - 1                     ' Each config
    ListBoxSelect.Selected(x) = True                           ' Set select
  Next
  CountSelected
End Sub

Private Sub CheckBoxFilter_Click()
  If CheckBoxFilter = True Then
    TextBoxFilter.Enabled = True
    TextBoxFilter.BackColor = vbWindowBackground
    CommandSelect.Enabled = True
    CommandClear.Enabled = True
  Else
    TextBoxFilter.Enabled = False
    TextBoxFilter.BackColor = vbButtonFace
    CommandSelect.Enabled = False
    CommandClear.Enabled = False
  End If
End Sub

Private Sub CommandClear_Click()
  For x = 0 To ListBoxSelect.ListCount - 1                     ' Each config
    For y = 1 To Len(ListBoxSelect.List(x, 0)) - Len(TextBoxFilter) + 1
      If UCase(Mid$(ListBoxSelect.List(x, 0), y, Len(TextBoxFilter))) _
             = UCase(TextBoxFilter) Then
        ListBoxSelect.Selected(x) = False                      ' Set select
      End If
    Next y
  Next x
  CountSelected
End Sub

Private Sub CommandSelect_Click()
  For x = 0 To ListBoxSelect.ListCount - 1                     ' Each config
    For y = 1 To Len(ListBoxSelect.List(x, 0)) - Len(TextBoxFilter) + 1
      If UCase(Mid$(ListBoxSelect.List(x, 0), y, Len(TextBoxFilter))) _
             = UCase(TextBoxFilter) Then
        ListBoxSelect.Selected(x) = True                       ' Set select
      End If
    Next y
  Next x
  CountSelected
End Sub

Private Sub CommandClose_Click()
  LabelStatus.Caption = ""                                     ' Clear status
  End
End Sub

Private Sub CommandExport_Click()
  ' Disable form controls during export
  CheckBoxFilter.Enabled = False
  TextBoxFilter.Enabled = False
  TextBoxFilter.BackColor = vbButtonFace
  CommandSelect.Enabled = False
  CommandClear.Enabled = False
  CommandSelectAll.Enabled = False
  CommandClearAll.Enabled = False
  ListBoxModExport.Enabled = False
  ListBoxDrawExport.Enabled = False
  CommandExport.Enabled = False
  CommandClose.Enabled = False
  ListBoxSelect.Locked = True
  TimeStart = Now                                              ' Start timer
  LabelStatus.Caption = ""                                     ' Clear status
  ' Count documents (models and drawings) selected
  
  SelDocCount = 0
  For x = 0 To ListBoxSelect.ListCount - 1                     ' each document
    If ListBoxSelect.Selected(x) = True Then                   ' doc selected
      SelDocCount = SelDocCount + 1
    End If
  Next x
  ' Export documents
  SelDocStat = 0
  ExpCount = 0
  
  For DocSelect = 0 To ListBoxSelect.ListCount - 1             ' each model
    If ListBoxSelect.Selected(DocSelect) = True Then           ' doc selected
      
      SelDocStat = SelDocStat + 1
      LabelAllDocStat.Caption = Str(SelDocStat) & " of" & Str(SelDocCount) _
                                & " documents."
      
      If UCase(Right$(ListBoxSelect.List(DocSelect, 0), 3)) = "PRT" Then
        ExportModels swDocPART
      ElseIf UCase(Right$(ListBoxSelect.List(DocSelect, 0), 3)) = "ASM" Then
        ExportModels swDocASSEMBLY
      ElseIf UCase(Right$(ListBoxSelect.List(DocSelect, 0), 3)) = "DRW" Then
        ExportDrawings swDocDRAWING
      Else
        MsgBox "Document type '" & UCase(Right$(ListBoxSelect.List(DocSelect, _
               0), 3)) & "' unknown.", vbExclamation
      End If
      
      ' Close all documents necessary for mirrored/referenced components
      Set OpenDoc = swApp.ActiveDoc()
      While Not OpenDoc Is Nothing
        swApp.QuitDoc (OpenDoc.GetTitle)
        Set OpenDoc = swApp.ActiveDoc()
      Wend
    End If
  Next DocSelect
  TimeEnd = Now                                                ' end timer
  LabelStatus.Caption = " Done: " & ExpCount & " documents exported in " & _
           Format(DateDiff("s", TimeStart, TimeEnd) / 60, "###.00") & _
           " minutes."
  CheckBoxFilter.Enabled = True
  TextBoxFilter.Enabled = True
  TextBoxFilter.BackColor = vbWindowBackground
  CommandSelect.Enabled = True
  CommandClear.Enabled = True
  CommandSelectAll.Enabled = True
  CommandClearAll.Enabled = True
  ListBoxModExport.Enabled = True
  ListBoxDrawExport.Enabled = True
  CommandExport.Enabled = True
  CommandClose.Enabled = True
  CheckBoxFilter_Click
  ListBoxSelect.Locked = False
End Sub

Private Sub ExportModels(ModType As Integer)
  Set Document = swApp.OpenDoc2(WorkDir + ListBoxSelect.List(DocSelect, 0), _
                  ModType, 1, 0, 1, longstatus)                ' load model
  DocName = Document.GetTitle                                  ' Doc title
  For i = 1 To Len(DocName)
    If Mid$(DocName, i, 1) = "." Then z = i                    ' last period
  Next i
  DocName = Left$(DocName, z - 1)                              ' remove ext
  DocPathname = Document.GetPathName                           ' Doc path
  For i = 1 To Len(DocPathname)
    If Mid$(DocPathname, i, 1) = "\" Then z = i                ' last slash
  Next i
  DocPath = Left$(DocPathname, z)                              ' remove doc nam
  numConfigs = Document.GetConfigurationCount()                ' Num configs
  ConfigNames = Document.GetConfigurationNames()               ' Configs Names
  
  ThisDocCount = numConfigs
  ThisDocStat = 0
  
  For cn = 0 To (numConfigs - 1)                               ' Each config
    ExportName = DocName & "_" & ConfigNames(cn)               ' Export name
    LabelStatus.Caption = " Regenerating: " & ExportName       ' Show name
    ThisDocStat = ThisDocStat + 1
    LabelThisDocStat.Caption = Str(ThisDocStat) & " of" & Str(ThisDocCount) _
                            & " configs."
    FormDocExportBatch.Repaint                                 ' Repaint form
    Document.ShowConfiguration2 ConfigNames(cn)                ' Show config
    Document.GraphicsRedraw2                                   ' Redraw screen
    For ExportType = 0 To ListBoxModExport.ListCount - 1       ' each type
      If ListBoxModExport.Selected(ExportType) = True Then     ' selected ?
        ExType = ListBoxModExport.List(ExportType, 1)
        If Right$(ExType, 1) = "*" Then
          ExType = Left$(ExType, Len(ExType) - 1)
          If ModType = swDocPART Then ExType = ExType & "prt"
          If ModType = swDocASSEMBLY Then ExType = ExType & "asm"
          If ModType = swDocDRAWING Then ExType = ExType & "drw"
        End If
        LabelStatus.Caption = " Exporting: " & ExportName & ExType ' show name
        FormDocExportBatch.Repaint                             ' Repaint form
        Document.SaveAs4 DocPath & ExportName & ExType, swSaveAsCurrentVersion, _
               swSaveAsOptions_Silent, Errors, Warnings        ' Create file
        ExpCount = ExpCount + 1
        LabelExportStat.Caption = Str(ExpCount) & " documents."
      End If
    Next ExportType
  Next cn
  Set Document = Nothing
End Sub

Private Sub ExportDrawings(ModType As Integer)
  Set Document = swApp.OpenDoc2(WorkDir + ListBoxSelect.List(DocSelect, 0), _
                  ModType, 1, 0, 1, longstatus)                ' load model
  DocName = Document.GetTitle                                  ' Doc title
  DocPathname = Document.GetPathName                           ' Doc path
  For i = 1 To Len(DocPathname)
    If Mid$(DocPathname, i, 1) = "\" Then z = i                ' last slash
  Next i
  DocPath = Left$(DocPathname, z)                              ' remove doc nam
  numSheets = Document.GetSheetCount()                         ' Num sheets
  SheetNames = Document.GetSheetNames()                        ' Configs Names
  
  ThisDocCount = numSheets
  ThisDocStat = 0
  
  For SH = 0 To (numSheets - 1)                                ' Each sheet
    ThisSheetName = SheetNames(SH)
    Document.ActivateSheet ThisSheetName                       ' Show sheet
    DocName = Document.GetTitle                                ' Doc title
    ExportName = DocName                                       ' Export name
    LabelStatus.Caption = " Regenerating: " & ExportName       ' Show name
    ThisDocStat = ThisDocStat + 1
    LabelThisDocStat.Caption = Str(ThisDocStat) & " of" & Str(ThisDocCount) _
                            & " sheets."
    FormDocExportBatch.Repaint                                 ' Repaint form
    Document.GraphicsRedraw2                                   ' Redraw screen
    
    ' Export each drawing sheet as an individual file by:
    '   * Activat sheet, Regenerate sheet, Export w/ sheet name
    For ExportType = 0 To ListBoxDrawExport.ListCount - 1      ' each type
      If ListBoxDrawExport.Selected(ExportType) = True _
      And UCase(ListBoxDrawExport.List(ExportType, 2)) = "MULTI" _
      Then                                                     ' selected ? 1.2
        LabelStatus.Caption = " Exporting: " & ExportName & _
               ListBoxDrawExport.List(ExportType, 1)           ' show name
        FormDocExportBatch.Repaint                             ' Repaint form
        Document.SaveAs4 DocPath & ExportName & _
               ListBoxDrawExport.List(ExportType, 1), _
               swSaveAsCurrentVersion, _
               swSaveAsOptions_Silent, Errors, Warnings        ' Create file
        ExpCount = ExpCount + 1
        LabelExportStat.Caption = Str(ExpCount) & " documents."
      End If
    Next ExportType
  Next SH
  
  ' Export complete drawing sheet as single file by:           ' v1.20
  '   * All sheets have been regenerated by steps above
  '   * Proceed with Exporting document
  ThisSheetName = SheetNames(0)                                '
  Document.ActivateSheet ThisSheetName                         ' Show sheet
  DocName = Document.GetTitle                                  ' Doc title
  ExportName = DocName                                         '
  For ExportType = 0 To ListBoxDrawExport.ListCount - 1        ' each type
    If ListBoxDrawExport.Selected(ExportType) = True _
    And UCase(ListBoxDrawExport.List(ExportType, 2)) = "SINGLE" _
    Then                                                       ' selected ?
      LabelStatus.Caption = " Exporting: " & ExportName & _
             ListBoxDrawExport.List(ExportType, 1)             ' show name
      FormDocExportBatch.Repaint                               ' Repaint form
      Document.SaveAs4 DocPath & ExportName & _
             ListBoxDrawExport.List(ExportType, 1), _
             swSaveAsCurrentVersion, _
             swSaveAsOptions_Silent, Errors, Warnings          ' Create file
      ExpCount = ExpCount + 1                                  '
      LabelExportStat.Caption = Str(ExpCount) & " documents."  '
    End If                                                     '
  Next ExportType                                              ' v1.20
  
  Set Document = Nothing
End Sub

Private Sub ListBoxUserSelect_Click()
  LabelStatus.Caption = ""
End Sub

Private Sub AddExportType(TypeList As Object, ExTitle As String, ExExt As String, ExQty As String)
  TypeList.AddItem ExTitle & " (*" & ExExt & ")"
  TypeList.List(TypeList.ListCount - 1, 1) = ExExt
  TypeList.List(TypeList.ListCount - 1, 2) = ExQty
End Sub

Private Sub GetExportTypes(TypeList As Object, ExportType As String)
  ' Setup export file types
  Dim ExpTyp As String
  Dim ExpExt As String
  Dim ExpQty As String                                         ' v1.20
  Source = swApp.GetCurrentMacroPathName                ' Get macro path & name
  Source = Left$(Source, Len(Source) - 3) + "ini"       ' Set source file name
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  If FileSys.FileExists(Source) Then                    ' Does source file exist?
    Open Source For Input As #1                         ' Open file for input.
    Do While Not EOF(1)                     ' Loop until end of file.
      Input #1, Reader                      ' Data into Reader variable
      ' Read in setting types and ranges
      If Reader = "[" & ExportType & "]" Then
        Do While Not EOF(1)                 ' Loop until end of file.
          Input #1, ExpTyp                  ' Data into ExpTyp variable
          If ExpTyp <> "" Then              ' ExpTyp variable valid
            Input #1, ExpExt                ' Data into ExpExt variable
            Input #1, ExpQty                ' Data into ExpQty ' v1.20
            AddExportType TypeList, ExpTyp, ExpExt, ExpQty     ' v1.20
          Else
            GoTo EndLoop                    ' Exit loop
          End If
        Loop
      End If
    Loop
  End If
EndLoop:
  Close #1    ' Close file.
End Sub

Private Sub ModelCheck()                                        ' v 1.20
  If CheckBoxShowParts = True Or CheckBoxShowAssem = True Then
    If ListBoxModExport.Enabled = False Then
      ListBoxModExport.Enabled = True
      ListBoxModExport.BackColor = vbWindowBackground
      ListBoxModExport.ForeColor = vbWindowText
    End If
  Else
    If ListBoxModExport.Enabled = True Then
      ListBoxModExport.Enabled = False
      ListBoxModExport.BackColor = vbButtonFace
      ListBoxModExport.ForeColor = vbGrayText
    End If
  End If
End Sub

Private Sub DrawCheck()                                         ' v 1.20
  If CheckBoxShowDraw = True Then
    ListBoxDrawExport.Enabled = True
    ListBoxDrawExport.BackColor = vbWindowBackground
    ListBoxDrawExport.ForeColor = vbWindowText
  Else
    ListBoxDrawExport.Enabled = False
    ListBoxDrawExport.BackColor = vbButtonFace
    ListBoxDrawExport.ForeColor = vbGrayText
  End If
End Sub

Private Sub PopulateFileList()
  ListBoxSelect.Clear
  ' Get list of documents and populate 'ListBoxSelect'
  
  If CheckBoxShowParts = True Then
    GetFileList "*.sldprt" ' get asm list
    GetFileList "*.prt"    ' get asm list
  End If
  If CheckBoxShowAssem = True Then
    GetFileList "*.sldasm" ' get asm list
    GetFileList "*.asm"    ' get asm list
  End If
  If CheckBoxShowDraw = True Then
    GetFileList "*.slddrw" ' get dwg list
    GetFileList "*.drw"    ' get dwg list
  End If
  ' Sort list of models 'ListBoxSelect'
  For x = 0 To ListBoxSelect.ListCount - 2
    For y = 0 To ListBoxSelect.ListCount - 2
      If ListBoxSelect.List(y, 0) > ListBoxSelect.List(y + 1, 0) Then
        SortTemp = ListBoxSelect.List(y, 0)
        ListBoxSelect.List(y, 0) = ListBoxSelect.List(y + 1, 0)
        ListBoxSelect.List(y + 1, 0) = SortTemp
      End If
    Next y
  Next x
  For x = 0 To ListBoxSelect.ListCount - 1
    ' Display file name without extension
    For i = 1 To Len(ListBoxSelect.List(x, 0))
      If Mid$(ListBoxSelect.List(x, 0), i, 1) = "." Then z = i
    Next i
    ListBoxSelect.List(x, 1) = Left$(ListBoxSelect.List(x, 0), z - 1)
    ' Display file type based on extension
    If UCase(Right$(ListBoxSelect.List(x, 0), 3)) = "PRT" Then
      ListBoxSelect.List(x, 2) = "Part Model"
    ElseIf UCase(Right$(ListBoxSelect.List(x, 0), 3)) = "ASM" Then
      ListBoxSelect.List(x, 2) = "Assembly Model"
    ElseIf UCase(Right$(ListBoxSelect.List(x, 0), 3)) = "DRW" Then
      ListBoxSelect.List(x, 2) = "Drawing file"
    Else
      ListBoxSelect.List(x, 2) = "Unknown"
    End If
  ' ListBoxSelect.List(x, 0)
  Next x
  CountSelected
End Sub

Private Sub GetFileList(FileType As String)
  Filename = Dir(WorkDir + FileType) ' get prt list
  Do While Filename <> ""
    ListBoxSelect.AddItem Filename
    Filename = Dir
  Loop
End Sub

Private Sub CountSelected()
  SelDocCount = 0
  For x = 0 To ListBoxSelect.ListCount - 1                     ' each document
    If ListBoxSelect.Selected(x) = True Then                   ' doc selected
      SelDocCount = SelDocCount + 1
    End If
  Next x
  LabelStatus.Caption = " " & SelDocCount & " selected of " & _
                        ListBoxSelect.ListCount & " listed."
End Sub

Private Sub CommandAbout_Click()
  FormAbout.Show
End Sub

Private Sub UserForm_Initialize()
  FormDocExportBatch.Caption = FormDocExportBatch.Caption + " " + Version
  WorkDir = swApp.GetCurrentWorkingDirectory
  ' Get list of documents and populate 'ListBoxSelect'
  PopulateFileList
  ' Populate model export types
  GetExportTypes ListBoxModExport, "MODEL"
  If ListBoxModExport.ListCount = 0 Then
    AddExportType ListBoxModExport, "STEP", ".step", "MULTI"        ' v1.20
    AddExportType ListBoxModExport, "IGES", ".igs", "MULTI"         ' v1.20
  End If
  ListBoxModExport.Selected(0) = True
  ' Populate drawing export types
  GetExportTypes ListBoxDrawExport, "DRAWING"
  If ListBoxDrawExport.ListCount = 0 Then
    AddExportType ListBoxDrawExport, "DWG vector", ".dwg", "MULTI"  ' v1.20
    AddExportType ListBoxDrawExport, "JPEG raster", ".jpg", "MULTI" ' v1.20
  End If
  ListBoxDrawExport.Selected(0) = True
  ' Clear selections and filter options
  CheckBoxFilter = False
  CommandClear_Click
  CheckBoxFilter_Click
  LabelStatus.Caption = ""                                    ' Clear status
  LabelAllDocStat.Caption = ""                                ' Clear status
  LabelThisDocStat.Caption = ""                               ' Clear status
  LabelExportStat.Caption = ""                                ' Clear status
  CommandClose.SetFocus
  CountSelected
End Sub
