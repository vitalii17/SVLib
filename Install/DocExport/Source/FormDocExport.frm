VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDocExport 
   Caption         =   "DocExport: CURRENT DOCUMENT"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   OleObjectBlob   =   "FormDocExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDocExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------------------------------03/09/2007
' DocExport.swp                     Written by Leonard Kikstra,
'                                   Copyright 2003-2007, Leonard Kikstra
'                                   Downloaded from Lenny's SolidWorks Resources
'                                        at http://www.lennyworks.com/solidworks
' ------------------------------------------------------------------------------
' Version   1.00 01/24/2004 Initial released
'                           Combined ConfigExport and DrawingExport macros
'------------------------------------------------------------------------------

Private Sub CommandAbout_Click()
  FormAbout.Show
End Sub

Private Sub CommandClearAll_Click()
  LabelStatus.Caption = ""                                     ' Clear status
  For x = 0 To ListBoxUserSelect.ListCount - 1                 ' Each config
    ListBoxUserSelect.Selected(x) = False                      ' Clear select
  Next
End Sub

Private Sub CommandSelectAll_Click()
  LabelStatus.Caption = ""                                     ' Clear status
  For x = 0 To ListBoxUserSelect.ListCount - 1                 ' Each config
    ListBoxUserSelect.Selected(x) = True                       ' Set select
  Next
End Sub

Private Sub CheckBoxFilter_Click()
  If CheckBoxFilter = True Then
    TextBoxFilter.Enabled = True
    TextBoxFilter.BackColor = vbWindowBackground
    CommandSelect.Enabled = True
    CommandClear.Enabled = True
    CheckBoxCaseSensitive.Enabled = True
  Else
    TextBoxFilter.Enabled = False
    TextBoxFilter.BackColor = vbButtonFace
    CommandSelect.Enabled = False
    CommandClear.Enabled = False
    CheckBoxCaseSensitive.Enabled = False
  End If
End Sub

Private Sub CommandClear_Click()
  For x = 0 To ListBoxUserSelect.ListCount - 1                 ' Each config
    For y = 1 To Len(ListBoxUserSelect.List(x, 0)) - Len(TextBoxFilter) + 1
      If CheckBoxCaseSensitive = True Then
        If Mid$(ListBoxUserSelect.List(x, 0), y, Len(TextBoxFilter)) _
               = TextBoxFilter Then
          ListBoxUserSelect.Selected(x) = False                ' Set select
        End If
      Else
        If UCase(Mid$(ListBoxUserSelect.List(x, 0), y, Len(TextBoxFilter))) _
               = UCase(TextBoxFilter) Then
          ListBoxUserSelect.Selected(x) = False                ' Set select
        End If
      End If
    Next y
  Next x
End Sub

Private Sub CommandSelect_Click()
  For x = 0 To ListBoxUserSelect.ListCount - 1                 ' Each config
    For y = 1 To Len(ListBoxUserSelect.List(x, 0)) - Len(TextBoxFilter) + 1
      If CheckBoxCaseSensitive = True Then
        If Mid$(ListBoxUserSelect.List(x, 0), y, Len(TextBoxFilter)) _
               = TextBoxFilter Then
          ListBoxUserSelect.Selected(x) = True                   ' Set select
        End If
      Else
        t = UCase(Mid$(ListBoxUserSelect.List(x, 0), y, Len(TextBoxFilter)))
        If UCase(Mid$(ListBoxUserSelect.List(x, 0), y, Len(TextBoxFilter))) _
               = UCase(TextBoxFilter) Then
          ListBoxUserSelect.Selected(x) = True                   ' Set select
        End If
      End If
    Next y
  Next x
End Sub

Private Sub CommandAllTypes_Click()
  LabelStatus.Caption = ""                                     ' Clear status
  For x = 0 To ListBoxExport.ListCount - 1                     ' Each Type
    ListBoxExport.Selected(x) = True                           ' Clear select
  Next
End Sub

Private Sub CommandClose_Click()
  LabelStatus.Caption = ""                                     ' Clear status
  End
End Sub

Private Sub CommandExport_Click()
  CheckBoxFilter.Enabled = False
  TextBoxFilter.Enabled = False
  TextBoxFilter.BackColor = vbButtonFace
  CommandSelect.Enabled = False
  CommandClear.Enabled = False
  CommandSelectAll.Enabled = False
  CommandClearAll.Enabled = False
  CommandAllTypes.Enabled = False
  ListBoxUserSelect.Locked = True
  ListBoxExport.Locked = True
  CommandExport.Enabled = False
  CommandClose.Enabled = False
  TimeStart = Now
  CT = 0
  For x = 0 To ListBoxUserSelect.ListCount - 1                 ' each item
    If ListBoxUserSelect.Selected(x) = True Then               ' item select
      CT = CT + 1
    End If
  Next                                                         ' Get next cfg
  LabelStatus.Caption = ""                                     ' Clear status
  
  DocPathname = Document.GetPathName                           ' Doc path
  For i = 1 To Len(DocPathname)
    If Mid$(DocPathname, i, 1) = "\" Then z = i                ' last slash
  Next i
  DocPath = Left$(DocPathname, z)                              ' remove doc nam
  
  Count = 0
  For x = 0 To ListBoxUserSelect.ListCount - 1                 ' each item
    If ListBoxUserSelect.Selected(x) = True Then               ' item select
      If DocMode = 0 Then               ' -- SolidWorks Models --
        DocName = Document.GetTitle                            ' Doc title
        For i = 1 To Len(DocName)
        If Mid$(DocName, i, 1) = "." Then z = i                ' last period
        Next i
        DocName = Left$(DocName, z - 1)                        ' remove ext
        ThisConfigName = ListBoxUserSelect.List(x, 0)
        ExportName = DocName & "_" & ThisConfigName            ' Exp name
        LabelStatus.Caption = " (" & Count + 1 & "/" & CT & ")" & _
                              " Exporting: " & ExportName      ' Show name
        FormDocExport.Repaint                                  ' Repaint form
        Document.ShowConfiguration2 ThisConfigName             ' Show config
        Document.GraphicsRedraw2                               ' Redraw screen
        For ExportType = 0 To ListBoxExport.ListCount - 1      ' each type
          If ListBoxExport.Selected(ExportType) = True Then    ' selected ?
            
            ExType = ListBoxExport.List(ExportType, 1)
            If Right$(ExType, 1) = "*" Then
              ExType = Left$(ExType, Len(ExType) - 1)
              If FileTyp = swDocPART Then ExType = ExType & "prt"
              If FileTyp = swDocASSEMBLY Then ExType = ExType & "asm"
              If FileTyp = swDocDRAWING Then ExType = ExType & "drw"
            End If
            
            LabelStatus.Caption = " Exporting: " & ExportName & ExType ' show name
            FormDocExport.Repaint                              ' Repaint form
            Document.SaveAs4 DocPath & ExportName & ExType, swSaveAsCurrentVersion, _
               swSaveAsOptions_Silent, Errors, Warnings        ' Create file
            Count = Count + 1
          End If
        Next ExportType
        CheckBoxFilter.Enabled = True
        ExType = "configurations"
      Else                              ' -- SolidWorks Drawings --
        FormDocExport.Repaint                                  ' Repaint form
        ThisSheetName = ListBoxUserSelect.List(x, 0)
        Document.ActivateSheet ThisSheetName                   ' Show sheet
        DocName = Document.GetTitle                            ' Doc title
        ExportName = DocName                                   ' Export name
        LabelStatus.Caption = " Exporting: " & ExportName      ' Show name
        Document.GraphicsRedraw2                               ' Redraw screen
        For ExportType = 0 To ListBoxExport.ListCount - 1      ' each type
          If ListBoxExport.Selected(ExportType) = True Then    ' selected ?
            LabelStatus.Caption = " Exporting: " & ExportName & _
              ListBoxExport.List(ExportType, 1)                ' show name
            FormDocExport.Repaint                              ' Repaint form
            Document.SaveAs4 DocPath & ExportName & _
               ListBoxExport.List(ExportType, 1), _
               swSaveAsCurrentVersion, _
               swSaveAsOptions_Silent, Errors, Warnings        ' Create file
            Count = Count + 1
          End If
        Next ExportType
        ExType = "sheets"
      End If
    End If
  Next                                                         ' Get next cfg
  TimeEnd = Now
  LabelStatus.Caption = " Done: " & Count & " " & ExType & " exported in " & _
           Format(DateDiff("s", TimeStart, TimeEnd) / 60, "###.00") & " minutes."
  CheckBoxFilter.Enabled = True
  CheckBoxFilter_Click
  CommandSelectAll.Enabled = True
  CommandClearAll.Enabled = True
  CommandAllTypes.Enabled = True
  ListBoxUserSelect.Locked = False
  ListBoxExport.Locked = False
  CommandExport.Enabled = True
  CommandClose.Enabled = True
End Sub

Private Sub ListBoxUserSelect_Click()
  LabelStatus.Caption = ""
End Sub

Private Sub AddExportType(ExTitle As String, ExExt As String, ExQty As String)
  ListBoxExport.AddItem ExTitle & " (*" & ExExt & ")"
  ListBoxExport.List(ListBoxExport.ListCount - 1, 1) = ExExt
  ListBoxExport.List(ListBoxExport.ListCount - 1, 2) = ExQty
End Sub

Private Sub GetExportTypes(ExportType As String)
  ' Setup export file types
  Dim ExpTyp As String
  Dim ExpExt As String
  Dim ExpQty As String
  Source = swApp.GetCurrentMacroPathName           ' Get macro path & name
  Source = Left$(Source, Len(Source) - 3) + "ini"  ' Set source file name
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  If FileSys.FileExists(Source) Then               ' Does source file exist?
    Open Source For Input As #1                    ' Open file for input.
    Do While Not EOF(1)                            ' Loop until end of file.
      Input #1, Reader                             ' Data into Reader variable
      ' Read in setting types and ranges
      If Reader = "[" & ExportType & "]" Then
        Do While Not EOF(1)                        ' Loop until end of file.
          Input #1, ExpTyp                         ' Data into ExpTyp variable
          If ExpTyp <> "" Then                     ' ExpTyp variable valid
            Input #1, ExpExt                       ' Data into ExpExt variable
            Input #1, ExpQty                       ' Data into ExpQty variable
            AddExportType ExpTyp, ExpExt, ExpQty
          Else
            GoTo EndLoop                           ' Exit loop
          End If
        Loop
      End If
    Loop
  End If
EndLoop:
  Close #1                                         ' Close file.
End Sub

Private Sub UserForm_Initialize()
  FormDocExport.Caption = FormDocExport.Caption + " " + Version
  If DocMode = 0 Then                   ' -- SolidWorks Models --
    ListBoxUserSelect.Clear                                    ' Clear form
    numConfigs = Document.GetConfigurationCount()              ' Num configs
    ConfigNames = Document.GetConfigurationNames()             ' Configs Names
    For x = 0 To (numConfigs - 1)                              ' Each config
      ListBoxUserSelect.AddItem ConfigNames(x)                 ' Add to list
    Next
    ' Sort entries in ListBoxUserSelect via simple bubble sort technique.
    For x = 0 To ListBoxUserSelect.ListCount - 2               ' Num of passes
      For y = 0 To ListBoxUserSelect.ListCount - 2             ' Pass each row
        If ListBoxUserSelect.List(y, 0) _
           > ListBoxUserSelect.List(y + 1, 0) Then             ' If greater,
          Temp = ListBoxUserSelect.List(y, 0)                  ' Swap entries
          ListBoxUserSelect.List(y, 0) = ListBoxUserSelect.List(y + 1, 0)
          ListBoxUserSelect.List(y + 1, 0) = Temp
        End If
      Next y
    Next x
    ' Automatically select current configuration
    Set Configuration = Document.GetActiveConfiguration        ' Active config
    CurrentConfigName = Configuration.Name                     ' name
    For x = 0 To (ListBoxUserSelect.ListCount - 1)             ' Each config
      If ListBoxUserSelect.List(x, 0) = CurrentConfigName Then ' Active conf?
        ListBoxUserSelect.Selected(x) = True                   ' Set select
      End If
    Next
    GetExportTypes "MODEL"
    If ListBoxExport.ListCount = 0 Then
      AddExportType "STEP", ".step", "MULTI"
      AddExportType "IGES", ".igs", "MULTI"
    End If
    ListBoxExport.Selected(0) = True
  Else                                  ' -- SolidWorks Drawings --
    ListBoxUserSelect.Clear                                    ' Clear form
    numSheets = Document.GetSheetCount()                       ' Num sheets
    SheetNames = Document.GetSheetNames()                      ' Configs Names
    For x = 0 To (numSheets - 1)                               ' Each config
      ListBoxUserSelect.AddItem SheetNames(x)                  ' Add to list
    Next
    Set Sheet = Document.GetCurrentSheet                       ' Active sheet
    CurrentSheetName = Sheet.GetName                           ' name
    For x = 0 To (ListBoxUserSelect.ListCount - 1)             ' Each config
      If ListBoxUserSelect.List(x, 0) = CurrentConfigName Then ' Active conf?
        ListBoxUserSelect.Selected(x) = True                   ' Set select
      End If
    Next
    GetExportTypes "DRAWING"
    If ListBoxExport.ListCount = 0 Then
      AddExportType "DWG vector", ".dwg", "MULTI"
      AddExportType "JPEG raster", ".jpg", "MULTI"
    End If
    ListBoxExport.Selected(0) = True
    CheckBoxFilter.Enabled = False                             ' Disable
    TextBoxFilter.Enabled = False                              ' all filter
    CommandSelect.Enabled = False                              ' tools for
    CommandClear.Enabled = False                               ' drawings
  End If
  ' -- All SolidWorks Documents --
  CheckBoxFilter = False
  CommandClearAll_Click
  CheckBoxFilter_Click
  LabelStatus.Caption = ""                                     ' Clear status
  CommandClose.SetFocus
End Sub
