Attribute VB_Name = "DocExportMain"
' --------------------------------------------------------------------03/09/2007
' DocExport.swp                     Written by Leonard Kikstra,
'                                   Copyright 2003-2007, Leonard Kikstra
'                                   Downloaded from Lenny's SolidWorks Resources
'                                        at http://www.lennyworks.com/solidworks
' ------------------------------------------------------------------------------
Global swApp As Object
Global Document As Object
Global Configuration As Object
Global FileTyp As String
Global numConfigs As Integer
Global ConfigNames As Variant
Global Retval As Integer
Global Errors As Long
Global Warnings As Long
Global DocMode As Integer
' SolidWorks constants
Global Const swDocPART = 1
Global Const swDocASSEMBLY = 2
Global Const swDocDRAWING = 3
Global Const swSaveAsCurrentVersion = 0
Global Const swSaveAsOptions_Silent = &H1
Global Const Version = "v1.30"

Sub Main()
  Set swApp = CreateObject("SldWorks.Application")            ' Attach to SWX
  Set Document = swApp.ActiveDoc                              ' Grab active doc
  If Document Is Nothing Then                                 ' Is doc loaded
    FormDocExportBatch.Show                                   ' go batch mode
  Else                                                        ' Doc loaded?
    FileTyp = Document.GetType                                ' Get doc type
    If FileTyp = swDocASSEMBLY Or FileTyp = swDocPART Then    ' Model ?
      DocMode = 0
      FormDocExport.Show
    ElseIf FileTyp = swDocDRAWING Then                        ' Drawing ?
      DocMode = 1
      FormDocExport.Show
    Else                                                      ' Else doc type
      MsgBox "Active file is not a SolidWorks document.", vbExclamation
    End If                                                    ' End doc type
  End If                                                      ' End doc load
End Sub
