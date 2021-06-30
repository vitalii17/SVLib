VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAbout 
   Caption         =   "DocExport: ABOUT"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   OleObjectBlob   =   "FormAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAbout"
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
' FormAbout:         Tell user about this macro and Lenny's SolidWorks Resources
' ------------------------------------------------------------------------------

Private Sub CommandClose_Click()
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
  ProgramVersion.Caption = "Version: v" & Version
End Sub
