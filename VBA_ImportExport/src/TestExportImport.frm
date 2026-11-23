VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestExportImport 
   Caption         =   "Test Form"
   ClientHeight    =   1305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "TestExportImport.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "TestExportImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Start_Click()
    MsgBox "Hallo World", vbInformation, "Mein erstes Makro!"
End Sub
