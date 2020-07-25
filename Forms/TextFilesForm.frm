VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TextFilesForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "TextFilesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TextFilesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Set FSO = CreateObject("Scripting.FileSystemObject")
'ActiveDocument.Name
'ActiveDocument.FullName
'ActiveDocument.Path
    S = ActiveDocument.Path
    MsgBox S
    TextPath = S & "\Test.txt"
    Set TextStream = FSO.CreateTextFile(TextPath) 'Создать файл
    CreateObject("wscript.shell").Run """" & TextPath & """" 'открыть в родной программе
End Sub

