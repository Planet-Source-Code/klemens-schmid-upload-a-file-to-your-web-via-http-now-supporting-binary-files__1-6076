VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Upload File"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This code uploads a file to an ASP script using http post
'You need to add a reference to Microsoft XML 3.0 to the project.
'Get this from http://msdn.microsoft.com/xml

Private Sub cmdSubmit_Click()
Dim strText As String
Dim strImage As String
Dim s$
Dim strBody As String
Dim aPostData() As Byte
Dim strFileName1 As String
Dim strFileName2 As String
Dim oHttp As XMLHTTP
Dim nFile As Integer

'make use of the XMLHTTPRequest object contained in msxml.dll
Set oHttp = New XMLHTTP
'read the whole text file
strFileName1 = InputBox("Text File Name:", "Upload a file", "c:\temp\test.txt")
nFile = FreeFile
Open strFileName1 For Binary As #nFile
strText = String(LOF(nFile), " ")
Get #nFile, , strText
Close #nFile
'read a GIF file
strFileName2 = InputBox("Image File Name:", "Upload a file", "c:\temp\test.gif")
nFile = FreeFile
Open strFileName2 For Binary As #nFile
strImage = String(LOF(nFile), " ")
Get #nFile, , strImage
Close #nFile
'fire of an http request
!!! adapt path in next line
oHttp.open "POST", "http://localhost/upload.asp", False
oHttp.setRequestHeader "Content-Type", "multipart/form-data, boundary=AaB03x"
'assemble the body. send one field and two files
strBody = _
   "--AaB03x" & vbCrLf & _
   "Content-Disposition: form-data; name=""field1""" & vbCrLf & vbCrLf & _
   "test field" & vbCrLf & _
   "--AaB03x" & vbCrLf & _
   "Content-Disposition: attachment; name=""FILE1""; filename=""" & strFileName1 & """" & vbCrLf & _
   "Content-Type: text/plain" & vbCrLf & vbCrLf & _
   strText & vbCrLf & _
   "--AaB03x" & vbCrLf & _
   "Content-Disposition: attachment; name=""FILE2""; filename=""" & strFileName2 & """" & vbCrLf & _
   "Content-Transfer-Encoding: binary" & vbCrLf & _
   "Content-Type: image/gif" & vbCrLf & vbCrLf & _
   strImage & vbCrLf & _
   "--AaB03x--"
'must convert to byte array because of binary zeros
aPostData = StrConv(strBody, vbFromUnicode)
'send it
oHttp.send aPostData
'check the feedback
Debug.Print oHttp.responseText

End Sub


