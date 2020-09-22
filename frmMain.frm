VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HTML Filter"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton remhtml 
      Caption         =   "Remove HTML Tags"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   5895
   End
   Begin VB.TextBox htmlsource 
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The SubString Function is the same as it comes
'in Javascript language
'It is useful for getting a sub string of a string
'starting in a detrmined position
'and ending in a determines position

Function SubString(str, atstart, atend)
Dim retvalue As String
retvalue = Mid(str, atstart, (atend - atstart) + 1)
SubString = retvalue
End Function

Function RemoveHTML(txt, fc, lc, tr)
'txt is the html content
'fc character is the tag first character like "<" in <html> or "&" in &nbsp;
'lc is the tag last character like ">" in </html> or ";" in &copy;
'tr is the tag replacer text, for example it could be a space symbol " "
'IMPORTANT: Do not set the tr variable empty like "", otherwise html tags will not dissapear
Dim nextpos
'lc position variable
For jk = 1 To Len(txt)
If Mid(txt, jk, 1) = fc Then
'If current character is equal to lc then do the next
'Get the lc position
nextpos = InStr(jk, txt, lc)
'If the lc is founded then do tha tag removal process
If nextpos <> 0 Then
'Replace the tag with the tag replacer text
txt = Left(txt, jk - 1) & tr & Mid(txt, (nextpos + 1), ((Len(txt) - nextpos) + 1))
End If
End If
Next
RemoveHTML = txt
End Function

Private Sub remhtml_Click()
Dim newsource As String
Dim finalsource As String
Dim htmlsectors As Variant
Dim htmlsectors2 As Variant
'Remove primary html tags like <body>, <pre>, <script>, etc
newsource = RemoveHTML(htmlsource.Text, "<", ">", " ")
'Remove secondary html tags like &nbsp;, &lt;, &gt;, &quot;, &copy;, etc
finalsource = RemoveHTML(newsource, "&", ";", " ")
htmlsource.Text = finalsource
End Sub
