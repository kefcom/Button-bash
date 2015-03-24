VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Starting..."
   ClientHeight    =   150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1740
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   150
   ScaleWidth      =   1740
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Screen.Width > 15360 Then
    Form1.BorderStyle = 1
Else
    If Screen.Width = 15360 Then
        Form1.BorderStyle = 0
        Form1.Caption = Form1.Caption
        Form1.Width = Screen.Width
        Form1.Height = Screen.Height
        Form1.Left = 0
        Form1.Top = 0
    End If
    If Screen.Width < 15360 Then
        MsgBox "This game can not run in this resolution, please set to at least 1024*768", vbCritical, "Error"
        End
    End If
End If
Form1.Show
Unload Form2
End Sub
