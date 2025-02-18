VERSION 5.00
Begin VB.Form frmStuff 
   Caption         =   "Git Hub Test"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdItalic 
      Caption         =   "Italic"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdBoldIt 
      Caption         =   "Bold"
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtText 
      Height          =   855
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objStuff As clsStuff


Private Sub cmdBoldIt_Click()
    If txtText.FontBold Then
        txtText.FontBold = False
    Else
        txtText.FontBold = True
    End If
End Sub

Private Sub cmdItalic_Click()
    If txtText.FontItalic Then
        txtText.FontItalic = False
    Else
        txtText.FontItalic = True
    End If
End Sub

Private Sub Form_Load()
    Set objStuff = New clsStuff
    
End Sub
