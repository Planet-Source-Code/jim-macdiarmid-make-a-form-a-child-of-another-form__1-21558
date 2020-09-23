VERSION 5.00
Begin VB.Form DlgChildLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login Dialog"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Login"
      Height          =   420
      Left            =   4860
      TabIndex        =   2
      Top             =   285
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   540
      TabIndex        =   1
      Top             =   1425
      Width           =   3930
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   540
      TabIndex        =   0
      Top             =   915
      Width           =   3930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Add some text here"
      Height          =   195
      Left            =   525
      TabIndex        =   3
      Top             =   345
      Width           =   1365
   End
End
Attribute VB_Name = "DlgChildLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


