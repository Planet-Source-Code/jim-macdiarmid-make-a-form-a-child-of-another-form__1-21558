VERSION 5.00
Begin VB.Form DlgLoginParent 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "DlgLoginParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    '*** DO NOT Close the Parent form without getting rid of the Child form, otherwise the app could crash.
    Unload DlgChildLogin
    End
End Sub
