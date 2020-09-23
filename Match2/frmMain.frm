VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Match2 - designed by Tim Harpur"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   574
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  gameState = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If gameState <> -2 Then Cancel = 1
End Sub

