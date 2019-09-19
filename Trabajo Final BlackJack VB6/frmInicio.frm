VERSION 5.00
Begin VB.Form frmInicio 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicio"
   ClientHeight    =   7095
   ClientLeft      =   8460
   ClientTop       =   3075
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInicio.frx":0000
   ScaleHeight     =   7095
   ScaleWidth      =   7470
   Begin VB.Timer tmrInicio 
      Interval        =   1200
      Left            =   6960
      Top             =   240
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tmrInicio_Timer()

Unload Me
frmBJ.Show

End Sub
