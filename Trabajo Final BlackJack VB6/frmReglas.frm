VERSION 5.00
Begin VB.Form frmReglas 
   BackColor       =   &H00004000&
   Caption         =   "Reglas del juego"
   ClientHeight    =   4995
   ClientLeft      =   10155
   ClientTop       =   3150
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   3960
   Begin VB.CommandButton cmdVolver 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "frmReglas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVolver_Click()

Unload Me

End Sub
