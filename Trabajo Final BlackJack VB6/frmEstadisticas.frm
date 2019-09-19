VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00004000&
   Caption         =   "Estadisticas del juego"
   ClientHeight    =   4020
   ClientLeft      =   9795
   ClientTop       =   3330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4560
   Begin VB.CommandButton cmdVolver 
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   855
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVolver_Click()

Unload Me

End Sub
