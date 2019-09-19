VERSION 5.00
Begin VB.Form frmBJ 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BlackJack"
   ClientHeight    =   10545
   ClientLeft      =   4875
   ClientTop       =   1575
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   14475
   Begin VB.PictureBox pctBotones 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   6240
      ScaleHeight     =   915
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   5400
      Width           =   5535
      Begin VB.CommandButton cmdDoble 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Doble"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPasar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pasar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPedir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pedir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRepartir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Repartir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   1215
      Left            =   -120
      Picture         =   "frmBJ.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   14595
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H8000000E&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAyuda 
         BackColor       =   &H8000000E&
         Caption         =   "Ayuda"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuevo 
         BackColor       =   &H8000000E&
         Caption         =   "Nuevo"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Image imgMazo 
      Height          =   2280
      Index           =   6
      Left            =   12480
      Picture         =   "frmBJ.frx":19CB
      Top             =   1800
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgMazo 
      Height          =   2280
      Index           =   5
      Left            =   11160
      Picture         =   "frmBJ.frx":35F4
      Top             =   2400
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgMazo 
      Height          =   2280
      Index           =   4
      Left            =   9840
      Picture         =   "frmBJ.frx":521D
      Top             =   1800
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgMazo 
      Height          =   2280
      Index           =   3
      Left            =   8520
      Picture         =   "frmBJ.frx":6E46
      Top             =   2400
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgMazo 
      Height          =   2280
      Index           =   2
      Left            =   7200
      Picture         =   "frmBJ.frx":8A6F
      Top             =   1800
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgMazo 
      Height          =   2280
      Index           =   1
      Left            =   5880
      Picture         =   "frmBJ.frx":A698
      Top             =   2400
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgJugador 
      Height          =   2295
      Index           =   6
      Left            =   12480
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Image imgJugador 
      Height          =   2295
      Index           =   5
      Left            =   11160
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Image imgJugador 
      Height          =   2295
      Index           =   4
      Left            =   9840
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Image imgJugador 
      Height          =   2295
      Index           =   3
      Left            =   8520
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Image imgJugador 
      Height          =   2295
      Index           =   2
      Left            =   7200
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Image imgJugador 
      Height          =   2295
      Index           =   1
      Left            =   5880
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Image imgJugador 
      Height          =   2295
      Index           =   0
      Left            =   4560
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Image imgBanca 
      Height          =   2295
      Index           =   6
      Left            =   12480
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Image imgMazo 
      Height          =   2280
      Index           =   0
      Left            =   4560
      Picture         =   "frmBJ.frx":C2C1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgKt 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":DEEA
      Top             =   7680
      Width           =   1680
   End
   Begin VB.Image imgKp 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":FAD8
      Top             =   7680
      Width           =   1680
   End
   Begin VB.Image imgKd 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":11731
      Top             =   7680
      Width           =   1680
   End
   Begin VB.Image imgQt 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":1343B
      Top             =   7080
      Width           =   1680
   End
   Begin VB.Image imgQp 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":1518A
      Top             =   7080
      Width           =   1680
   End
   Begin VB.Image imgQd 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":16F76
      Top             =   7080
      Width           =   1680
   End
   Begin VB.Image imgJp 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":18D4A
      Top             =   6480
      Width           =   1680
   End
   Begin VB.Image imgJt 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":1AA64
      Top             =   6480
      Width           =   1680
   End
   Begin VB.Image imgJd 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":1C74B
      Top             =   6480
      Width           =   1680
   End
   Begin VB.Image img10p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":1E4A1
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Image img10d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":1FA59
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Image img10t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":20FA9
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Image img9t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":226C5
      Top             =   5400
      Width           =   1680
   End
   Begin VB.Image img9d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":23C72
      Top             =   5400
      Width           =   1680
   End
   Begin VB.Image img9p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":250B6
      Top             =   5400
      Width           =   1680
   End
   Begin VB.Image img8t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":2650E
      Top             =   4680
      Width           =   1680
   End
   Begin VB.Image img8d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":2797F
      Top             =   4680
      Width           =   1680
   End
   Begin VB.Image img8p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":28C81
      Top             =   4680
      Width           =   1680
   End
   Begin VB.Image img7t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":2A078
      Top             =   3960
      Width           =   1680
   End
   Begin VB.Image img7d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":2B3A0
      Top             =   3840
      Width           =   1680
   End
   Begin VB.Image img7p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":2C5B6
      Top             =   3960
      Width           =   1680
   End
   Begin VB.Image img6t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":2D8D2
      Top             =   3240
      Width           =   1680
   End
   Begin VB.Image img6d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":2EB24
      Top             =   3240
      Width           =   1680
   End
   Begin VB.Image img6p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":2FC8F
      Top             =   3240
      Width           =   1680
   End
   Begin VB.Image img5t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":30E96
      Top             =   2640
      Width           =   1680
   End
   Begin VB.Image img5d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":31FD0
      Top             =   2640
      Width           =   1680
   End
   Begin VB.Image img5p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":33018
      Top             =   2520
      Width           =   1680
   End
   Begin VB.Image img4p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":340EE
      Top             =   1800
      Width           =   1680
   End
   Begin VB.Image img4d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":350A6
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Image img4t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":36001
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Image img3t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":37026
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Image img3p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":37F3D
      Top             =   1200
      Width           =   1680
   End
   Begin VB.Image img3d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":38DFF
      Top             =   1200
      Width           =   1680
   End
   Begin VB.Image img2t 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":39C17
      Top             =   840
      Width           =   1680
   End
   Begin VB.Image imgKc 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":3A9ED
      Top             =   7680
      Width           =   1680
   End
   Begin VB.Image imgQc 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":3C788
      Top             =   7200
      Width           =   1680
   End
   Begin VB.Image imgJc 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":3E50E
      Top             =   6720
      Width           =   1680
   End
   Begin VB.Image img10c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":401C3
      Top             =   6120
      Width           =   1680
   End
   Begin VB.Image img9c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":419B2
      Top             =   5520
      Width           =   1680
   End
   Begin VB.Image img8c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":4305C
      Top             =   4800
      Width           =   1680
   End
   Begin VB.Image img7c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":445F2
      Top             =   4200
      Width           =   1680
   End
   Begin VB.Image img6c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":45A02
      Top             =   3480
      Width           =   1680
   End
   Begin VB.Image img5c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":46D2C
      Top             =   2880
      Width           =   1680
   End
   Begin VB.Image img2p 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":47EDF
      Top             =   720
      Width           =   1680
   End
   Begin VB.Image img4c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":48CBF
      Top             =   2160
      Width           =   1680
   End
   Begin VB.Image img2d 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":49D52
      Top             =   720
      Width           =   1680
   End
   Begin VB.Image img3c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":4AABB
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Image imgAt 
      Height          =   2295
      Left            =   20400
      Picture         =   "frmBJ.frx":4BA7B
      Top             =   240
      Width           =   1680
   End
   Begin VB.Image imgAp 
      Height          =   2295
      Left            =   16320
      Picture         =   "frmBJ.frx":4C72C
      Top             =   240
      Width           =   1680
   End
   Begin VB.Image imgAd 
      Height          =   2295
      Left            =   18360
      Picture         =   "frmBJ.frx":4D3E7
      Top             =   240
      Width           =   1680
   End
   Begin VB.Image img2c 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":4E06B
      Top             =   840
      Width           =   1680
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   14640
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image imgAc 
      Height          =   2295
      Left            =   14520
      Picture         =   "frmBJ.frx":4EF23
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label lblPuntosJ 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label lblPuntosB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label lblJugador 
      BackStyle       =   0  'Transparent
      Caption         =   "Jugador:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label lblBanca 
      BackStyle       =   0  'Transparent
      Caption         =   "Repartidor:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Image imgBanca 
      Height          =   2295
      Index           =   5
      Left            =   11160
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Image imgBanca 
      Height          =   2295
      Index           =   4
      Left            =   9840
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Image imgBanca 
      Height          =   2295
      Index           =   3
      Left            =   8520
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Image imgBanca 
      Height          =   2295
      Index           =   2
      Left            =   7200
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Image imgBanca 
      Height          =   2295
      Index           =   1
      Left            =   5880
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Image imgBanca 
      Height          =   2295
      Index           =   0
      Left            =   4560
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Menu mnuJuego 
      Caption         =   "Juego"
      Begin VB.Menu mnuNuevo 
         Caption         =   "Nuevo Juego"
      End
      Begin VB.Menu mnuEstadistica 
         Caption         =   "Estadisticas del juego"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuReglas 
         Caption         =   "Reglas del juego"
      End
   End
End
Attribute VB_Name = "frmBJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNuevo_Click()

    nombre = InputBox("Ingrese su nombre")  'Pido al usuario que ingrese su nombre
    
    While nombre = ""      'Mientras que el usuario no ingrese su nombre no puede continar el juego
        nombre = InputBox("Ingrese su nombre")
    Wend
    
    lblJugador.Caption = nombre & ":"    'Mando el nombre a el lblJugador
    cmdRepartir.Enabled = True           'Hago que el boton repartir este habilitado

End Sub

Private Sub cmdPasar_Click()
Dim i As Integer

    Call PartidaDeBanca        'Llamo al sub PartidaDeBanca
    
    For i = 0 To 6             'Hago que las cartas de la banca esten visibles ocultando la imgMazo
        imgMazo(i).Visible = False
    Next i
    
    lblPuntosB.Caption = puntosBtotal 'Hago que se muestren los puntos de la banca
    
    
    If (((puntosJtotal > 21) And (puntosBtotal > 21)) Or (puntosJtotal = puntosBtotal)) Then     'Si los puntos del jugador y de la banca superan los 21 o la banca y el jugador tienen el mismo puntaje entonces es un empate
            MsgBox "¡Empate!", vbOKOnly, "BlackJack"
    Else
        If (puntosBtotal > 21) Then   'Si los puntos de la banca superan los 21 entonces el jugador gano
            MsgBox "¡Felicidades, Ganaste!", vbOKOnly, "BlackJack"
        Else
            If (puntosJtotal < puntosBtotal) Then   'si los puntos del jugador son menores que los de la banca el jugador perdio
                MsgBox "¡Que lastima, Perdiste!", vbOKOnly, "BlackJack"
            Else   'Si los puntos del jugador son mayores que la banca entonces el jugador gano
                MsgBox "¡Felicidades, Ganaste!", vbOKOnly, "BlackJack"
            End If
        End If
    End If
    'deshabilito cmdPedir, cmdPasar, cmdDoble, dejando solo el boton Repartir
    cmdPedir.Enabled = False
    cmdPasar.Enabled = False
    cmdDoble.Enabled = False

End Sub

Private Sub cmdPedir_Click()
    
    cmdDoble.Enabled = False    'deshabilito el boton Doble
    
    If puntosJtotal < 21 Then    'Cuando se pide una carta se pregunta si los puntos del jugador es menor que 21
        If imgJugador(2).Visible = False Then   'si el puntaje es menor que 21 y la tercera carta no esta visible entonces se la muestra y se suma el puntaje
            imgJugador(2).Visible = True
            puntosJtotal = puntosJtotal + puntosJ(2)
            lblPuntosJ.Caption = puntosJtotal
        Else
            If imgJugador(3).Visible = False Then   'si la cuarta carta no esta visible entonces se la muestra y se suma el puntaje
                imgJugador(3).Visible = True
                puntosJtotal = puntosJtotal + puntosJ(3)
                lblPuntosJ.Caption = puntosJtotal
            Else
                If imgJugador(4).Visible = False Then   'si la quinta carta no esta visible entonces se la muestra y se suma el puntaje
                    imgJugador(4).Visible = True
                    puntosJtotal = puntosJtotal + puntosJ(4)
                    lblPuntosJ.Caption = puntosJtotal
                Else
                    If imgJugador(4).Visible = False Then   'si la sexta carta no esta visible entonces se la muestra y se suma el puntaje
                        imgJugador(4).Visible = True
                        puntosJtotal = puntosJtotal + puntosJ(5)
                        lblPuntosJ.Caption = puntosJtotal
                    Else
                        If imgJugador(4).Visible = False Then   'si la septima carta no esta visible entonces se la muestra y se suma el puntaje
                            imgJugador(4).Visible = True
                            puntosJtotal = puntosJtotal + puntosJ(6)
                            lblPuntosJ.Caption = puntosJtotal
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Call PartidaDeBanca     'Se llama al sub PartidaDeBanca
    
    If (puntosJtotal > 21) Then    'Despues de haber pedido la carta se pregunta si el puntaje es mayor de 21 y si lo es se muestra un msgbox diciendo que perdio
        MsgBox "Tienes " & Val(lblPuntosJ.Caption) & ", has perdido", vbOKOnly, "BlackJack"
        cmdPedir.Enabled = False        'se deshabilitan los botones pedir, pasar y doble
        cmdPasar.Enabled = False
        cmdDoble.Enabled = False
        For i = 0 To 6      'Se muestran las cartas de la banca
            imgMazo(i).Visible = False
        Next i
        cmdRepartir.SetFocus                'Se pasa el foco a el comando Repartir
        lblPuntosB.Caption = puntosBtotal   'Se muestran los puntos de la banca
        ReDim puntajesJ(partidas)
        puntajes = puntosJtotal            'Se carga al arreglo con los puntos para despues mostrarlo en frmEstadisticas
    End If

End Sub

Private Sub cmdRepartir_Click()
Dim i As Integer
    
    partidas = partidas + 1
    lblPuntosB.Caption = ""     'Se borra el puntaje de la banca
    GenerarCartas               'se llama al sub GenerarCartas
    
    cmdPedir.Enabled = True     'Se habilitan los botones pedir, pasar y doble
    cmdPasar.Enabled = True
    cmdDoble.Enabled = True
    
    imgMazo(0).Visible = True   'Se tapan las cartas de la banca
    imgMazo(1).Visible = True
    
    For i = 2 To 6
        imgBanca(i).Visible = False     'Se repartieron 7 cartas a la banca y al jugador, de las cuales se ocultan 5
        imgJugador(i).Visible = False
    Next i
    puntosJtotal = puntosJ(0) + puntosJ(1)   'Se carga la variable con el puntaje obtenido hasta ahora por el jugador
    lblPuntosJ.Caption = puntosJtotal     'Se muestra el puntaje del Jugador

End Sub

Private Sub cmdSalir_Click()

    Unload Me   'Se descarga el formulario

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If (MsgBox("¿Está seguro de que desea salir?", vbOKCancel, "Salir") = vbOK) Then
        Cancel = 0
    Else
        Cancel = 1
    End If
    
End Sub

Private Sub mnuEstadistica_Click()

frmEstadisticas.Show

End Sub

Private Sub mnuNuevo_Click()

Call cmdNuevo_Click   'Se llama el sub del boton Nuevo

End Sub

Private Sub mnuReglas_Click()

frmReglas.Show

End Sub

Private Sub mnuSalir_Click()

    Unload Me   'Se descarga el formulario
    
End Sub
