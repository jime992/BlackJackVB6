Attribute VB_Name = "Module1"
Option Explicit

Public nombre As String
Public cartasJ() As String
Public cartasB() As String
Public cartas() As Integer
Public puntosJ() As Integer
Public puntosB() As Integer
Public puntosBtotal As Integer
Public Aux() As Integer
Public puntajesJ() As Integer
Public puntosJtotal As Integer
Public partidas As Integer





Public Sub GenerarCartas()
Dim i As Integer
ReDim cartasJ(6)
ReDim cartasB(6)
ReDim puntosJ(6)
ReDim puntosB(6)
ReDim cartas(13)

    cartas = PedirNumeros(Val(14))  'Pido 14 numeros entre el 1 y el 52

    cartasJ(0) = cartas(1)   'le asigno el primer numero que se obtuvo al jugador
    cartasB(0) = cartas(2)   'le asigno el segunto numero que se obtuvo a la banca
    cartasJ(1) = cartas(3)   'le asigno el tercero numero que se obtuvo al jugador
    cartasB(1) = cartas(4)   'le asigno el cuarto numero que se obtuvo a la banca
    cartasJ(2) = cartas(5)   'le asigno el quinto numero que se obtuvo al jugador
    cartasB(2) = cartas(6)   'le asigno el sexto numero que se obtuvo a la banca
    cartasJ(3) = cartas(7)   'le asigno el septimo numero que se obtuvo al jugador
    cartasB(3) = cartas(8)   'le asigno el octavo numero que se obtuvo a la banca
    cartasJ(4) = cartas(9)   'le asigno el noveno numero que se obtuvo al jugador
    cartasB(4) = cartas(10)  'le asigno el décimo numero que se obtuvo a la banca
    cartasJ(5) = cartas(11)  'le asigno el unodécimo numero que se obtuvo al jugador
    cartasB(5) = cartas(12)  'le asigno el duodécimo numero que se obtuvo a la banca
    cartasJ(6) = cartas(13)  'le asigno el decimotercer numero que se obtuvo al jugador
    cartasB(6) = cartas(14)  'le asigno el decimocuarto numero que se obtuvo a la banca
    
     For i = 0 To UBound(cartasJ())
        
        Select Case cartasJ(i)      'segun el numero que le toco al jugador le toca un puntaje y la imagen de la carta que le toco
            
            Case 1                  'carta A de corazones
                puntosJ(i) = 11
                frmBJ.imgJugador(i).Picture = frmBJ.imgAc.Picture
            
            Case 2                  'carta 2 de corazones
                puntosJ(i) = 2
                frmBJ.imgJugador(i).Picture = frmBJ.img2c.Picture
            
            Case 3                  'carta 3 de corazones
                puntosJ(i) = 3
                frmBJ.imgJugador(i).Picture = frmBJ.img3c.Picture
            
            Case 4                  'carta 4 de corazones
                puntosJ(i) = 4
                frmBJ.imgJugador(i).Picture = frmBJ.img4c.Picture
            
            Case 5                  'carta 5 de corazones
                puntosJ(i) = 5
                frmBJ.imgJugador(i).Picture = frmBJ.img5c.Picture
            
            Case 6                  'carta 6 de corazones
                puntosJ(i) = 6
                frmBJ.imgJugador(i).Picture = frmBJ.img6c.Picture
            
            Case 7                  'carta 7 de corazones
                puntosJ(i) = 7
                frmBJ.imgJugador(i).Picture = frmBJ.img7c.Picture
            
            Case 8                  'carta 8 de corazones
                puntosJ(i) = 8
                frmBJ.imgJugador(i).Picture = frmBJ.img8c.Picture
            
            Case 9                  'carta 9 de corazones
                puntosJ(i) = 9
                frmBJ.imgJugador(i).Picture = frmBJ.img9c.Picture
            
            Case 10                 'carta 10 de corazones
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.img10c.Picture
            
            Case 11                 'carta J de corazones
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgJc.Picture
            
            Case 12                 'carta Q de corazones
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgQc.Picture
            
            Case 13                 'carta K de corazones
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgKc.Picture
            
            Case 14                 'carta A de trebol
                puntosJ(i) = 11
                frmBJ.imgJugador(i).Picture = frmBJ.imgAt.Picture
            
            Case 15                 'carta 2 de trebol
                puntosJ(i) = 2
                frmBJ.imgJugador(i).Picture = frmBJ.img2t.Picture
            
            Case 16                 'carta 3 de trebol
                puntosJ(i) = 3
                frmBJ.imgJugador(i).Picture = frmBJ.img3t.Picture
            
            Case 17                 'carta 4 de trebol
                puntosJ(i) = 4
                frmBJ.imgJugador(i).Picture = frmBJ.img4t.Picture
            
            Case 18                 'carta 5 de trebol
                puntosJ(i) = 5
                frmBJ.imgJugador(i).Picture = frmBJ.img5t.Picture
            
            Case 19                 'carta 6 de trebol
                puntosJ(i) = 6
                frmBJ.imgJugador(i).Picture = frmBJ.img6t.Picture
            
            Case 20                 'carta 7 de trebol
                puntosJ(i) = 7
                frmBJ.imgJugador(i).Picture = frmBJ.img7t.Picture
            
            Case 21                 'carta 8 de trebol
                puntosJ(i) = 8
                frmBJ.imgJugador(i).Picture = frmBJ.img8t.Picture
            
            Case 22                 'carta 9 de trebol
                puntosJ(i) = 9
                frmBJ.imgJugador(i).Picture = frmBJ.img9t.Picture
            
            Case 23                 'carta 10 de trebol
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.img10t.Picture
            
            Case 24                 'carta J de trebol
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgJt.Picture
            
            Case 25                 'carta Q de trebol
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgQt.Picture
            
            Case 26                 'carta K de trebol
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgKt.Picture
            
            Case 27                 'carta A de pica
                puntosJ(i) = 11
                frmBJ.imgJugador(i).Picture = frmBJ.imgAp.Picture
            
            Case 28                 'carta 2 de pica
                puntosJ(i) = 2
                frmBJ.imgJugador(i).Picture = frmBJ.img2p.Picture
            
            Case 29                 'carta 3 de pica
                puntosJ(i) = 3
                frmBJ.imgJugador(i).Picture = frmBJ.img3p.Picture
            
            Case 30                 'carta 4 de pica
                puntosJ(i) = 4
                frmBJ.imgJugador(i).Picture = frmBJ.img4p.Picture
            
            Case 31                 'carta 5 de pica
                puntosJ(i) = 5
                frmBJ.imgJugador(i).Picture = frmBJ.img5p.Picture
            
            Case 32                 'carta 6 de pica
                puntosJ(i) = 6
                frmBJ.imgJugador(i).Picture = frmBJ.img6p.Picture
            
            Case 33                 'carta 7 de pica
                puntosJ(i) = 7
                frmBJ.imgJugador(i).Picture = frmBJ.img7p.Picture
            
            Case 34                 'carta 8 de pica
                puntosJ(i) = 8
                frmBJ.imgJugador(i).Picture = frmBJ.img8p.Picture
            
            Case 35                 'carta 9 de pica
                puntosJ(i) = 9
                frmBJ.imgJugador(i).Picture = frmBJ.img9p.Picture
            
            Case 36                 'carta 10 de pica
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.img10p.Picture
            
            Case 37                 'carta J de pica
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgJp.Picture
            
            Case 38                 'carta Q de pica
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgQp.Picture
            
            Case 39                 'carta K de pica
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgKp.Picture
            
            Case 40                 'carta A de diamante
                puntosJ(i) = 11
                frmBJ.imgJugador(i).Picture = frmBJ.imgAd.Picture
            
            Case 41                 'carta 2 de diamante
                puntosJ(i) = 2
                frmBJ.imgJugador(i).Picture = frmBJ.img2d.Picture
            
            Case 42                 'carta 3 de diamante
                puntosJ(i) = 3
                frmBJ.imgJugador(i).Picture = frmBJ.img3d.Picture
            
            Case 43                 'carta 4 de diamante
                puntosJ(i) = 4
                frmBJ.imgJugador(i).Picture = frmBJ.img4d.Picture
            
            Case 44                 'carta 5 de diamante
                puntosJ(i) = 5
                frmBJ.imgJugador(i).Picture = frmBJ.img5d.Picture
            
            Case 45                 'carta 6 de diamante
                puntosJ(i) = 6
                frmBJ.imgJugador(i).Picture = frmBJ.img6d.Picture
            
            Case 46                 'carta 7 de diamante
                puntosJ(i) = 7
                frmBJ.imgJugador(i).Picture = frmBJ.img7d.Picture
            
            Case 47                 'carta 8 de diamante
                puntosJ(i) = 8
                frmBJ.imgJugador(i).Picture = frmBJ.img8d.Picture
            
            Case 48                 'carta 9 de diamante
                puntosJ(i) = 9
                frmBJ.imgJugador(i).Picture = frmBJ.img9d.Picture
                 
            Case 49                 'carta 10 de diamante
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.img10d.Picture
            
            Case 50                 'carta J de diamante
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgJd.Picture
            
            Case 51                 'carta Q de diamante
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgQd.Picture
    
            Case 52                 'carta K de diamante
                puntosJ(i) = 10
                frmBJ.imgJugador(i).Picture = frmBJ.imgKd.Picture
            
        End Select
    Next i
    
     For i = 0 To UBound(cartasB())
        
        Select Case cartasB(i)      'segun la carta que le toco a la banca le toca un puntaje y la imagen de la carta que le toco
            
            Case 1  'carta A de corazones
                puntosB(i) = 11
                frmBJ.imgBanca(i).Picture = frmBJ.imgAc.Picture
            
            Case 2  'carta 2 de corazones
                puntosB(i) = 2
                frmBJ.imgBanca(i).Picture = frmBJ.img2c.Picture
            
            Case 3  'carta 3 de corazones
                puntosB(i) = 3
                frmBJ.imgBanca(i).Picture = frmBJ.img3c.Picture
            
            Case 4  'carta 4 de corazones
                puntosB(i) = 4
                frmBJ.imgBanca(i).Picture = frmBJ.img4c.Picture
            
            Case 5  'carta 5 de corazones
                puntosB(i) = 5
                frmBJ.imgBanca(i).Picture = frmBJ.img5c.Picture
            
            Case 6  'carta 6 de corazones
                puntosB(i) = 6
                frmBJ.imgBanca(i).Picture = frmBJ.img6c.Picture
            
            Case 7  'carta 7 de corazones
                puntosB(i) = 7
                frmBJ.imgBanca(i).Picture = frmBJ.img7c.Picture
            
            Case 8  'carta 8 de corazones
                puntosB(i) = 8
                frmBJ.imgBanca(i).Picture = frmBJ.img8c.Picture
            
            Case 9  'carta 9 de corazones
                puntosB(i) = 9
                frmBJ.imgBanca(i).Picture = frmBJ.img9c.Picture
            
            Case 10 'carta 10 de corazones
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.img10c.Picture
            
            Case 11 'carta J de corazones
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgJc.Picture
            
            Case 12 'carta Q de corazones
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgQc.Picture
            
            Case 13 'carta K de corazones
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgKc.Picture
            
            Case 14 'carta A de trebol
                puntosB(i) = 11
                frmBJ.imgBanca(i).Picture = frmBJ.imgAt.Picture
            
            Case 15 'carta 2 de trebol
                puntosB(i) = 2
                frmBJ.imgBanca(i).Picture = frmBJ.img2t.Picture
            
            Case 16 'carta 3 de trebol
                puntosB(i) = 3
                frmBJ.imgBanca(i).Picture = frmBJ.img3t.Picture
            
            Case 17 'carta 4 de trebol
                puntosB(i) = 4
                frmBJ.imgBanca(i).Picture = frmBJ.img4t.Picture
            
            Case 18 'carta 5 de trebol
                puntosB(i) = 5
                frmBJ.imgBanca(i).Picture = frmBJ.img5t.Picture
            
            Case 19 'carta 6 de trebol
                puntosB(i) = 6
                frmBJ.imgBanca(i).Picture = frmBJ.img6t.Picture
            
            Case 20 'carta 7 de trebol
                puntosB(i) = 7
                frmBJ.imgBanca(i).Picture = frmBJ.img7t.Picture
            
            Case 21 'carta 8 de trebol
                puntosB(i) = 8
                frmBJ.imgBanca(i).Picture = frmBJ.img8t.Picture
            
            Case 22 'carta 9 de trebol
                puntosB(i) = 9
                frmBJ.imgBanca(i).Picture = frmBJ.img9t.Picture
            
            Case 23 'carta 10 de trebol
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.img10t.Picture
            
            Case 24 'carta J de trebol
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgJt.Picture
            
            Case 25 'carta Q de trebol
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgQt.Picture
            
            Case 26 'carta K de trebol
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgKt.Picture
            
            Case 27 'carta A de pica
                puntosB(i) = 11
                frmBJ.imgBanca(i).Picture = frmBJ.imgAp.Picture
            
            Case 28 'carta 2 de pica
                puntosB(i) = 2
                frmBJ.imgBanca(i).Picture = frmBJ.img2p.Picture
            
            Case 29 'carta 3 de pica
                puntosB(i) = 3
                frmBJ.imgBanca(i).Picture = frmBJ.img3p.Picture
            
            Case 30 'carta 4 de pica
                puntosB(i) = 4
                frmBJ.imgBanca(i).Picture = frmBJ.img4p.Picture
            
            Case 31 'carta 5 de pica
                puntosB(i) = 5
                frmBJ.imgBanca(i).Picture = frmBJ.img5p.Picture
            
            Case 32 'carta 6 de pica
                puntosB(i) = 6
                frmBJ.imgBanca(i).Picture = frmBJ.img6p.Picture
            
            Case 33 'carta 7 de pica
                puntosB(i) = 7
                frmBJ.imgBanca(i).Picture = frmBJ.img7p.Picture
            
            Case 34 'carta 8 de pica
                puntosB(i) = 8
                frmBJ.imgBanca(i).Picture = frmBJ.img8p.Picture
            
            Case 35 'carta 9 de pica
                puntosB(i) = 9
                frmBJ.imgBanca(i).Picture = frmBJ.img9p.Picture
            
            Case 36 'carta 10 de pica
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.img10p.Picture
            
            Case 37 'carta J de pica
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgJp.Picture
            
            Case 38 'carta Q de pica
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgQp.Picture
            
            Case 39 'carta K de pica
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgKp.Picture
            
            Case 40 'carta A de diamante
                puntosB(i) = 11
                frmBJ.imgBanca(i).Picture = frmBJ.imgAd.Picture
            
            Case 41 'carta 2 de diamante
                puntosB(i) = 2
                frmBJ.imgBanca(i).Picture = frmBJ.img2d.Picture
            
            Case 42 'carta 3 de diamante
                puntosB(i) = 3
                frmBJ.imgBanca(i).Picture = frmBJ.img3d.Picture
            
            Case 43 'carta 4 de diamante
                puntosB(i) = 4
                frmBJ.imgBanca(i).Picture = frmBJ.img4d.Picture
            
            Case 44 'carta 5 de diamante
                puntosB(i) = 5
                frmBJ.imgBanca(i).Picture = frmBJ.img5d.Picture
            
            Case 45 'carta 6 de diamante
                puntosB(i) = 6
                frmBJ.imgBanca(i).Picture = frmBJ.img6d.Picture
            
            Case 46 'carta 7 de diamante
                puntosB(i) = 7
                frmBJ.imgBanca(i).Picture = frmBJ.img7d.Picture
            
            Case 47 'carta 8 de diamante
                puntosB(i) = 8
                frmBJ.imgBanca(i).Picture = frmBJ.img8d.Picture
            
            Case 48 'carta 9 de diamante
                puntosB(i) = 9
                frmBJ.imgBanca(i).Picture = frmBJ.img9d.Picture
                 
            Case 49 'carta 10 de diamante
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.img10d.Picture
            
            Case 50 'carta J de diamante
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgJd.Picture
            
            Case 51 'carta Q de diamante
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgQd.Picture
    
            Case 52 'carta K de diamante
                puntosB(i) = 10
                frmBJ.imgBanca(i).Picture = frmBJ.imgKd.Picture
        
        End Select
    Next i


End Sub

Function PedirNumeros(n As Integer) As Integer()
Dim i As Integer
ReDim Aux(n) As Integer     'arreglo aux para guardar los numeros aleatorios
    For i = LBound(Aux()) + 1 To UBound(Aux())  'Recorre todo el arreglo
        n = i
        Do
            n = n - 1      'compara desde atras si tiene numeros repetidos , y si los tiene genera uno nuevo
            If Aux(i) = Aux(n) Then
                Aux(i) = Int((Rnd * 52) + 1)    'pido numeros entre el 1 y el 52
                n = i
            End If
        Loop Until n = 0
    Next
PedirNumeros = Aux

End Function


Public Sub PartidaDeBanca()

puntosBtotal = puntosB(0) + puntosB(1)

If (puntosBtotal <= 16) Then                    'Si los puntos de la banca son menores o igual a 16 entonces pide otra carta
    puntosBtotal = puntosB(2) + puntosBtotal    'Se le suma los puntos de la tercera carta
    frmBJ.imgBanca(2).Visible = True            'se hace visible la tercera carta
    frmBJ.imgMazo(2).Visible = True             'se cubre la tercera carta con una imagen del mazo
    
    If ((puntosBtotal) <= 16) Then                  'Si los puntos de la banca siguen siendo menores o igual a 16 entonces pide otra carta
        puntosBtotal = puntosB(3) + puntosBtotal    'Se le suma los puntos de la cuarta carta
        frmBJ.imgBanca(3).Visible = True            'se hace visible la cuarta carta
        frmBJ.imgMazo(3).Visible = True             'se cubre la cuarta carta con una imagen del mazo
        
        If ((puntosBtotal) <= 16) Then                      'Si los puntos de la banca siguen siendo menores o igual a 16 entonces pide otra carta
            puntosBtotal = puntosB(4) + puntosBtotal        'Se le suma los puntos de la quinta carta
            frmBJ.imgBanca(4).Visible = True                'se hace visible la quinta carta
            frmBJ.imgMazo(4).Visible = True                 'se cubre la quinta carta con una imagen del mazo
            
            If ((puntosBtotal) <= 16) Then                      'Si los puntos de la banca siguen siendo menores o igual a 16 entonces pide otra carta
                puntosBtotal = puntosB(5) + puntosBtotal        'Se le suma los puntos de la sexta carta
                frmBJ.imgBanca(5).Visible = True                'se hace visible la sexta carta
                frmBJ.imgMazo(5).Visible = True                 'se cubre la septima carta con una imagen del mazo
                
                If ((puntosBtotal) <= 16) Then                      'Si los puntos de la banca siguen siendo menores o igual a 16 entonces pide otra carta
                    puntosBtotal = puntosB(6) + puntosBtotal        'Se le suma los puntos de la tercera carta
                    frmBJ.imgBanca(6).Visible = True                'se hace visible la septima carta
                    frmBJ.imgMazo(6).Visible = True                 'se cubre la septima carta con una imagen del mazo
                End If
                
            End If
            
        End If
        
    End If
    
End If

End Sub
