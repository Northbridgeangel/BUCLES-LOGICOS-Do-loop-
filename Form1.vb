Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim RESPUESTA As Integer
        Dim C As Integer
        'HAZLO Y REPITE HASTA A C= C+1
        Do
            C += 1 'SERIA IGUAL A C= C+1
            RESPUESTA = MsgBox(" LO HACEMOS OTRA VEZ Y VAN YA " & C, vbYesNo)
        Loop Until RESPUESTA <> vbYes
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim RESPUESTA As Integer
        Dim C As Integer
        'HAZLO MIENTRAS
        Do While RESPUESTA <> vbNo
            C += 1 'SERIA IGUAL A C= C+1
            RESPUESTA = MsgBox(" LO HACEMOS OTRA VEZ Y VAN YA " & C, vbYesNo)
        Loop
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim RESPUESTA As Integer
        Dim C As Integer
        'MIENTRAS QUE HAZLO
        While RESPUESTA <> vbNo
            C += 1 'SERIA IGUAL A C= C+1
            RESPUESTA = MsgBox(" LO HACEMOS OTRA VEZ Y VAN YA " & C, vbYesNo)
        End While
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim NUMERO As Integer
        Dim RESTO = 1
        'HAZLO Y REPITE HASTA C= C+1
        While RESTO <> 0
            NUMERO = InputBox("METE UN NUMERO")
            RESTO = NUMERO Mod 2
        End While
      
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim VECTOR(10) As String
        For I = 0 To 10
            VECTOR(I) = InputBox("METE JUGADOR" & I)
        Next
        For I = 0 To 10
            ComboBox1.Items.Add(VECTOR(I))
        Next
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim N As Integer 'CUANTOS VOY A ALMACENAR EN EL VECTOR
        Dim NS As Integer  'Nº SUPERIOR AL AZAR
        Dim NI As Integer   'Nº INFERIOR AL AZAR
        N = InputBox("CUANTOS NUMEROS AL AZAR ALMACENAMOS", "CUANTOS NUMEROS AL AZAR ALMACENAMOS")


        Dim VECTOR(N) As Integer
        While NI >= NS 'MIENTRAS QUE INFERIOR SEA 0 O MAYOR QUE SUPERIOR
            NS = InputBox("NUMERO AL AZAR", "AZAR SUPERIOR")
        NI = InputBox("NUMERO AL AZAR INFERIOR", "AZAR INFERIOR")
            If NI >= NS Then MsgBox("ERROR, HAS METIDO EN NUMERO INFERIOR UN Nº MAYOR O IGUAL")
        End While

        For I = 1 To N
            VECTOR(I) = Int(Rnd() * (NS - NI) + NI)
        Next

        For I = 1 To N
            ComboBox2.Items.Add(VECTOR(I))
        Next

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim LISTA(5, 3)
        For I = 1 To 5
            LISTA(I, 1) = InputBox("Mete Nombre")
            LISTA(I, 2) = InputBox("Mete DNI")
            LISTA(I, 3) = InputBox("Mete TELÉFONO")
        Next

        For I = 1 To 5
            ComboBox3.Items.Add(LISTA(I, 1))
            ComboBox4.Items.Add(LISTA(I, 2))
            ComboBox5.Items.Add(LISTA(I, 3))
        Next

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        'AL SELECCIONAR UN ELEMENTO DEL ÍNDICE DE COMBOBOX5, TE MUESTRA SUS IGUALES EN COMBOS 3 Y 4
        ComboBox4.SelectedIndex = ComboBox5.SelectedIndex
        ComboBox3.SelectedIndex = ComboBox5.SelectedIndex
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        'AL SELECCIONAR UN ELEMENTO DEL ÍNDICE DE COMBOBOX4, TE MUESTRA SUS IGUALES EN COMBOS 5 Y 3
        ComboBox5.SelectedIndex = ComboBox4.SelectedIndex
        ComboBox3.SelectedIndex = ComboBox4.SelectedIndex
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        'AL SELECCIONAR UN ELEMENTO DEL ÍNDICE DE COMBOBOX3, TE MUESTRA SUS IGUALES EN COMBOS 4 Y 5
        ComboBox4.SelectedIndex = ComboBox3.SelectedIndex
        ComboBox5.SelectedIndex = ComboBox3.SelectedIndex
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Randomize()
        Dim VECTOR(500)
        Dim NMAXIMO, NMINIMO, POSIMAXIMO, POSIMINIMO As Integer
        NMAXIMO = 0
        NMINIMO = 1000
        'RELLENO EL VECTOR
        For I = 1 To 500
            VECTOR(I) = Int((Rnd() * 500) + 200)
            ComboBox6.Items.Add(VECTOR(I))
        Next

        'BUSCAR EL MAYOR Y EL MENOR Y SUS POSICIONES (POSCICIÓN I)
        For I = 1 To 500
            If VECTOR(I) > NMAXIMO Then
                NMAXIMO = VECTOR(I)
                POSIMAXIMO = I
            End If
            If VECTOR(I) < NMINIMO Then
                NMINIMO = VECTOR(I)
                POSIMINIMO = I
            End If
        Next
        'MUESTRA LOS RESULTADOS
        MsgBox(" EL MAS GRANDE ES " & NMAXIMO)
        MsgBox(" EL MAS PEQUEÑO ES " & NMINIMO)
        MsgBox(" LA POSICION DEL GRANDE ES " & POSIMAXIMO)
        MsgBox(" LA POSICION DEL GRANDE ES " & POSIMINIMO)

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

    End Sub
End Class
