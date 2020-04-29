Imports System.Data.OleDb
Imports System.Data
Public Class cabler
    Dim Con_access As New OleDbConnection
    Dim comando As New OleDbCommand
    Dim adaptador As New OleDbDataAdapter
    Dim datos As New DataTable
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        button1.Enabled = False
        textBox1.Visible = True
        textBox2.Visible = True
        textBox3.Visible = True
        textBox4.Visible = False
        textBox5.Visible = False
        textBox6.Visible = False
        textBox7.Visible = False
        textBox8.Visible = False
        textBox9.Visible = False
        label3.Visible = False
        label4.Visible = False
        label5.Visible = False
        label6.Visible = False
        label7.Visible = False
        label8.Visible = False
        comboBox1.Text = ""
        Button3.BackColor = Color.Red

        label12.Text = DateTime.Now.ToString("dd/MM/yyyy") & " " & DateTime.Now.ToLongTimeString


        Try
            Con_access.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb"
            Con_access.Open()
            comando.Connection = Con_access
            comando.CommandType = CommandType.Text
            comando.CommandText = "SELECT * FROM Part_Number"
            adaptador.SelectCommand = comando
            adaptador.Fill(datos)
            Me.textBox10.DataSource = datos
            Me.textBox10.DisplayMember = "PN"
            Me.textBox10.ValueMember = "ID"
            Me.textBox10.ResetText()
            Me.textBox10.SelectedText = "Seleccione NP"
            Con_access.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            'MessageBox.Show("Problema al conectar a BD")
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboBox1.SelectedIndexChanged
        If comboBox1.Text = "2" Then


            textBox2.Visible = True
            textBox3.Visible = True
            textBox4.Visible = False
            textBox5.Visible = False
            textBox6.Visible = False
            textBox7.Visible = False
            textBox8.Visible = False
            textBox9.Visible = False
            TextBox21.Visible = True
            TextBox29.Visible = True
            TextBox37.Visible = True
            TextBox20.Visible = True
            TextBox28.Visible = True
            TextBox36.Visible = True

            textBox4.Visible = False

            TextBox19.Visible = False
            TextBox27.Visible = False

            TextBox35.Visible = False
            TextBox18.Visible = False
            TextBox26.Visible = False

            TextBox34.Visible = False

            TextBox17.Visible = False
            TextBox25.Visible = False

            TextBox33.Visible = False
            TextBox16.Visible = False
            TextBox24.Visible = False

            TextBox33.Visible = False
            TextBox14.Visible = False
            TextBox23.Visible = False

            TextBox31.Visible = False
            TextBox13.Visible = False
            TextBox22.Visible = False
            TextBox30.Visible = False

            label3.Visible = False
            label4.Visible = False
            label5.Visible = False
            label6.Visible = False
            label7.Visible = False
            label8.Visible = False
            TextBox32.Visible = False

            label9.Visible = True
        End If


        If comboBox1.Text = "4" Then


            textBox2.Visible = True
            textBox3.Visible = True
            textBox4.Visible = True
            textBox5.Visible = True
            textBox6.Visible = False
            textBox7.Visible = False
            textBox8.Visible = False
            textBox9.Visible = False
            label3.Visible = True
            label4.Visible = True
            label5.Visible = False
            label6.Visible = False
            label7.Visible = False
            label8.Visible = False

            TextBox19.Visible = True
            TextBox27.Visible = True

            TextBox35.Visible = True
            TextBox18.Visible = True
            TextBox26.Visible = True

            TextBox34.Visible = True
            TextBox32.Visible = False
            TextBox17.Visible = False
            TextBox25.Visible = False

            TextBox33.Visible = False
            TextBox16.Visible = False
            TextBox24.Visible = False

            TextBox33.Visible = False
            TextBox14.Visible = False
            TextBox23.Visible = False

            TextBox31.Visible = False
            TextBox13.Visible = False
            TextBox22.Visible = False
            TextBox30.Visible = False



        End If



        If comboBox1.Text = "8" Then


            textBox2.Visible = True
            textBox3.Visible = True
            textBox4.Visible = True
            textBox5.Visible = True
            textBox6.Visible = True
            textBox7.Visible = True
            textBox8.Visible = True
            textBox9.Visible = True
            label3.Visible = True
            label4.Visible = True
            label5.Visible = True
            label6.Visible = True
            label7.Visible = True
            label8.Visible = True
            TextBox19.Visible = True
            TextBox27.Visible = True

            TextBox35.Visible = True
            TextBox18.Visible = True
            TextBox26.Visible = True
            TextBox34.Visible = True

            TextBox17.Visible = True
            TextBox25.Visible = True

            TextBox33.Visible = True
            TextBox16.Visible = True
            TextBox24.Visible = True

            TextBox33.Visible = True
            TextBox14.Visible = True
            TextBox23.Visible = True

            TextBox31.Visible = True
            TextBox13.Visible = True
            TextBox22.Visible = True
            TextBox30.Visible = True
            TextBox32.Visible = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles button1.Click

        If textBox10.Text = "Seleccione NP" Then
            MsgBox("Seleccione un numero de parte a cargar", MsgBoxStyle.Information, "AVISO")
            Exit Sub
        End If

        If textBox1.Text = "" Or TextBox11.Text = "" Then
            MsgBox("Los datos de Trazabilidad o cantidad de carretes estan vacios.", MsgBoxStyle.Information, "AVISO")
            Exit Sub
        End If

        If comboBox1.Text = "2" Then
            If textBox2.Text = "" Or textBox3.Text = "" Or TextBox21.Text = "" Or TextBox29.Text = "" Or TextBox37.Text = "" Or TextBox20.Text = "" Or TextBox28.Text = "" Or TextBox36.Text = "" Then
                MsgBox("Los datos de Twinax estan vacios. La informacion no se ha cargado.", MsgBoxStyle.Information, "AVISO")
                Exit Sub
            End If
        End If

        If comboBox1.Text = "4" Then
            If textBox2.Text = "" Or textBox3.Text = "" Or textBox4.Text = "" Or textBox5.Text = "" Or TextBox21.Text = "" Or TextBox29.Text = "" Or TextBox37.Text = "" Or TextBox20.Text = "" Or TextBox28.Text = "" Or TextBox36.Text = "" Or TextBox19.Text = "" Or TextBox27.Text = "" Or TextBox35.Text = "" Or TextBox18.Text = "" Or TextBox26.Text = "" Or TextBox34.Text = "" Then
                MsgBox("Los datos de Twinax estan vacios. La informacion no se ha cargado.", MsgBoxStyle.Information, "AVISO")
                Exit Sub
            End If
        End If

        If comboBox1.Text = "8" Then
            If textBox2.Text = "" Or textBox3.Text = "" Or textBox4.Text = "" Or textBox5.Text = "" Or textBox6.Text = "" Or textBox7.Text = "" Or textBox8.Text = "" Or textBox9.Text = "" Or TextBox21.Text = "" Or TextBox29.Text = "" Or TextBox37.Text = "" Or TextBox20.Text = "" Or TextBox28.Text = "" Or TextBox36.Text = "" Or TextBox19.Text = "" Or TextBox27.Text = "" Or TextBox35.Text = "" Or TextBox18.Text = "" Or TextBox26.Text = "" Or TextBox34.Text = "" Or TextBox17.Text = "" Or TextBox25.Text = "" Or TextBox33.Text = "" Or TextBox16.Text = "" Or TextBox24.Text = "" Or TextBox32.Text = "" Or TextBox14.Text = "" Or TextBox23.Text = "" Or TextBox31.Text = "" Or TextBox13.Text = "" Or TextBox23.Text = "" Or TextBox30.Text = "" Then
                MsgBox("Los datos de Twinax estan vacios. La informacion no se ha cargado.", MsgBoxStyle.Information, "AVISO")
                Exit Sub
            End If
        End If



        Dim conect_string As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb"
        Dim conect As New OleDb.OleDbConnection(conect_string)
        conect.Open()

        Dim query As String = "INSERT INTO Cabler (N_pares, NP_cabler,Traza_cabler,Twinax1,Twinax2,Twinax3,Twinax4,Twinax5,Twinax6,Twinax7,Twinax8,Cant_Cabler,Fecha_Cabler) VALUES (@pares, @npclaber,@trazacabler,@tw1,@tw2,@tw3,@tw4,@tw5,@tw6,@tw7,@tw8,@cant,@fecha)"
        Dim Comando As New OleDb.OleDbCommand(query, conect)
        Comando.Parameters.AddWithValue("@pares", comboBox1.Text)
        Comando.Parameters.AddWithValue("@npclaber", textBox10.Text)
        Comando.Parameters.AddWithValue("@trazacabler", textBox1.Text)
        Comando.Parameters.AddWithValue("@tw1", textBox2.Text & "-" & TextBox21.Text & "-" & TextBox29.Text & "-" & TextBox37.Text)
        Comando.Parameters.AddWithValue("@tw2", textBox3.Text & "-" & TextBox20.Text & "-" & TextBox28.Text & "-" & TextBox36.Text)
        Comando.Parameters.AddWithValue("@tw3", textBox4.Text & "-" & TextBox19.Text & "-" & TextBox27.Text & "-" & TextBox35.Text)
        Comando.Parameters.AddWithValue("@tw4", textBox5.Text & "-" & TextBox18.Text & "-" & TextBox26.Text & "-" & TextBox34.Text)
        Comando.Parameters.AddWithValue("@tw5", textBox6.Text & "-" & TextBox17.Text & "-" & TextBox25.Text & "-" & TextBox33.Text)
        Comando.Parameters.AddWithValue("@tw6", textBox7.Text & "-" & TextBox16.Text & "-" & TextBox24.Text & "-" & TextBox32.Text)
        Comando.Parameters.AddWithValue("@tw7", textBox8.Text & "-" & TextBox14.Text & "-" & TextBox23.Text & "-" & TextBox31.Text)
        Comando.Parameters.AddWithValue("@tw8", textBox9.Text & "-" & TextBox13.Text & "-" & TextBox22.Text & "-" & TextBox30.Text)
        Comando.Parameters.AddWithValue("@cant", TextBox11.Text)
        Comando.Parameters.AddWithValue("@fecha", DateTime.Now.ToString("dd/MM/yyyy") & " " & DateTime.Now.ToLongTimeString)
        Comando.ExecuteNonQuery()
        conect.Close()

        Dim response
        response = MsgBox("LOS DATOS FUERON CARGADOS CON EXITO EN LA BASE DE DATOS. ¿DESEA CARGAR OTRO CARRETE CON LOS MISMOS DATOS DE TWINAX?", MsgBoxStyle.YesNo, "NOTIFICACION")

        If response = vbYes Then
            TextBox11.Text = ""
            textBox1.Text = ""
        Else
            textBox1.Text = ""
            textBox2.Text = ""
            textBox3.Text = ""
            textBox4.Text = ""
            textBox5.Text = ""
            textBox6.Text = ""
            textBox7.Text = ""
            textBox8.Text = ""
            textBox9.Text = ""
            textBox10.Text = ""
            TextBox11.Text = ""
            TextBox19.Text = ""
            TextBox27.Text = ""

            TextBox21.Text = ""
            TextBox29.Text = ""
            TextBox37.Text = ""
            TextBox20.Text = ""
            TextBox28.Text = ""
            TextBox36.Text = ""

            TextBox35.Text = ""
            TextBox18.Text = ""
            TextBox26.Text = ""
            TextBox34.Text = ""

            TextBox17.Text = ""
            TextBox25.Text = ""

            TextBox33.Text = ""
            TextBox16.Text = ""
            TextBox24.Text = ""

            TextBox33.Text = ""
            TextBox14.Text = ""
            TextBox23.Text = ""

            TextBox31.Text = ""
            TextBox13.Text = ""
            TextBox22.Text = ""
            TextBox30.Text = ""
            TextBox32.Text = ""
            button1.Enabled = False
            Button3.Enabled = True
            Button3.BackColor = Color.Red
            TextBox12.Text = ""
            TextBox15.Text = ""
        End If


    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles textBox10.TextChanged
        If textBox10.Text = "Seleccione NP" Then
            TextBox12.Text = ""
            TextBox15.Text = ""
            comboBox1.Text = ""
        End If

        Try

            Dim cadenaconexion As String
            cadenaconexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb"
            Dim miconexion As OleDbConnection
            miconexion = New OleDbConnection(cadenaconexion)

            Dim PNtableadapter As OleDbDataAdapter
            PNtableadapter = New OleDbDataAdapter

            PNtableadapter.SelectCommand = New OleDbCommand("SELECT * FROM Part_Number WHERE PN = '" & textBox10.Text & "'", miconexion)
            Dim PNdataset As DataSet
            PNdataset = New DataSet

            PNdataset.Tables.Add("Part_Number")
            PNtableadapter.Fill(PNdataset.Tables("Part_Number"))

            TextBox12.DataBindings.Clear()
            TextBox12.DataBindings.Add("Text", PNdataset.Tables("Part_Number"), "Twinax")

            TextBox15.DataBindings.Clear()
            TextBox15.DataBindings.Add("Text", PNdataset.Tables("Part_Number"), "Filler")

            comboBox1.DataBindings.Clear()
            comboBox1.DataBindings.Add("Text", PNdataset.Tables("Part_Number"), "Cant_Pares")



            miconexion.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            ' MessageBox.Show("Sucedio un error al cargar los datos, por favor vuelva a intentar")
        End Try
    End Sub

    Private Sub textBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub


    Private Sub TextBox29_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox29.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox37_TextChanged(sender As Object, e As EventArgs) Handles TextBox37.TextChanged

    End Sub

    Private Sub TextBox37_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox37.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub textBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox28_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox28.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox36_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox36.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub textBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox27_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox27.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox35_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox35.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub textBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox26.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox34_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox34.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub textBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox33_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox33.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub textBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox32_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox32.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub textBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox31_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox31.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub textBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox30_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox30.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub textBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles textBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If comboBox1.Text = "Seleccione NP" Or TextBox12.Text = "" Or TextBox15.Text = "" Then
            MsgBox("NO SE HA SELECCIONADO UN NUMERO DE PARTE O LOS DATOS DEL BOM NO SE HAN CARGADO.")
            Exit Sub

        End If


        Dim cadenaOriginal As String = TextBox12.Text
        'Dim tam_var = cadenaOriginal.Length
        Dim final As String = Strings.Right(cadenaOriginal, 3)
        Dim inicio As String = Strings.Left(cadenaOriginal, 6)
        TextBox38.Text = final
        TextBox39.Text = inicio

        If comboBox1.Text = 2 Then
            If LCase(TextBox21.Text) = "1p" & inicio & "1" & final And LCase(TextBox20.Text) = "1p" & inicio & "2" & final Then
                button1.Enabled = True
                MsgBox("DATOS VALIDADOS.", MsgBoxStyle.Information, "AVISO")
                Button3.BackColor = Color.Green
                Button3.Enabled = False
            Else
                MsgBox("LOS DATOS DE TWINAX NO COINCIDEN", MsgBoxStyle.Information, "AVISO")
                button1.Enabled = False
            End If
        End If


        If comboBox1.Text = 4 Then
            If LCase(TextBox21.Text) = "1p" & inicio & "1" & final And LCase(TextBox20.Text) = "1p" & inicio & "2" & final And LCase(TextBox19.Text) = "1p" & inicio & "3" & final And LCase(TextBox18.Text) = "1p" & inicio & "4" & final Then
                button1.Enabled = True
                MsgBox("DATOS VALIDADOS", MsgBoxStyle.Information, "AVISO")
                Button3.BackColor = Color.Green
                Button3.Enabled = False
            Else
                MsgBox("LOS DATOS DE TWINAX NO COINCIDEN", MsgBoxStyle.Information, "AVISO")
                button1.Enabled = False
            End If
        End If

        If comboBox1.Text = 8 Then
            If LCase(TextBox21.Text) = "1p" & inicio & "1" & final And LCase(TextBox20.Text) = "1p" & inicio & "2" & final And LCase(TextBox19.Text) = "1p" & inicio & "3" & final And LCase(TextBox18.Text) = "1p" & inicio & "4" & final And LCase(TextBox17.Text) = "1p" & inicio & "5" & final And LCase(TextBox16.Text) = "1p" & inicio & "6" & final And LCase(TextBox14.Text) = "1p" & inicio & "7" & final And LCase(TextBox13.Text) = "1p" & inicio & "8" & final Then
                button1.Enabled = True
                MsgBox("DATOS VALIDADOS", MsgBoxStyle.Information, "AVISO")
                Button3.BackColor = Color.Green
                Button3.Enabled = False
            Else
                MsgBox("LOS DATOS DE TWINAX NO COINCIDEN", MsgBoxStyle.Information, "AVISO")
                button1.Enabled = False
            End If
        End If



    End Sub
End Class
