Imports System.Data.OleDb
Imports System.Data
Public Class jacket
    Private dv As New DataView
    Dim comando As New OleDbCommand
    Private Sub Jacket_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'TrazaDataSet.Cabler' table. You can move, or remove it, as needed.
        Dim cnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb")
        Dim da As New OleDbDataAdapter("SELECT * FROM Cabler", cnn)
        Dim ds As New DataSet
        da.Fill(ds)
        'DataGridView1.DataSource = ds.Tables(0)
        dv.Table = ds.Tables(0)
        DataGridView1.DataSource = dv
        'DataGridView2.DataSource = dv
        cnn.Close()

        Label4.Text = DateTime.Now.ToString("dd/MM/yyyy") & " " & DateTime.Now.ToLongTimeString

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        'TODO: This line of code loads data into the 'TrazaDataSet.Cabler' table. You can move, or remove it, as needed.
        'Dim cnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb")
        'Dim da As New OleDbDataAdapter("SELECT * FROM Cabler", cnn)
        'Dim ds As New DataSet
        'da.Fill(ds)
        ''DataGridView1.DataSource = ds.Tables(0)
        'dv.Table = ds.Tables(0)
        'DataGridView1.DataSource = dv
        'cnn.Close()
        dv.RowFilter = String.Format("Traza_cabler like '%{0}%'", TextBox4.Text)
        'dv.RowFilter = String.Format("Cant_Cabler like '%{0}%'", TextBox5.Text)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox6.Text = "" Then
            MsgBox("LOS DATOS DE TRAZABILIDAD, NP O CANTIDAD ESTAN VACIOS. NO SE HA GUARDADO NINGUN REGISTRO.", MsgBoxStyle.Information, "AVISO")
            Exit Sub
        End If

        If TextBox3.Text = "" Or TextBox5.Text = "" Then
            MsgBox("LOS DATOS DE TRAZABILIDAD DE CABLER ESTAN VACIOS. ESTOS SON NECESARIOS PARA ACTUALIZAR LA BASE DE DATOS. NO FUE ACTUALIZADO NINGUN REGISTRO.", MsgBoxStyle.Information, "AVISO")
            Exit Sub
        End If
        Try
            Dim cnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb")
            cnn.Open()
            Dim actualizar As String
            'actualizar = "UPDATE Cabler SET NP_Jacket='" & TextBox1.Text & "', Traza_Jacket='" & TextBox2.Text & "', Cant_Jacket='" & TextBox6.Text & "', Fecha_Jacket='" & Label4.Text & "' WHERE Traza_cabler='" & TextBox3.Text & "'"
            actualizar = "UPDATE Cabler SET NP_Jacket='" & TextBox1.Text & "', Traza_Jacket='" & TextBox2.Text & "', Cant_Jacket='" & TextBox6.Text & "', Fecha_Jacket='" & Label4.Text & "' WHERE Traza_cabler='" & TextBox3.Text & "' AND Cant_Cabler='" & TextBox5.Text & "'"
            comando = New OleDbCommand(actualizar, cnn)
            comando.ExecuteNonQuery()
            MsgBox("REGISTRO DE JACKET CARGADO CON EXITO EN LA BASE DE DATOS.", MsgBoxStyle.Information, "AVISO")
            'cnn.Close()


            Dim da As New OleDbDataAdapter("SELECT * FROM Cabler", cnn)
            Dim ds As New DataSet
            da.Fill(ds)
            'DataGridView1.DataSource = ds.Tables(0)
            dv.Table = ds.Tables(0)
            DataGridView1.DataSource = dv
            cnn.Close()
            Label4.Text = DateTime.Now.ToString("dd/MM/yyyy") & " " & DateTime.Now.ToLongTimeString

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""


        Catch ex As Exception
            MsgBox("SUCEDIO UN ERROR AL CARGAR LOS DATOS.")
        End Try

    End Sub

    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs)
        'TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value()
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value()
        TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value() ' & "FT"
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb")
        Dim da As New OleDbDataAdapter("SELECT * FROM Cabler", cnn)
        Dim ds As New DataSet
        da.Fill(ds)
        'DataGridView1.DataSource = ds.Tables(0)
        dv.Table = ds.Tables(0)
        DataGridView1.DataSource = dv
        'DataGridView2.DataSource = dv
        cnn.Close()

        Label4.Text = DateTime.Now.ToString("dd/MM/yyyy") & " " & DateTime.Now.ToLongTimeString
    End Sub
End Class