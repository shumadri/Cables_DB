Imports System.Data.OleDb
Public Class braider

    Dim comando As New OleDbCommand
    Private dv As New DataView
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim cnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb")
        cnn.Open()
        Dim actualizar As String
        actualizar = "UPDATE Cabler SET NP_Braider='" & TextBox1.Text & "', Traza_Braider='" & TextBox2.Text & "', Fecha_Braider='" & Label4.Text & "' WHERE Traza_cabler='" & TextBox3.Text & "'"
        comando = New OleDbCommand(actualizar, cnn)
        comando.ExecuteNonQuery()
        MsgBox("Registro cargado con exito")
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
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        'dv.RowFilter = String.Format("Traza_cabler like '%{0}%'", TextBox3.Text)
    End Sub

    Private Sub braider_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'TrazaDataSet.Cabler' table. You can move, or remove it, as needed.
        Dim cnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb")
        Dim da As New OleDbDataAdapter("SELECT * FROM Cabler", cnn)
        Dim ds As New DataSet
        da.Fill(ds)
        'DataGridView1.DataSource = ds.Tables(0)
        dv.Table = ds.Tables(0)
        DataGridView1.DataSource = dv
        cnn.Close()

        Label4.Text = DateTime.Now.ToString("dd/MM/yyyy") & " " & DateTime.Now.ToLongTimeString
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        dv.RowFilter = String.Format("Traza_cabler like '%{0}%'", TextBox4.Text)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class