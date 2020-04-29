Imports System.Data.OleDb
Imports System.Data
Public Class buscar
    Dim dv As New DataView
    Private Sub buscar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'TrazaDataSet.Cabler' table. You can move, or remove it, as needed.
        Dim cnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\TempFlex\Daniel H\Trazabilidad.mdb")
        Dim da As New OleDbDataAdapter("SELECT * FROM Cabler", cnn)
        Dim ds As New DataSet
        da.Fill(ds)
        'DataGridView1.DataSource = ds.Tables(0)
        dv.Table = ds.Tables(0)
        DataGridView1.DataSource = dv


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        dv.RowFilter = String.Format("Traza_Jacket like '%{0}%'", TextBox1.Text)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class