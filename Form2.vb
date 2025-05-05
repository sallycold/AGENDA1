Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class Form2

    Dim CN As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OleDb.12.0; Data Source= Database51.accdb;")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Or ComboBox3.Text = "" Or TextBox1.Text.Trim = "" Or ComboBox2.Text = "" Then
            MessageBox.Show("Completa todos los campos de búsqueda.", "Faltan datos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim campo As String = ComboBox1.Text
        Dim tipoBusqueda As String = ComboBox3.Text
        Dim palabraClave As String = TextBox1.Text.Trim()
        Dim operadorLike As String = ""

        ' Determinar cómo armar el LIKE
        Select Case tipoBusqueda
            Case "que comience"
                operadorLike = palabraClave & "%"
            Case "que termine"
                operadorLike = "%" & palabraClave
            Case "que contenga"
                operadorLike = "%" & palabraClave & "%"
        End Select

        ' Campo para ordenar
        Dim campoOrden As String = ComboBox2.Text

        ' Ordenamiento según CheckBox
        Dim ordenSql As String = ""
        If CheckBox1.Checked AndAlso Not CheckBox2.Checked Then
            ordenSql = "ASC"
        ElseIf CheckBox2.Checked AndAlso Not CheckBox1.Checked Then
            ordenSql = "DESC"
        Else
            MessageBox.Show("Seleccioná solo una opción de orden: Ascendente o Descendente.", "Orden inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Armar consulta
        Dim consulta As String = $"SELECT * FROM AGENDA WHERE [{campo}] LIKE ? ORDER BY [{campoOrden}] {ordenSql}"
            'Dim consulta As String = "SELECT * FROM AGENDA WHERE campo LIKE ? ORDER BY " & campoOrden & "" & ordenSql
            

        Dim comando As New OleDbCommand(consulta, CN)
        comando.Parameters.AddWithValue("?", operadorLike)

        Try
            CN.Open()
            Dim adaptador As New OleDbDataAdapter(comando)
            Dim tabla As New DataTable()
            adaptador.Fill(tabla)
            DataGridView1.DataSource = tabla
        Catch ex As Exception
            MessageBox.Show("Error al buscar: " & ex.Message)
        Finally
            CN.Close()
        End Try


    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim campos() As String = {"NOMBRE", "DIRECCION", "LOCALIDAD", "CUMPLEAÑOS", "CONTACTO", "TELEFONO", "TELEFONO2", "TELEFONO3"}

        ComboBox1.Items.AddRange(campos) ' Campo para filtrar (WHERE)
        ComboBox3.Items.AddRange({"que comience", "que termine", "que contenga"}) ' Tipo de búsqueda
        ComboBox2.Items.AddRange(campos) ' Campo para ordenar (ORDER BY)

        'ComboBox1.SelectedIndex = 0
       ' ComboBox3.SelectedIndex = 2
        'ComboBox2.SelectedIndex = 0

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim NOMBRE, DIRECCION, LOCALIDAD, CONTACTO, AGRUPADO, E_MAIL, CUMPLEAÑOS As String
        Dim TELEFONO, TELEFONO2, TELEFONO3 As Integer


        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)


            AGRUPADO = selectedRow.Cells(0).Value.ToString()
            NOMBRE = selectedRow.Cells(1).Value.ToString()
            DIRECCION = selectedRow.Cells(2).Value.ToString()
            LOCALIDAD = selectedRow.Cells(3).Value.ToString()
            TELEFONO = selectedRow.Cells(4).Value.ToString()
            TELEFONO2 = selectedRow.Cells(5).Value.ToString()
            TELEFONO3 = selectedRow.Cells(6).Value.ToString()
            CONTACTO = selectedRow.Cells(7).Value.ToString()
            CUMPLEAÑOS = selectedRow.Cells(8).Value.ToString()
            E_MAIL = selectedRow.Cells(9).Value.ToString()



            Form1.ComboBox1.Text = AGRUPADO
            Form1.TextBox1.Text = NOMBRE
            Form1.TextBox2.Text = DIRECCION
            Form1.TextBox3.Text = LOCALIDAD
            Form1.TextBox4.Text = TELEFONO
            Form1.TextBox7.Text = TELEFONO2
            Form1.TextBox8.Text = TELEFONO3
            Form1.TextBox5.Text = CONTACTO
            Form1.TextBox9.Text = CUMPLEAÑOS
            Form1.TextBox6.Text = E_MAIL



        Else
            MessageBox.Show("Seleccioná una fila completa en el DataGridView.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Form1.Show()
            Form1.RestablecerComboBox1()
            Me.Hide()

        Catch ex As Exception
            MessageBox.Show("Error al mostrar Form1: " & ex.Message)
        End Try
    End Sub


End Class

        
