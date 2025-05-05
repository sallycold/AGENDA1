Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Public Class Form1

    Dim CN As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OleDb.12.0; Data Source= Database51.accdb;")

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ActualizarLabel6()
    End Sub
    Public Sub ActualizarLabel6()
        If ComboBox1.Text = "Personal" Then
            TextBox9.Visible = True
            TextBox5.Visible = False
            Label6.Text = "Cumpleaños"
        Else
            TextBox9.Visible = False
            TextBox5.Visible = True
            Label6.Text = "Contacto"
        End If
    End Sub

    
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        LIMPIAR()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim fechaCumple As String = "NULL"
        If TextBox9.Visible AndAlso Not String.IsNullOrWhiteSpace(TextBox9.Text) Then
            Dim fechaTemp As Date
            If Date.TryParse(TextBox9.Text, fechaTemp) Then
                fechaCumple = "#" & fechaTemp.ToString("MM/dd/yyyy") & "#"
            Else
                MessageBox.Show("Ingrese una fecha válida en el campo Cumpleaños.")
                Exit Sub
            End If
        End If

        Dim CMD As New OleDb.OleDbCommand("UPDATE AGENDA SET NOMBRE = '" & TextBox1.Text & "', LOCALIDAD = '" & TextBox2.Text & "', DIRECCION = '" & TextBox3.Text & "', TELEFONO = '"
            & TextBox7.Text & "', TELEFONO2 = '" & TextBox8.Text & "', TELEFONO3 = '" & TextBox4.Text & "', CONTACTO = '" & TextBox5.Text & "', CUMPLEAÑOS = " & fechaCumple & ", [E-MAIL] = '" & TextBox6.Text &
            "' WHERE NOMBRE = '" & TextBox1.Text & "'", CN)

        CN.Open()
        CMD.ExecuteNonQuery()
        CN.Close()

        MsgBox("EL DATO SE ACTUALIZÓ")
        LIMPIAR()
        TextBox1.Enabled = False
        ComboBox1.Enabled = False




    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox5.Focus()
            e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        End
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox7.Visible = True
        TextBox8.Visible = True
    End Sub
    Sub LIMPIAR()
        ComboBox1.Text = ""
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox3.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox7.Visible = False
        TextBox8.Visible = False
        PictureBox3.Visible = False
        TextBox9.Visible = False
        ActualizarLabel6()
    End Sub

    Private Sub TextBox_KeyDown(sender As Object, e As KeyEventArgs) _
        Handles TextBox1.KeyDown, TextBox2.KeyDown, TextBox3.KeyDown, TextBox4.KeyDown, TextBox5.KeyDown, TextBox9.KeyDown, TextBox6.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(CType(sender, Control), True, True, True, True)
            TextBox1.Text = CapitalizarCadaPalabra(TextBox1.Text)
            TextBox2.Text = CapitalizarCadaPalabra(TextBox2.Text)
            TextBox3.Text = CapitalizarCadaPalabra(TextBox3.Text)
            TextBox5.Text = CapitalizarCadaPalabra(TextBox5.Text)
            e.SuppressKeyPress = True
        End If

        If ComboBox1.Text <> "" Then
            ComboBox1.Enabled = False
        End If

    End Sub
    Private Sub RedondearControl(control As Control, radio As Integer)
        Dim path As New Drawing2D.GraphicsPath()

        path.StartFigure()
        path.AddArc(New Rectangle(0, 0, radio, radio), 180, 90)
        path.AddArc(New Rectangle(control.Width - radio, 0, radio, radio), 270, 90)
        path.AddArc(New Rectangle(control.Width - radio, control.Height - radio, radio, radio), 0, 90)
        path.AddArc(New Rectangle(0, control.Height - radio, radio, radio), 90, 90)
        path.CloseFigure()

        control.Region = New Region(path)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim tel2 As Integer
        Dim tel3 As Integer
        Dim fechaCumple As String = "NULL"


        If Not Integer.TryParse(TextBox4.Text, tel2) Then
            MessageBox.Show("TELEFONO3 debe ser un número.")
            Exit Sub
        End If

        If Not Integer.TryParse(TextBox7.Text, tel3) Then
            MessageBox.Show("TELEFONO3 debe ser un número.")
            Exit Sub
        End If
        If Not Integer.TryParse(TextBox8.Text, tel3) Then
            MessageBox.Show("TELEFONO3 debe ser un número.")
            Exit Sub
        End If


        If TextBox9.Visible AndAlso Not String.IsNullOrWhiteSpace(TextBox9.Text) Then
            Dim fechaTemp As Date
            If Date.TryParse(TextBox9.Text, fechaTemp) Then
                fechaCumple = "#" & fechaTemp.ToString("MM/dd/yyyy") & "#"
            Else
                MessageBox.Show("Ingrese una fecha válida en el campo Cumpleaños.")
                Exit Sub
            End If
        End If

        If Not TextBox6.Text.Contains("@gmail.com") Then
            MessageBox.Show("Falta el '@gmail.com' en el correo electrónico.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If


        Dim query As String = "INSERT INTO AGENDA (AGRUPADO, NOMBRE, DIRECCION, LOCALIDAD, TELEFONO, TELEFONO2, TELEFONO3, CONTACTO, CUMPLEAÑOS, [E-MAIL]) " &
                          "VALUES ('" & ComboBox1.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "'," &
                          tel2 & "," & tel3 & ",'" & TextBox5.Text & "'," & fechaCumple & ",'" & TextBox6.Text & "')"


        Try
            CN.Open()
            Dim CMD As New OleDb.OleDbCommand(query, CN)
            CMD.ExecuteNonQuery()
            CN.Close()

            MessageBox.Show("El dato se agregó correctamente.", "Éxito")
            LIMPIAR()
        Catch ex As Exception
            MessageBox.Show("Error al insertar: " & ex.Message, "Error")
            CN.Close()
        End Try

        If ComboBox1.Text = "" Then
            ComboBox1.Enabled = True
        End If
        TextBox1.Enabled = True

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Enabled = True
        ComboBox1.Enabled = True
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then

            If TextBox1.Text.Trim() <> "" Then
                TextBox1.Enabled = False
                Me.SelectNextControl(TextBox1, True, True, True, True)
                e.SuppressKeyPress = True
            Else
                MessageBox.Show("Por favor, ingrese un nombre antes de continuar.")
            End If
        End If

    End Sub

    'poner las primeras palabras en mayuscula
    Private Function CapitalizarCadaPalabra(texto As String) As String
        Return StrConv(texto.Trim().ToLower(), VbStrConv.ProperCase)
    End Function


End Class
