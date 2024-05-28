Imports System.Data.SqlClient
Imports System.Configuration


Public Class Form3
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not String.IsNullOrWhiteSpace(txtNacionalidad.Text) Then
            InsertNacionalidad(txtNacionalidad.Text)
            LoadNacionalidades()
        Else
            MessageBox.Show("Por favor, ingrese una nacionalidad.")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim id As Integer
        If Integer.TryParse(txtIdNacionalidad.Text, id) AndAlso Not String.IsNullOrWhiteSpace(txtNacionalidad.Text) Then
            UpdateNacionalidad(id, txtNacionalidad.Text)
            LoadNacionalidades()
        Else
            MessageBox.Show("Por favor, ingrese un identificador de nacionalidad válido y una nacionalidad.")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim id As Integer
        If Integer.TryParse(txtIdNacionalidad.Text, id) Then
            DeleteNacionalidad(id)
            LoadNacionalidades()
        Else
            MessageBox.Show("Por favor, seleccione una nacionalidad válida para eliminar.")
        End If
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadNacionalidades()
    End Sub

    Private Sub LoadNacionalidades()
        Dim dataTable As DataTable = GetNacionalidades()
        DataGridView1.DataSource = dataTable
    End Sub

    Public Sub InsertNacionalidad(nacionalidad As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spInsertNacionalidad", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@TipoNacionalidad", nacionalidad)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function GetNacionalidades() As DataTable
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Dim dataTable As New DataTable()
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spGetNacionalidades", connection)
                command.CommandType = CommandType.StoredProcedure
                Using reader As SqlDataReader = command.ExecuteReader()
                    dataTable.Load(reader)
                End Using
            End Using
        End Using
        Return dataTable
    End Function

    Public Sub UpdateNacionalidad(idNacionalidad As Integer, nacionalidad As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spUpdateNacionalidad", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdNacionalidad", idNacionalidad)
                command.Parameters.AddWithValue("@TipoNacionalidad", nacionalidad)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub DeleteNacionalidad(idNacionalidad As Integer)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spDeleteNacionalidad", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdNacionalidad", idNacionalidad)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub dataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            txtIdNacionalidad.Text = row.Cells("IdNacionalidad").Value.ToString()
            txtNacionalidad.Text = row.Cells("TipoNacionalidad").Value.ToString()
        End If
    End Sub

End Class