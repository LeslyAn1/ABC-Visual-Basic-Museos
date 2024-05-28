Imports System.Data.SqlClient
Imports System.Configuration
Public Class Form1
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadEstados()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not String.IsNullOrWhiteSpace(txtNombreEstado.Text) Then
            InsertEstado(txtNombreEstado.Text)
            LoadEstados()
        Else
            MessageBox.Show("Por favor, ingrese un nombre de estado.")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim id As Integer
        If Integer.TryParse(txtIdEstado.Text, id) AndAlso Not String.IsNullOrWhiteSpace(txtNombreEstado.Text) Then
            UpdateEstado(id, txtNombreEstado.Text)
            LoadEstados()
        Else
            MessageBox.Show("Por favor, ingrese un identificador de estado válido y un nombre de estado.")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim id As Integer
        If Integer.TryParse(txtIdEstado.Text, id) Then
            DeleteEstado(id)
            LoadEstados()
        Else
            MessageBox.Show("Por favor, seleccione un estado válido para eliminar.")
        End If
    End Sub

    Private Sub LoadEstados()
        Dim dataTable As DataTable = GetEstados()
        DataGridView1.DataSource = dataTable
        DataGridView1.Refresh()
    End Sub

    Public Sub InsertEstado(nombreEstado As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spInsertEstado", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@NombreEstado", nombreEstado)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function GetEstados() As DataTable
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Dim dataTable As New DataTable()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand("spGetEstados", connection)
                    command.CommandType = CommandType.StoredProcedure
                    Using reader As SqlDataReader = command.ExecuteReader()
                        dataTable.Load(reader)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al obtener estados: " & ex.Message)
        End Try
        Return dataTable
    End Function

    Public Sub UpdateEstado(idEstado As Integer, nombreEstado As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spUpdateEstado", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdEstado", idEstado)
                command.Parameters.AddWithValue("@NombreEstado", nombreEstado)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub DeleteEstado(idEstado As Integer)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spDeleteEstado", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdEstado", idEstado)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub dataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            txtIdEstado.Text = row.Cells("IdEstado").Value.ToString()
            txtNombreEstado.Text = row.Cells("NombreEstado").Value.ToString()
        End If
    End Sub
End Class
