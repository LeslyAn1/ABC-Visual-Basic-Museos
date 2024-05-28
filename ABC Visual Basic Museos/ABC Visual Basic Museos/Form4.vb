Imports System.Data.SqlClient
Imports System.Configuration
Public Class Form4
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not String.IsNullOrWhiteSpace(txtModalidad.Text) Then
            InsertTipo(txtModalidad.Text)
            LoadTipos()
        Else
            MessageBox.Show("Por favor, ingrese una modalidad.")
        End If
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadTipos()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim id As Integer
        If Integer.TryParse(txtIdTipo.Text, id) AndAlso Not String.IsNullOrWhiteSpace(txtModalidad.Text) Then
            UpdateTipo(id, txtModalidad.Text)
            LoadTipos()
        Else
            MessageBox.Show("Por favor, ingrese un identificador de tipo válido y una modalidad.")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim id As Integer
        If Integer.TryParse(txtIdTipo.Text, id) Then
            DeleteTipo(id)
            LoadTipos()
        Else
            MessageBox.Show("Por favor, seleccione un tipo válido para eliminar.")
        End If
    End Sub

    Private Sub LoadTipos()
        Dim dataTable As DataTable = GetTipos()
        DataGridView1.DataSource = dataTable
    End Sub

    Public Sub InsertTipo(modalidad As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spInsertTipo", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@Modalidad", modalidad)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function GetTipos() As DataTable
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Dim dataTable As New DataTable()
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spGetTipos", connection)
                command.CommandType = CommandType.StoredProcedure
                Using reader As SqlDataReader = command.ExecuteReader()
                    dataTable.Load(reader)
                End Using
            End Using
        End Using
        Return dataTable
    End Function

    Public Sub UpdateTipo(idTipo As Integer, modalidad As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spUpdateTipo", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdTipo", idTipo)
                command.Parameters.AddWithValue("@Modalidad", modalidad)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub DeleteTipo(idTipo As Integer)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spDeleteTipo", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdTipo", idTipo)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub dataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            txtIdTipo.Text = row.Cells("IdTipo").Value.ToString()
            txtModalidad.Text = row.Cells("Modalidad").Value.ToString()
        End If
    End Sub

End Class