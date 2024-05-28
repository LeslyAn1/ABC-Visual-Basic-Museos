Imports System.Data.SqlClient
Imports System.Configuration

Public Class Form5
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not String.IsNullOrWhiteSpace(txtNombre.Text) AndAlso
           Not String.IsNullOrWhiteSpace(txtUsername.Text) AndAlso
           Not String.IsNullOrWhiteSpace(txtPassword.Text) Then
            InsertUsuario(txtNombre.Text, txtUsername.Text, txtPassword.Text)
            LoadUsuarios()
        Else
            MessageBox.Show("Por favor, ingrese todos los datos del usuario.")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim id As Integer
        If Integer.TryParse(txtIdUsuario.Text, id) AndAlso
           Not String.IsNullOrWhiteSpace(txtNombre.Text) AndAlso
           Not String.IsNullOrWhiteSpace(txtUsername.Text) AndAlso
           Not String.IsNullOrWhiteSpace(txtPassword.Text) Then
            UpdateUsuario(id, txtNombre.Text, txtUsername.Text, txtPassword.Text)
            LoadUsuarios()
        Else
            MessageBox.Show("Por favor, ingrese un identificador de usuario válido y todos los datos del usuario.")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim id As Integer
        If Integer.TryParse(txtIdUsuario.Text, id) Then
            DeleteUsuario(id)
            LoadUsuarios()
        Else
            MessageBox.Show("Por favor, seleccione un usuario válido para eliminar.")
        End If
    End Sub

    Private Sub LoadUsuarios()
        Dim dataTable As DataTable = GetUsuarios()
        DataGridView1.DataSource = dataTable
    End Sub

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadUsuarios()
    End Sub

    Public Sub InsertUsuario(nombre As String, username As String, password As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spInsertUsuario", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@Nombre", nombre)
                command.Parameters.AddWithValue("@Username", username)
                command.Parameters.AddWithValue("@Password", password)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function GetUsuarios() As DataTable
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Dim dataTable As New DataTable()
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spGetUsuarios", connection)
                command.CommandType = CommandType.StoredProcedure
                Using reader As SqlDataReader = command.ExecuteReader()
                    dataTable.Load(reader)
                End Using
            End Using
        End Using
        Return dataTable
    End Function

    Public Sub UpdateUsuario(idUsuario As Integer, nombre As String, username As String, password As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spUpdateUsuario", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdUsuario", idUsuario)
                command.Parameters.AddWithValue("@Nombre", nombre)
                command.Parameters.AddWithValue("@Username", username)
                command.Parameters.AddWithValue("@Password", password)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub DeleteUsuario(idUsuario As Integer)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spDeleteUsuario", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdUsuario", idUsuario)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub dataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            txtIdUsuario.Text = row.Cells("IdUsuario").Value.ToString()
            txtNombre.Text = row.Cells("Nombre").Value.ToString()
            txtUsername.Text = row.Cells("Username").Value.ToString()
            txtPassword.Text = row.Cells("Password").Value.ToString()
        End If
    End Sub

End Class