Imports System.Data.SqlClient
Imports System.Configuration
Public Class Form2
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadMuseos()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not String.IsNullOrWhiteSpace(txtNombreMuseo.Text) Then
            InsertMuseo(txtNombreMuseo.Text)
            LoadMuseos()
        Else
            MessageBox.Show("Por favor, ingrese un nombre de museo.")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim id As Integer
        If Integer.TryParse(txtIdMuseo.Text, id) AndAlso Not String.IsNullOrWhiteSpace(txtNombreMuseo.Text) Then
            UpdateMuseo(id, txtNombreMuseo.Text)
            LoadMuseos()
        Else
            MessageBox.Show("Por favor, ingrese un identificador de museo válido y un nombre de museo.")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim id As Integer
        If Integer.TryParse(txtIdMuseo.Text, id) Then
            DeleteMuseo(id)
            LoadMuseos()
        Else
            MessageBox.Show("Por favor, seleccione un museo válido para eliminar.")
        End If
    End Sub

    Private Sub LoadMuseos()
        Dim dataTable As DataTable = GetMuseos()
        DataGridView1.DataSource = dataTable
    End Sub

    Public Sub InsertMuseo(nombreMuseo As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spInsertMuseo", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@NombreMuseo", nombreMuseo)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function GetMuseos() As DataTable
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Dim dataTable As New DataTable()
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spGetMuseos", connection)
                command.CommandType = CommandType.StoredProcedure
                Using reader As SqlDataReader = command.ExecuteReader()
                    dataTable.Load(reader)
                End Using
            End Using
        End Using
        Return dataTable
    End Function

    Public Sub UpdateMuseo(idMuseo As Integer, nombreMuseo As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spUpdateMuseo", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdMuseo", idMuseo)
                command.Parameters.AddWithValue("@NombreMuseo", nombreMuseo)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub DeleteMuseo(idMuseo As Integer)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand("spDeleteMuseo", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@IdMuseo", idMuseo)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub dataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            txtIdMuseo.Text = row.Cells("IdMuseo").Value.ToString()
            txtNombreMuseo.Text = row.Cells("NombreMuseo").Value.ToString()
        End If
    End Sub
End Class