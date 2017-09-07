Imports System
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.VisualBasic


Public Class Form1
    Private cnx As New OleDbConnection
    Private cmd As OleDbCommand
    Private dta As OleDbDataAdapter
    Private dts As New DataSet
    Private sql As String
    Private dtt As DataTable
    Private dtr As DataRow
    Private rownum As Integer
    Private cnxstr As String
    Private cmdb As OleDbCommandBuilder



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
      
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM Users_Login WHERE Username = '" & ComboBox1.Text & "' AND Password_u = '" & TextBox2.Text & "'", cnx)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        Dim userFound As Boolean = False
        Dim FirstName As String = ""
        Dim LastName As String = ""

        'if found:
        While dr.Read
            userFound = True
            FirstName = dr("FirstName").ToString
            LastName = dr("LastName").ToString
            Form2.Label7.Text = "Bienvenu " & FirstName & " " & LastName
        End While

        'checking the result
        If userFound = True Then
            Form2.Show()
            Me.Hide()




        Else
            MsgBox("Sorry, username or password not found", MsgBoxStyle.Critical, "Invalid Login")
            Label4.Visible = True

        End If
    End Sub

    

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        cnx = New OleDbConnection
        cnx.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; data source =C:\Users\Nanovic-182\Documents\vb-Prog\Gestion_du_stock\Gestion_du_stock\db1.mdb"
        cnx.Open()
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM Users_Login", cnx)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        Dim username As String
        While dr.Read

            username = dr("Username").ToString
            ComboBox1.Items.Add(username)
            ComboBox1.Text = "admin"


        End While







    End Sub

   
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        
        If (TextBox2.UseSystemPasswordChar = True) Then
            TextBox2.UseSystemPasswordChar = False

        ElseIf (TextBox2.UseSystemPasswordChar = False) Then
            TextBox2.UseSystemPasswordChar = True

        End If


    End Sub

   
   

    
    
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Form3.Show()

    End Sub

   

End Class
