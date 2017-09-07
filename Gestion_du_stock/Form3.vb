Imports System
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.VisualBasic

Public Class Form3
    Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\Users\Nanovic-182\Documents\vb-Prog\Gestion_du_stock\Gestion_du_stock\db1.mdb"
    Private cnx As OleDbConnection
    Private cmd As OleDbCommand
    Private dta As OleDbDataAdapter
    Private dts As New DataSet
    Private dtt As DataTable
    Private dtr As DataRow
    Private rownum As Integer
    Private cmdb As OleDbCommandBuilder

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (TextBox5.Text = TextBox4.Text) Then
            dtr = dts.Tables("Users_Login").NewRow
            dtr("Username") = TextBox3.Text
            dtr("Password_u") = TextBox4.Text

            dtr("LastName") = TextBox1.Text
            dtr("FirstName") = TextBox2.Text




            dts.Tables("Users_Login").Rows.Add(dtr)
            cmdb = New OleDbCommandBuilder(dta)
            Try
                dta.Update(dts, "Users_Login")
                dts.Clear()
                dta.Fill(dts, "Users_Login")
                MsgBox("Ajouter avec succées")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Hide()
            Form1.Show()
            Form1.ComboBox1.Items.Add(TextBox3.Text)

        Else
            MsgBox("Password doesn't Match, try again", MsgBoxStyle.Critical, "Error")
            TextBox4.Text = ""
            TextBox5.Text = ""

        End If
        

       
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cnx = New OleDbConnection
        cnx.ConnectionString = connString
        cnx.Open()
        cmd = New OleDbCommand("Select * From Users_Login")
        dta = New OleDbDataAdapter(cmd)
        cmd.Connection() = cnx
        dta.Fill(dts, "Users_Login")
        dtt = dts.Tables("Users_Login")
        Me.Text = "Sign up"

    End Sub
End Class