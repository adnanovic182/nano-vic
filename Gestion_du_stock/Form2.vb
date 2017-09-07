Imports System
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.VisualBasic


Public Class Form2

    'For datagrid view
    Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\Users\Nanovic-182\Documents\vb-Prog\Gestion_du_stock\Gestion_du_stock\db3.mdb"
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim tables As DataTableCollection
    Dim source1 As New BindingSource

    'for 
    Dim row As Integer

    Private cnx As OleDbConnection
    Private cmd As OleDbCommand
    Private dta As OleDbDataAdapter
    Private dts As New DataSet
    Private dtt As DataTable
    Private dtr As DataRow
    Private rownum As Integer
    Private cmdb As OleDbCommandBuilder


      




    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        dtr = dts.Tables("Stock_Info").NewRow
        dtr("Code_Article") = TextBox1.Text
        dtr("Referance") = TextBox2.Text
        dtr("Quantite") = TextBox3.Text
        dtr("Prix") = TextBox4.Text
        dtr("Date_livraison") = DateTimePicker1.Value
        dtr("Observation") = TextBox5.Text
        dts.Tables("Stock_Info").Rows.Add(dtr)
        cmdb = New OleDbCommandBuilder(dta)
        dta.Update(dts, "Stock_Info")
        dts.Clear()
        dta.Fill(dts, "Stock_Info")
        Button1.PerformClick()


    End Sub

    Private Sub Form2_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        cnx = New OleDbConnection
        cnx.ConnectionString = connString
        cnx.Open()
        cmd = New OleDbCommand("Select * From Stock_Info")
        dta = New OleDbDataAdapter(cmd)
        cmd.Connection() = cnx
        dta.Fill(dts, "Stock_Info")
        dtt = dts.Tables("Stock_Info")
        Me.Text = "AeroTechnic Stock"
       
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MyConn = New OleDbConnection
        MyConn.ConnectionString = connString
        ds = New DataSet
        tables = ds.Tables
        da = New OleDbDataAdapter("Select * from Stock_Info", MyConn)
        da.Fill(ds, "Stock_Info")
        Dim view As New DataView(tables(0))
        source1.DataSource = view
        DataGridView1.DataSource = view


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        row = DataGridView1.CurrentCell.RowIndex

        dts.Tables("Stock_Info").Rows(row).Delete()
        cmdb = New OleDbCommandBuilder(dta)
        dta.Update(dts, "Stock_Info")
        Button1.PerformClick()


    End Sub
  
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        row = DataGridView1.CurrentCell.RowIndex

        dtr = dts.Tables("Stock_Info").Rows(row)
        dtr("Code_Article") = TextBox1.Text
        dtr("Referance") = TextBox2.Text
        dtr("Quantite") = TextBox3.Text
        dtr("Prix") = TextBox4.Text
        dtr("Date_livraison") = DateTimePicker1.Value
        dtr("Observation") = TextBox5.Text

        cmdb = New OleDbCommandBuilder(dta)
        dta.Update(dts, "Stock_Info")
        dts.Clear()
        dta.Fill(dts, "Stock_Info")
        Button1.PerformClick()



    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""



    End Sub

    
End Class
