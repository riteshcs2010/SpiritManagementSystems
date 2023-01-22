﻿Imports System.Data.SqlClient


Public Class frmCustomerLedger

    Dim a, b, c As String
   
    Sub Reset()
        dtpDateFrom.Text = Today
        dtpDateTo.Text = Today
        txtCustomerName.Text = ""
        txtCustomerID.Text = ""
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            If Len(Trim(txtCustomerName.Text)) = 0 Then
                MessageBox.Show("Please retrieve Customer Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCustomerName.Focus()
                Exit Sub
            End If
            getcustomer()
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select PartyID from LedgerBook where PartyID=@d1 and Date >=@d2 and Date < @d3"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", txtCustomerID.Text)
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d3", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date.AddDays(1)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If Not rdr.Read() Then
                MessageBox.Show("Sorry...No record found", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select Date, Name, LedgerNo, Label, Credit, Debit from LedgerBook where Date >=@d1 and Date < @d2 and PartyID=@d3 order by Date,LedgerNo", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date.AddDays(1)
            cmd.Parameters.AddWithValue("@d3", txtCustomerID.Text)
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("CustomerLedger.xml")
            Dim rpt As New rptCustomerLedger
            rpt.SetDataSource(ds)
            rpt.SetParameterValue("p1", dtpDateFrom.Value.Date)
            rpt.SetParameterValue("p2", dtpDateTo.Value.Date)
            rpt.SetParameterValue("p3", txtCustomerID.Text)
            rpt.SetParameterValue("p4", txtCustomerName.Text)
            rpt.SetParameterValue("p5", a)
            rpt.SetParameterValue("p6", b)
            rpt.SetParameterValue("p7", c)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmSalesReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
    Sub getcustomer()
        Try
            a = ""
            b = ""
            c = ""
            'txtCustomerID.Text = ""
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RTRIM(CustomerID),RTRIM(Address),RTRIM(City),RTRIM(ContactNo) FROM Customer WHERE CustomerID=@d1"
            cmd.Parameters.AddWithValue("@d1", txtCustomerID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtCustomerID.Text = rdr.GetValue(0)
                a = rdr.GetValue(1)
                b = rdr.GetValue(2)
                c = rdr.GetValue(3)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub cmbSupplierName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)
       
    End Sub

    Private Sub btnSelection_Click(sender As System.Object, e As System.EventArgs) Handles btnSelection.Click
        frmCustomerRecord.lblSet.Text = "Customer Ledger"
        'frmCustomerRecord.btnAddCustomer.Visible = False
        frmCustomerRecord.Reset()
        frmCustomerRecord.ShowDialog()
    End Sub
End Class
