Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlCustomer
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class diaCustomers

    'Dim sqlConn As String = "DRIVER={MySQL ODBC 3.51 Driver};" _
    '                          & "SERVER=localhost;" _
    '                          & "UID=root;" _
    '                          & "PWD=password;" _
    '                          & "DATABASE=wg_tms_db_90;"

    Dim sqlConn As String = frmConfig.theConnection

    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 8, FontStyle.Regular)
    Dim detailFont As New Font("Segoe UI", 6, FontStyle.Regular)

    Private Sub diaCustomers_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        diaPROJECTS.reload_cmbDB()
    End Sub

    'Private Sub diaCustomers_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
    '    If e.KeyCode = Keys.Escape Then
    '        If txt2.Enabled = True Then
    '            closeForm()
    '        Else
    '            Me.Close()
    '        End If
    '    End If
    'End Sub

    Private Sub txt2_KeyDown(sender As Object, e As KeyEventArgs) Handles txt2.KeyDown
        If e.KeyCode = Keys.Escape Then
            If txt2.Enabled = True Then
                closeForm()
            Else
                Me.Close()
            End If
        End If
    End Sub

    Private Sub diaCustomers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ValidateExisting.Text = "" Then
            Me.Text = "CUSTOMER DETAILS ENTRY"
            txt1.Enabled = True
            txt1.Focus()
            txt1.SelectAll()
        Else
            Me.Text = "CUSTOMER DETAILS ENTRY - " & ValidateExisting.Text
            load_Details()

        End If

    End Sub
    Private Sub load_Details()
        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Dim str_cust_detl As String = "SELECT dbcode, db_name1, db_name2, db_add, db_tel1, db_tel2, db_hp, db_fax, email, remark, deleted FROM m_armaster where dbcode = '" & ValidateExisting.Text.ToString & "';"

        connect.Open()
        command1 = New OdbcCommand(str_cust_detl, connect)

        myDataSet = New DataSet()
        myDataSet.Tables.Clear()
        myAdapter = New OdbcDataAdapter()
        myAdapter.SelectCommand = command1
        myAdapter.Fill(myDataSet, "db_detl")

        txt1.Text = myDataSet.Tables(0).Rows(0)("dbcode").ToString
        txt2.Text = myDataSet.Tables(0).Rows(0)("db_name1").ToString
        txt3.Text = myDataSet.Tables(0).Rows(0)("db_name2").ToString
        txt4.Text = myDataSet.Tables(0).Rows(0)("db_tel1").ToString
        txt5.Text = myDataSet.Tables(0).Rows(0)("db_tel2").ToString
        txt6.Text = myDataSet.Tables(0).Rows(0)("db_hp").ToString
        txt7.Text = myDataSet.Tables(0).Rows(0)("db_fax").ToString
        txt8.Text = myDataSet.Tables(0).Rows(0)("email").ToString
        txt9.Text = myDataSet.Tables(0).Rows(0)("remark").ToString
        txt10.Text = myDataSet.Tables(0).Rows(0)("db_add").ToString

        txt1.Enabled = False
        txt2.Focus()
        txt2.SelectAll()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Dim theCode As String = txt1.Text.ToUpper.Trim.ToString

        Dim updtCommand As OdbcCommand
        Dim updtAdapter As OdbcDataAdapter

        If txt1.Text = "" And txt2.Text = "" Then
            validation1.Visible = True
            validation2.Visible = True
            txt1.Focus()
            txt1.SelectAll()
        End If

        If txt1.Text <> "" And txt2.Text = "" Then
            validation1.Visible = False
            validation2.Visible = True
            txt2.Focus()
            txt2.SelectAll()
        End If

        If txt1.Text = "" And txt2.Text <> "" Then
            validation1.Visible = True
            validation2.Visible = False
            txt1.Focus()
            txt1.SelectAll()
        End If

        If txt1.Text <> "" And txt2.Text <> "" Then
            validation1.Visible = False
            validation2.Visible = False

            'THIS IS TO DIFFERENTIATE IF USER IS ADDING OR EDITING
            'THIS IS WHEN ADDING
            If ValidateExisting.Text = "" Then
                Try

                    connect = New OdbcConnection(sqlConn)

                    Dim inscommand2 As OdbcCommand
                    Dim insadapter2 As OdbcDataAdapter

                    Dim dC As String = txt1.Text.ToString.Trim
                    Dim dN1 As String = txt2.Text.ToString.Trim
                    Dim dN2 As String = txt3.Text.ToString.Trim
                    Dim dT1 As String = txt4.Text.ToString.Trim
                    Dim dT2 As String = txt5.Text.ToString.Trim
                    Dim dH As String = txt6.Text.ToString.Trim
                    'added
                    Dim dF As String = txt7.Text.ToString.Trim
                    Dim dE As String = txt8.Text.ToString.Trim
                    Dim dR As String = txt9.Text.ToString.Trim
                    Dim dA As String = txt10.Text.ToString.Trim

                    connect.Open()
                    Dim insCust As String = "insert into m_armaster(dbcode,db_name1,db_name2,db_tel1,db_tel2,db_hp,db_fax,email,remark,db_add) values(?,?,?,?,?,?,?,?,?,?)"
                    inscommand2 = New OdbcCommand(insCust, connect)
                    insadapter2 = New OdbcDataAdapter()

                    inscommand2.Parameters.Add("@dbcode", OdbcType.VarChar, 10).Value = dC.ToString
                    inscommand2.Parameters.Add("@db_name1", OdbcType.VarChar, 100).Value = dN1.ToString
                    inscommand2.Parameters.Add("@db_name2", OdbcType.VarChar, 100).Value = dN2.ToString
                    inscommand2.Parameters.Add("@db_tel1", OdbcType.VarChar, 12).Value = dT1.ToString
                    inscommand2.Parameters.Add("@db_tel2", OdbcType.VarChar, 12).Value = dT2.ToString
                    inscommand2.Parameters.Add("@db_hp", OdbcType.VarChar, 12).Value = dH.ToString
                    'added
                    inscommand2.Parameters.Add("@db_fax", OdbcType.VarChar, 12).Value = dF.ToString
                    inscommand2.Parameters.Add("@email", OdbcType.VarChar, 100).Value = dE.ToString
                    inscommand2.Parameters.Add("@remark", OdbcType.VarChar, 250).Value = dR.ToString
                    inscommand2.Parameters.Add("@db_add", OdbcType.VarChar, 250).Value = dA.ToString

                    insadapter2.InsertCommand = inscommand2
                    insadapter2.InsertCommand.ExecuteNonQuery()

                    MessageBox.Show("Customer Details Saved.", "ADD NEW CUSTOMER", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    txt1.Clear()
                    txt2.Clear()
                    txt3.Clear()
                    txt4.Clear()
                    txt5.Clear()
                    txt6.Clear()
                    txt7.Clear()
                    txt8.Clear()
                    txt9.Clear()
                    txt10.Clear()

                    txt1.Focus()
                    connect.Close()
                    'END OF SAVE FILE INTO DATABASE
                Catch ex As Exception
                    MessageBox.Show("Duplicated Customer Code. Please choose another Code", "ADD NEW CUSTOMER", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txt1.Focus()
                    txt1.SelectAll()
                    'MessageBox.Show(ex.ToString)
                End Try
                'connect.Close()
            End If

            'THIS IS WHEN EDITING
            If ValidateExisting.Text <> "" Then

                Dim result As Integer = MessageBox.Show("Confirm to Edit Customer Details?", "EDIT CUSTOMER", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation)
                If result = DialogResult.OK Then
                    Try

                        connect = New OdbcConnection(sqlConn)
                        connect.Open()

                        Dim dC As String = txt1.Text.ToString.Trim
                        Dim dN1 As String = txt2.Text.ToString.Trim
                        Dim dN2 As String = txt3.Text.ToString.Trim
                        Dim dT1 As String = txt4.Text.ToString.Trim
                        Dim dT2 As String = txt5.Text.ToString.Trim
                        Dim dH As String = txt6.Text.ToString.Trim
                        'added
                        Dim dF As String = txt7.Text.ToString.Trim
                        Dim dE As String = txt8.Text.ToString.Trim
                        Dim dR As String = txt9.Text.ToString.Trim
                        Dim dA As String = txt10.Text.ToString.Trim

                        'Dim updtCust As String = "update m_armaster Set db_name1 = ?, db_name2 = ?, db_tel1 = ?, db_tel2 = ?, db_hp = ? where dbcode = '" & ValidateExisting.Text.ToString.ToUpper & "';"
                        Dim updtCust As String = "update m_armaster set db_name1 = ?, db_name2 = ?, db_tel1 = ?, db_tel2 = ?, db_hp = ?, db_fax = ?, email = ?, remark = ?, db_add = ? where dbcode = '" & ValidateExisting.Text.ToString.ToUpper & "';"

                        updtCommand = New OdbcCommand(updtCust, connect)
                        updtAdapter = New OdbcDataAdapter()

                        updtCommand.Parameters.AddWithValue("@db_name1", dN1.ToString)
                        updtCommand.Parameters.AddWithValue("@db_name2", dN2.ToString)
                        updtCommand.Parameters.AddWithValue("@db_tel1", dT1.ToString)
                        updtCommand.Parameters.AddWithValue("@db_tel2", dT2.ToString)
                        updtCommand.Parameters.AddWithValue("@db_hp", dH.ToString)
                        'added
                        updtCommand.Parameters.AddWithValue("@db_fax", dF.ToString)
                        updtCommand.Parameters.AddWithValue("@email", dE.ToString)
                        updtCommand.Parameters.AddWithValue("@remark", dR.ToString)
                        updtCommand.Parameters.AddWithValue("@db_add", dA.ToString)

                        updtAdapter.UpdateCommand = updtCommand
                        updtAdapter.UpdateCommand.ExecuteNonQuery()

                        MessageBox.Show("Customer Details Saved.", "EDIT CUSTOMER", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        connect.Close()

                    Catch ex As Exception
                        MessageBox.Show("Error in saving file." & vbCrLf & vbCrLf & ex.ToString, "EDIT CUSTOMER", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    connect.Close()

                    load_Details()
                    txt1.Enabled = False
                    txt2.Enabled = False
                    txt3.Enabled = False
                    txt4.Enabled = False
                    txt5.Enabled = False
                    txt6.Enabled = False
                    txt7.Enabled = False
                    txt8.Enabled = False
                    txt9.Enabled = False
                    txt10.Enabled = False

                    SaveToolStripMenuItem.Visible = False
                    CancelToolStripMenuItem.Visible = False
                    EditToolStripMenuItem.Visible = True

                    'txt1.Clear()
                    'txt2.Clear()
                    'txt3.Clear()
                    'txt4.Clear()
                    'txt5.Clear()
                    'txt6.Clear()
                    'txt7.Clear()
                    'txt8.Clear()
                    'txt9.Clear()
                    'txt10.Clear()

                    'frmCustomers.txtDB.Text = ""
                    'Me.Close()

                ElseIf result = DialogResult.Cancel Then
                    'do nothing
                End If

            End If

        End If
    End Sub

    Private Sub CancelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelToolStripMenuItem.Click
        closeForm()

    End Sub

    Private Sub closeForm()

        If ValidateExisting.Text = "" Then
            Dim result As Integer = MessageBox.Show("Cancel changes?", "EDIT CUSTOMER", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation)
            If result = DialogResult.OK Then
                txt1.Clear()
                txt2.Clear()
                txt3.Clear()
                txt4.Clear()
                txt5.Clear()
                txt6.Clear()
                txt7.Clear()
                txt8.Clear()
                txt9.Clear()
                txt10.Clear()
                Me.Close()
            ElseIf result = DialogResult.Cancel Then
                'do nothing
            End If
        Else
            Dim result As Integer = MessageBox.Show("Cancel changes?", "EDIT CUSTOMER", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation)
            If result = DialogResult.OK Then
                load_Details()
                txt1.Enabled = False
                txt2.Enabled = False
                txt3.Enabled = False
                txt4.Enabled = False
                txt5.Enabled = False
                txt6.Enabled = False
                txt7.Enabled = False
                txt8.Enabled = False
                txt9.Enabled = False
                txt10.Enabled = False
                SaveToolStripMenuItem.Visible = False
                CancelToolStripMenuItem.Visible = False
                EditToolStripMenuItem.Visible = True

                txt2.Focus()
            ElseIf result = DialogResult.Cancel Then
                'do nothing
            End If
        End If

    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        SaveToolStripMenuItem.Visible = True
        CancelToolStripMenuItem.Visible = True
        EditToolStripMenuItem.Visible = False

        txt1.Enabled = False
        txt2.Enabled = True
        txt3.Enabled = True
        txt4.Enabled = True
        txt5.Enabled = True
        txt6.Enabled = True
        txt7.Enabled = True
        txt8.Enabled = True
        txt9.Enabled = True
        txt10.Enabled = True
    End Sub

End Class