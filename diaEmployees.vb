Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class diaEmployees

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 8, FontStyle.Regular)
    Dim detailFont As New Font("Segoe UI", 6, FontStyle.Regular)

    'Dim theEmpCode As String = ValidateExisting.Text.ToString.Trim()

    Private Sub diaEmployees_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If ValidateExisting.Text = "" Then
            'DO NOTHING
            Me.Text = "EMPLOYEE DETAILS ENTRY"
            txt1.Enabled = True
            txt1.Focus()
            txt1.SelectAll()
        Else
            Me.Text = "EMPLOYEE DETAILS ENTRY - " & ValidateExisting.Text
            'Select database and load the info into the textboxes
            load_Details()
        End If

    End Sub

    Private Sub load_Details()
        connect = New OdbcConnection(sqlConn)

        Try

            connect.Open()

            Dim selCommand As OdbcCommand
            Dim selAdapt As OdbcDataAdapter
            Dim selDS As DataSet

            Dim selString As String = "select username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, position, gender, email, ic_no, oldic_no, race, nationality, remark from m_employee where username = '" & ValidateExisting.Text.ToString.Trim() & "';"

            selCommand = New OdbcCommand(selString, connect)

            selDS = New DataSet()
            selDS.Tables.Clear()
            selAdapt = New OdbcDataAdapter()
            selAdapt.SelectCommand = selCommand
            selAdapt.Fill(selDS, "employee")

            If selDS.Tables("employee").Rows.Count <> 0 Then

                Dim dtRetrievedData As DataTable = selDS.Tables("employee")

                Dim dtData As New DataTable
                Dim dtDataRows As DataRow

                dtData.TableName = "employee"
                dtData.Columns.Add("Code")
                dtData.Columns.Add("name1")
                dtData.Columns.Add("name2")
                dtData.Columns.Add("tel1")
                dtData.Columns.Add("tel2")
                dtData.Columns.Add("hp")
                dtData.Columns.Add("wage")
                'added
                dtData.Columns.Add("pos")
                dtData.Columns.Add("gen")
                dtData.Columns.Add("mail")
                dtData.Columns.Add("ic")
                dtData.Columns.Add("oldic")
                dtData.Columns.Add("race")
                dtData.Columns.Add("nation")
                dtData.Columns.Add("remark")

                For Each dtDataRows In dtRetrievedData.Rows

                    Dim code As String = dtDataRows("username").ToString().Trim()
                    Dim name1 As String = dtDataRows("emp_name1").ToString.Trim()
                    Dim name2 As String = dtDataRows("emp_name2").ToString.Trim()
                    Dim tel1 As String = dtDataRows("emp_tel1").ToString.Trim()
                    Dim tel2 As String = dtDataRows("emp_tel2").ToString()
                    Dim hp As String = dtDataRows("emp_hp").ToString.Trim()
                    Dim wage As String = dtDataRows("emp_wage").ToString.Trim()
                    'added
                    Dim pos As String = dtDataRows("position").ToString.Trim()
                    Dim gen As String = dtDataRows("gender").ToString.Trim()
                    Dim mail As String = dtDataRows("email").ToString.Trim()
                    Dim ic As String = dtDataRows("ic_no").ToString.Trim()
                    Dim oldic As String = dtDataRows("oldic_no").ToString.Trim()
                    Dim race As String = dtDataRows("race").ToString.Trim()
                    Dim nation As String = dtDataRows("nationality").ToString.Trim()
                    Dim remark As String = dtDataRows("remark").ToString.Trim()

                    'dtData.Rows.Add(New Object() {code.ToString.Trim(), name1.ToString.Trim(), name2.ToString.Trim(), tel1.ToString.Trim(), tel2.ToString.Trim(), hp.ToString.Trim(), wage.ToString.Trim()})
                    dtData.Rows.Add(New Object() {code.ToString.Trim(), name1.ToString.Trim(), name2.ToString.Trim(), tel1.ToString.Trim(), tel2.ToString.Trim(), hp.ToString.Trim(), wage.ToString.Trim(), pos.ToString.Trim(), gen.ToString.Trim(), mail.ToString.Trim(), ic.ToString.Trim(), oldic.ToString.Trim(), race.ToString.Trim(), nation.ToString.Trim(), remark.ToString.Trim()})
                    txt1.Text = code.ToString.Trim()
                    txt1.Enabled = False
                    txt2.Focus()
                    txt2.SelectAll()

                    txt2.Text = name1.ToString.Trim()
                    txt3.Text = name2.ToString.Trim()
                    'added
                    If gen.ToString = "MALE" Then
                        cmb4.SelectedIndex = 0
                    End If
                    If gen.ToString = "FEMALE" Then
                        cmb4.SelectedIndex = 1
                    End If

                    txt5.Text = ic.ToString.Trim()
                    txt6.Text = oldic.ToString.Trim()
                    txt7.Text = race.ToString.Trim()
                    txt8.Text = nation.ToString.Trim()

                    txt9.Text = tel1.ToString.Trim()
                    txt10.Text = tel2.ToString.Trim()
                    txt11.Text = hp.ToString.Trim()
                    'added
                    txt12.Text = mail.ToString.Trim()
                    txt13.Text = pos.ToString.Trim()
                    txt15.Text = remark.ToString.Trim()

                    If frmMain.theUser = "SUPERVISOR" Then
                        txt14.Enabled = True
                        txt14.Text = wage.ToString.Trim()
                    ElseIf frmMain.theUser = txt1.Text.ToString Then
                        txt14.Enabled = False
                        txt14.Text = wage.ToString.Trim()
                    Else
                        txt14.Enabled = False
                        txt14.PasswordChar = "*"
                        txt14.Text = wage.ToString.Trim()
                    End If

                Next

            Else

                txt1.Text = ValidateExisting.Text.ToString.Trim.ToUpper
                txt1.Enabled = False
                txt2.Focus()
                txt2.SelectAll()

            End If

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        connect.Close()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Dim theCode As String = txt1.Text.ToUpper.Trim.ToString

        Dim updtCommand As OdbcCommand
        Dim updtAdapter As OdbcDataAdapter

        If txt1.Text = "" And txt2.Text = "" Then
            validation1.Visible = True
            validation2.Visible = True
        End If

        If txt1.Text <> "" And txt2.Text = "" Then
            validation1.Visible = False
            validation2.Visible = True
        End If

        If txt1.Text = "" And txt2.Text <> "" Then
            validation1.Visible = True
            validation2.Visible = False
        End If

        If txt1.Text <> "" And txt2.Text <> "" Then
            validation1.Visible = False
            validation2.Visible = False

            'IF INFORMATION IS INSIDE THE m_employee table
            If ValidateExisting.Text <> "" Then

                Dim result As Integer = MessageBox.Show("Confirm to Edit Employee Details?", "EDIT EMPLOYEE", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation)
                If result = DialogResult.OK Then
                    Try

                        connect = New OdbcConnection(sqlConn)
                        connect.Open()

                        Dim dC As String = txt1.Text.ToString.Trim
                        Dim dN1 As String = txt2.Text.ToString.Trim
                        Dim dN2 As String = txt3.Text.ToString.Trim
                        'added
                        'MsgBox(cmb4.SelectedItem.ToString)
                        Dim gen As String = cmb4.SelectedItem.ToString
                        Dim icno As String = txt5.Text.ToString
                        Dim oldicno As String = txt6.Text.ToString
                        Dim race As String = txt7.Text.ToString
                        Dim nation As String = txt8.Text.ToString

                        Dim dT1 As String = txt9.Text.ToString.Trim
                        Dim dT2 As String = txt10.Text.ToString.Trim
                        Dim dH As String = txt11.Text.ToString.Trim
                        'added
                        Dim mail As String = txt12.Text.ToString.Trim
                        Dim position As String = txt13.Text.ToString.Trim
                        Dim dL As String = txt14.Text.ToString.Trim
                        Dim remark As String = txt15.Text.ToString

                        Dim updtCust As String = "update m_employee set emp_name1 = ?, emp_name2 = ?, emp_tel1 = ?, emp_tel2 = ?, emp_hp = ?, emp_wage = ?, position = ?, gender = ?, email = ?, ic_no = ?, oldic_no = ?, race = ?, nationality = ?, remark = ? where username = '" & ValidateExisting.Text.ToString.ToUpper & "';"

                        updtCommand = New OdbcCommand(updtCust, connect)
                        updtAdapter = New OdbcDataAdapter()

                        updtCommand.Parameters.AddWithValue("@emp_name1", dN1.ToString)
                        updtCommand.Parameters.AddWithValue("@emp_name2", dN2.ToString)
                        updtCommand.Parameters.AddWithValue("@emp_tel1", dT1.ToString)
                        updtCommand.Parameters.AddWithValue("@emp_tel2", dT2.ToString)
                        updtCommand.Parameters.AddWithValue("@emp_hp", dH.ToString)
                        updtCommand.Parameters.AddWithValue("@emp_wage", dL.ToString)
                        'added
                        updtCommand.Parameters.AddWithValue("@position", position.ToString)
                        updtCommand.Parameters.AddWithValue("@gender", gen.ToString)
                        updtCommand.Parameters.AddWithValue("@email", mail.ToString)
                        updtCommand.Parameters.AddWithValue("@ic_no", icno.ToString)
                        updtCommand.Parameters.AddWithValue("@oldic_no", oldicno.ToString)
                        updtCommand.Parameters.AddWithValue("@race", race.ToString)
                        updtCommand.Parameters.AddWithValue("@nationality", nation.ToString)
                        updtCommand.Parameters.AddWithValue("@remark", remark.ToString)

                        updtAdapter.UpdateCommand = updtCommand
                        updtAdapter.UpdateCommand.ExecuteNonQuery()

                        MessageBox.Show("Employee Details Saved.", "EDIT EMPLOYEE", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        connect.Close()

                    Catch ex As Exception
                        MessageBox.Show("Error in saving file", "EDIT EMPLOYEE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    connect.Close()

                    load_Details()
                    txt1.Enabled = False
                    txt2.Enabled = False
                    txt3.Enabled = False
                    cmb4.Enabled = False
                    txt5.Enabled = False
                    txt6.Enabled = False
                    txt7.Enabled = False
                    txt8.Enabled = False
                    txt9.Enabled = False
                    txt10.Enabled = False
                    txt11.Enabled = False
                    txt12.Enabled = False
                    txt13.Enabled = False
                    txt14.Enabled = False
                    txt15.Enabled = False

                    SaveToolStripMenuItem.Visible = False
                    CancelToolStripMenuItem.Visible = False
                    EDITToolStripMenuItem.Visible = True
                    'txt1.Clear()
                    'txt2.Clear()
                    'txt3.Clear()
                    'cmb4.SelectedIndex = -1
                    'txt5.Clear()
                    'txt6.Clear()
                    'txt7.Clear()
                    'txt8.Clear()
                    'txt9.Clear()
                    'txt10.Clear()
                    'txt11.Clear()
                    'txt12.Clear()
                    'txt13.Clear()
                    'txt14.Clear()
                    'txt15.Clear()

                    'frmEmployee.txtDB.Text = ""
                    'Me.Close()

                ElseIf result = DialogResult.Cancel Then
                    'do nothing
                End If

            End If
            'frmEmployee
        End If
    End Sub

    Private Sub CancelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelToolStripMenuItem.Click
        closeForm()
    End Sub

    Private Sub closeForm()
        If ValidateExisting.Text = "" Then
            Dim result As Integer = MessageBox.Show("Cancel changes?", "EDIT EMPLOYEE", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation)
            If result = DialogResult.OK Then
                txt1.Clear()
                txt2.Clear()
                txt3.Clear()
                cmb4.SelectedIndex = -1
                txt5.Clear()
                txt6.Clear()
                txt7.Clear()
                txt8.Clear()
                txt9.Clear()
                txt10.Clear()
                txt11.Clear()
                txt12.Clear()
                txt13.Clear()
                txt14.Clear()
                txt15.Clear()
                txt1.Focus()
                Me.Close()
            ElseIf result = DialogResult.Cancel Then
                'do nothing
            End If
        Else
            Dim result As Integer = MessageBox.Show("Cancel changes?", "EDIT EMPLOYEE", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation)
            If result = DialogResult.OK Then
                load_Details()
                txt1.Enabled = False
                txt2.Enabled = False
                txt3.Enabled = False
                cmb4.Enabled = False
                txt5.Enabled = False
                txt6.Enabled = False
                txt7.Enabled = False
                txt8.Enabled = False
                txt9.Enabled = False
                txt10.Enabled = False
                txt11.Enabled = False
                txt12.Enabled = False
                txt13.Enabled = False
                txt14.Enabled = False
                txt15.Enabled = False
                SaveToolStripMenuItem.Visible = False
                CancelToolStripMenuItem.Visible = False
                EDITToolStripMenuItem.Visible = True
            ElseIf result = DialogResult.Cancel Then
                'do nothing
            End If
        End If

    End Sub

    Private Sub EDITToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EDITToolStripMenuItem.Click
        SaveToolStripMenuItem.Visible = True
        CancelToolStripMenuItem.Visible = True
        EDITToolStripMenuItem.Visible = False

        txt1.Enabled = False
        txt2.Enabled = True
        txt3.Enabled = True
        cmb4.Enabled = True
        txt5.Enabled = True
        txt6.Enabled = True
        txt7.Enabled = True
        txt8.Enabled = True
        txt9.Enabled = True
        txt10.Enabled = True
        txt11.Enabled = True
        txt12.Enabled = True
        txt13.Enabled = True
        txt14.Enabled = True
        txt15.Enabled = True

    End Sub
End Class