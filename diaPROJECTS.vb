Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl
Imports System.Drawing.Printing

Public Class diaPROJECTS

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 8, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Dim myDataSet5 As DataSet
    Public myXML As String = Application.StartupPath & "\prnProject.XML"

    Private finalDataSet As New DataSet
    'This is the printing variables

    Public Sub btnPrintPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintPreview.Click

        checkGrid()
        PrintPreview()

    End Sub

    Public Sub PrintPreview()
        Try
            Dim jsDoc As String = txtDocCode.Text.ToString.Trim
            Dim jsRef As String = txtDocRef.Text.ToString.Trim

            prnPROJ.Label12.Text = jsDoc
            prnPROJ.Label13.Text = jsRef
            prnPROJ.Show()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        
    End Sub

    Private Sub CancelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelToolStripMenuItem.Click

        SaveToolStripMenuItem.Visible = False
        CloseProjectToolStripMenuItem.Visible = False
        EditProjectToolStripMenuItem.Visible = True
        CancelToolStripMenuItem.Visible = False
        CLOSEToolStripMenuItem.Visible = True

        txtDocCode.ReadOnly = True
        txtDocRef.ReadOnly = True
        txtRemarks.ReadOnly = True
        txtHours.ReadOnly = True
        txtMins.ReadOnly = True
        txtEstHours.Enabled = False
        txtEstHours.BackColor = Color.White
        txtEstMinutes.Enabled = False
        txtEstMinutes.BackColor = Color.White
        txtProjPrice.Enabled = False
        txtProjPrice.BackColor = Color.White

        cmbDB.Enabled = False
        cmbDB.SelectedValue = lblDebtor.Text
        cmbProj.SelectedValue = lblProjectType.Text

        cmbStatus.Enabled = False
        If cmbStatus.Enabled = False Then
            cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList
            cmbStatus.AutoCompleteMode = AutoCompleteMode.None
            cmbStatus.AutoCompleteSource = AutoCompleteSource.None
            cmbStatus.BackColor = Color.White
        Else
            cmbStatus.DropDownStyle = ComboBoxStyle.DropDown
        End If

        cmbProj.Enabled = False
        If cmbProj.Enabled = False Then
            cmbProj.DropDownStyle = ComboBoxStyle.DropDownList
            cmbProj.AutoCompleteMode = AutoCompleteMode.None
            cmbProj.AutoCompleteSource = AutoCompleteSource.None
            cmbProj.BackColor = Color.White
        Else
            cmbProj.DropDownStyle = ComboBoxStyle.DropDown
        End If

        dtpDate.Enabled = False
        dtpDate.CalendarTitleBackColor = Color.White

        btnAdd.Enabled = False
        btnAddProject.Enabled = False

        txtRemarks.TabStop = False

        GridView1.Focus()
        GridView1.Select()

        loadHeaderGrid()
        checkGrid()

    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click

        If cmbDB.SelectedIndex = -1 Then
            validation1.Visible = True
            Exit Sub
        End If

        If cmbProj.SelectedIndex = -1 Then
            validation2.Visible = True
            Exit Sub
        End If

        Dim increment As String = "1"

        If Label12.Text = "" And Label13.Text = "" Then
            Dim type As String = "DOC"
            Dim docdoc As String = txtDocCode.Text.ToString.Trim
            Dim docref As String = txtDocRef.Text.ToString.Trim
            Dim finalDocRef As String
            Dim docdate1 As DateTime = dtpDate.Value
            Dim docdate As String = docdate1.ToString("yyyy-MM-dd")
            Dim dbcoD As String = cmbDB.SelectedValue.ToString.Trim
            Dim projcoD As String = cmbProj.SelectedValue.ToString.Trim
            Dim totlhours As Integer = 0
            Dim totlmins As Integer = 0
            Dim stat As String = cmbStatus.Text.ToString.Trim

            Dim est_hr As Double
            If txtEstHours.Text = "" Then
                est_hr = Convert.ToDouble("0.00")
            Else
                est_hr = txtEstHours.Text
            End If

            Dim est_min As Double
            If txtEstMinutes.Text = "" Then
                est_min = Convert.ToDouble("0.00")
            Else
                est_min = txtEstMinutes.Text
            End If

            Dim proj_price As Double
            If txtProjPrice.Text = "" Then
                proj_price = Convert.ToDouble("0.00")
            Else
                proj_price = txtProjPrice.Text
            End If

            Try

                connect = New OdbcConnection(sqlConn)
                connect.Open()

                Dim command9 As OdbcCommand
                Dim myAdapter9 As OdbcDataAdapter
                Dim myDataSet9 As DataSet

                Dim str_find_runno = "select run_no from m_runno where doc_doc = '" & docdoc & "';"

                command9 = New OdbcCommand(str_find_runno, connect)

                myDataSet9 = New DataSet()
                myDataSet9.Tables.Clear()
                myAdapter9 = New OdbcDataAdapter()
                myAdapter9.SelectCommand = command9
                myAdapter9.Fill(myDataSet9, "runno")

                Dim theCurrentNumber As String = myDataSet9.Tables("runno").Rows(0)(0).ToString

                Dim command10 As OdbcCommand
                Dim myAdapter10 As OdbcDataAdapter
                Dim myDataSet10 As DataSet

                Dim str_max_project_runno = "select max(doc_ref) from t_projects;"

                command10 = New OdbcCommand(str_max_project_runno, connect)

                myDataSet10 = New DataSet()
                myDataSet10.Tables.Clear()
                myAdapter10 = New OdbcDataAdapter()
                myAdapter10.SelectCommand = command10
                myAdapter10.Fill(myDataSet10, "tran")

                Dim themax As String = myDataSet10.Tables("tran").Rows(0)(0).ToString

                'Dim valueOfRunnoOnTextbox As String = txtDocRef.Text.ToString.Trim
                Dim finalTrans As String
                Dim finalRunno As String

                'MessageBox.Show(themaxtranno & " / " & theCurrentNumber)
                Dim themaxtranno As String
                If themax = "" Then
                    themaxtranno = "0"
                Else
                    themaxtranno = themax
                End If

                If Convert.ToDouble(themaxtranno) = Convert.ToDouble(theCurrentNumber) Then

                    finalTrans = Val(Convert.ToDouble(theCurrentNumber) + Convert.ToDouble(increment)).ToString.Trim
                    finalRunno = Val(Convert.ToDouble(theCurrentNumber) + Convert.ToDouble(increment)).ToString.Trim

                    finalDocRef = finalTrans.PadLeft(6, "0")

                    'MessageBox.Show("Number existing already. Will add as " & finalTrans & ". The next running number will be " & finalRunno & ": " & finalDocRef)

                    Try
                        'UPDATE RUNNING NUMBER TABLE
                        Dim updtcommand7 As OdbcCommand
                        Dim updtadapter7 As OdbcDataAdapter

                        Dim update_running_no As String = "update m_runno set run_no = '" & finalRunno & "' where doc_doc = '" & docdoc.ToString.Trim & "';"

                        updtcommand7 = New OdbcCommand(update_running_no, connect)
                        updtadapter7 = New OdbcDataAdapter()

                        updtadapter7.UpdateCommand = updtcommand7
                        updtadapter7.UpdateCommand.ExecuteNonQuery()
                        'END OF UPDATE THE RUNNING NO

                        'INSERT TO T_PROJECTS
                        Dim inscommand7 As OdbcCommand
                        Dim insadapter7 As OdbcDataAdapter

                        Dim str_Project_Trans_Insert As String = "insert into t_projects(doc_type, doc_doc, doc_ref, doc_date, dbcode, project_code, " _
                                                                 & "t_projects.desc, ttl_hours, ttl_minutes, status, deleted, cancel, closed, " _
                                                                 & "proj_price, est_hours, est_minutes) VALUES ('" & type & "','" & docdoc & "', '" _
                                                                 & finalDocRef & "','" & docdate & "','" & dbcoD & "','" & projcoD & "',?" _
                                                                 & ",0.00,0.00,'" & stat & "',0,0,0," & proj_price & "," & est_hr & "," & est_min & ");"

                        inscommand7 = New OdbcCommand(str_Project_Trans_Insert, connect)
                        insadapter7 = New OdbcDataAdapter()

                        inscommand7.Parameters.AddWithValue("?", txtRemarks.Text)
                        insadapter7.InsertCommand = inscommand7
                        insadapter7.InsertCommand.ExecuteNonQuery()
                        'END OF INSERT

                        Label12.Text = docdoc.ToString.Trim.ToUpper
                        Label13.Text = docref.ToString.Trim.ToUpper
                        EditProjectToolStripMenuItem.Visible = True
                        CancelToolStripMenuItem.Visible = False
                        CLOSEToolStripMenuItem.Visible = True
                        CancelNEWToolStripMenuItem1.Visible = False
                        frmPROJECTS.disableAll()

                        If cmbStatus.SelectedIndex.ToString = 0 Then
                            CloseProjectToolStripMenuItem.Visible = False
                        ElseIf cmbStatus.SelectedIndex.ToString = 1 Then
                            CloseProjectToolStripMenuItem.Visible = True
                        Else
                            CloseProjectToolStripMenuItem.Visible = False
                        End If

                        txtDocRef.Text = theCurrentNumber.PadLeft(6, "0")
                        'frmPROJECTS.refresh_gridview1()
                    Catch ex As Exception
                        MessageBox.Show("Error: Problem in saving new project. " & vbCrLf & vbCrLf & "Error Message: " & ex.ToString)
                    End Try
                End If

                If Convert.ToDouble(themaxtranno) < Convert.ToDouble(theCurrentNumber) Then

                    finalTrans = Val(Convert.ToDouble(theCurrentNumber))
                    finalRunno = Val(Convert.ToDouble(theCurrentNumber) + Convert.ToDouble(increment)).ToString.Trim

                    finalDocRef = finalTrans.PadLeft(6, "0")

                    'MessageBox.Show("Number existing already. Will add as " & finalTrans & ". The next running number will be " & finalRunno & ": " & finalDocRef)

                    Try
                        'UPDATE RUNNING NUMBER TABLE
                        Dim updtcommand7 As OdbcCommand
                        Dim updtadapter7 As OdbcDataAdapter

                        Dim update_running_no As String = "update m_runno set run_no = '" & finalRunno & "' where doc_doc = '" & docdoc.ToString.Trim & "';"

                        updtcommand7 = New OdbcCommand(update_running_no, connect)
                        updtadapter7 = New OdbcDataAdapter()

                        updtadapter7.UpdateCommand = updtcommand7
                        updtadapter7.UpdateCommand.ExecuteNonQuery()
                        'END OF UPDATE THE RUNNING NO

                        'INSERT TO T_PROJECTS
                        Dim inscommand7 As OdbcCommand
                        Dim insadapter7 As OdbcDataAdapter

                        Dim str_Project_Trans_Insert As String = "insert into t_projects(doc_type, doc_doc, doc_ref, doc_date, dbcode, project_code, " _
                                                                 & "t_projects.desc, ttl_hours, ttl_minutes, status, deleted, cancel, closed, " _
                                                                 & "proj_price, est_hours, est_minutes) VALUES ('" & type & "','" & docdoc & "', '" _
                                                                 & finalDocRef & "','" & docdate & "','" & dbcoD & "','" & projcoD & "',?" _
                                                                 & ",0.00,0.00,'" & stat & "',0,0,0," & proj_price & "," & est_hr & "," & est_min & ");"

                        inscommand7 = New OdbcCommand(str_Project_Trans_Insert, connect)
                        insadapter7 = New OdbcDataAdapter()

                        inscommand7.Parameters.AddWithValue("?", txtRemarks.Text)
                        insadapter7.InsertCommand = inscommand7
                        insadapter7.InsertCommand.ExecuteNonQuery()
                        'END OF INSERT

                        Label12.Text = docdoc.ToString.Trim.ToUpper
                        Label13.Text = docref.ToString.Trim.ToUpper
                        EditProjectToolStripMenuItem.Visible = True
                        CancelToolStripMenuItem.Visible = False
                        CLOSEToolStripMenuItem.Visible = True
                        CancelNEWToolStripMenuItem1.Visible = False
                        frmPROJECTS.disableAll()

                        If cmbStatus.SelectedIndex.ToString = 0 Then
                            CloseProjectToolStripMenuItem.Visible = False
                        ElseIf cmbStatus.SelectedIndex.ToString = 1 Then
                            CloseProjectToolStripMenuItem.Visible = True
                        Else
                            CloseProjectToolStripMenuItem.Visible = False
                        End If

                        txtDocRef.Text = theCurrentNumber.PadLeft(6, "0")
                        'frmPROJECTS.refresh_gridview1()
                    Catch ex As Exception
                        MessageBox.Show("Error: Problem in saving new project. " & vbCrLf & vbCrLf & "Error Message: " & ex.ToString)
                    End Try
                End If

                If Convert.ToDouble(themaxtranno) > Convert.ToDouble(theCurrentNumber) Then

                    finalTrans = Val(Convert.ToDouble(theCurrentNumber))
                    finalRunno = Val(Convert.ToDouble(theCurrentNumber) + Convert.ToDouble(increment).ToString.Trim)

                    finalDocRef = finalTrans.PadLeft(6, "0")

                    Try
                        'INSERT TO T_PROJECTS
                        Dim inscommand7 As OdbcCommand
                        Dim insadapter7 As OdbcDataAdapter

                        Dim str_Project_Trans_Insert As String = "insert into t_projects(doc_type, doc_doc, doc_ref, doc_date, dbcode, project_code, " _
                                                                 & "t_projects.desc, ttl_hours, ttl_minutes, status, deleted, cancel, closed, " _
                                                                 & "proj_price, est_hours, est_minutes) VALUES ('" & type & "','" & docdoc & "', '" _
                                                                 & finalDocRef & "','" & docdate & "','" & dbcoD & "','" & projcoD & "',?" _
                                                                 & ",0.00,0.00,'" & stat & "',0,0,0," & proj_price & "," & est_hr & "," & est_min & ");"

                        inscommand7 = New OdbcCommand(str_Project_Trans_Insert, connect)
                        insadapter7 = New OdbcDataAdapter()

                        inscommand7.Parameters.AddWithValue("?", txtRemarks.Text)
                        insadapter7.InsertCommand = inscommand7
                        insadapter7.InsertCommand.ExecuteNonQuery()
                        'END OF INSERT

                        'UPDATE RUNNING NUMBER TABLE
                        Dim updtcommand7 As OdbcCommand
                        Dim updtadapter7 As OdbcDataAdapter

                        Dim update_running_no As String = "update m_runno set run_no = '" & finalRunno & "' where doc_doc = '" & docdoc.ToString.Trim & "';"

                        updtcommand7 = New OdbcCommand(update_running_no, connect)
                        updtadapter7 = New OdbcDataAdapter()

                        updtadapter7.UpdateCommand = updtcommand7
                        updtadapter7.UpdateCommand.ExecuteNonQuery()
                        'END OF UPDATE THE RUNNING NO

                        Label12.Text = docdoc.ToString.Trim.ToUpper
                        Label13.Text = docref.ToString.Trim.ToUpper
                        EditProjectToolStripMenuItem.Visible = True
                        CancelToolStripMenuItem.Visible = False
                        CLOSEToolStripMenuItem.Visible = True
                        CancelNEWToolStripMenuItem1.Visible = False

                        frmPROJECTS.disableAll()

                        If cmbStatus.SelectedIndex.ToString = 0 Then
                            CloseProjectToolStripMenuItem.Visible = False
                        ElseIf cmbStatus.SelectedIndex.ToString = 1 Then
                            CloseProjectToolStripMenuItem.Visible = True
                        Else
                            CloseProjectToolStripMenuItem.Visible = False
                        End If

                        txtDocRef.Text = theCurrentNumber.PadLeft(6, "0")
                        'frmPROJECTS.refresh_gridview1()
                    Catch ex As Exception
                        MessageBox.Show("Error: Problem in saving new project. " & vbCrLf & vbCrLf & "Error Message: " & ex.ToString)
                    End Try

                End If

                EditProjectToolStripMenuItem.Visible = True
                checkGrid()

            Catch ex As Exception
                MessageBox.Show("Problem occured during Save. Please find the following problem as per error message below." & vbCrLf & vbCrLf & ex.ToString, "Error", MessageBoxButtons.OK)
            End Try

            connect.Close()

        Else
            Dim type As String = "DOC"
            Dim docdoc As String = txtDocCode.Text.ToString.Trim
            Dim docref As String = txtDocRef.Text.ToString.Trim
            Dim finalDocRef As String = docref.PadLeft(6, "0")
            Dim docdate1 As DateTime = dtpDate.Value
            Dim docdate As String = docdate1.ToString("yyyy-MM-dd")
            Dim dbcoD As String = cmbDB.SelectedValue.ToString.Trim
            Dim projcoD As String = cmbProj.SelectedValue.ToString.Trim
            Dim projDesc As String = txtRemarks.Text.ToString.Trim
            Dim totlhours As Integer = 0
            Dim totlmins As Integer = 0
            Dim stat As String = cmbStatus.Text.ToString.Trim

            Dim est_hr As Double

            If txtEstHours.Text = "" Then
                est_hr = Convert.ToDouble("0.00")
            Else
                est_hr = txtEstHours.Text
            End If

            Dim est_min As Double

            If txtEstMinutes.Text = "" Then
                est_min = Convert.ToDouble("0.00")
            Else
                est_min = txtEstMinutes.Text
            End If

            Dim proj_price As Double
            If txtProjPrice.Text = "" Then
                proj_price = Convert.ToDouble("0.00")
            Else
                proj_price = txtProjPrice.Text
            End If

            Dim updcommand7 As OdbcCommand
            Dim updadapter7 As OdbcDataAdapter

            connect = New OdbcConnection(sqlConn)
            connect.Open()

            Try
                Dim str_Project_Trans_Update As String = "update t_projects set doc_date = '" & docdate & "', dbcode = '" & dbcoD & "', project_code = '" & projcoD & "', " _
                                                         & "t_projects.desc = ?, t_projects.status = '" & stat & "', proj_price = " & proj_price _
                                                         & ", est_hours = " & est_hr & ", est_minutes = " & est_min & " WHERE doc_doc = '" & Label12.Text.ToString.Trim.ToUpper & "' " _
                                                         & " AND doc_type = '" & type & "' AND doc_ref = '" & Label13.Text.ToString.Trim.ToUpper & "';"

                updcommand7 = New OdbcCommand(str_Project_Trans_Update, connect)
                updadapter7 = New OdbcDataAdapter()

                updcommand7.Parameters.AddWithValue("?", txtRemarks.Text)
                updadapter7.InsertCommand = updcommand7
                updadapter7.InsertCommand.ExecuteNonQuery()

                frmPROJECTS.disableAll()

                If cmbStatus.SelectedIndex.ToString = 0 Then
                    CloseProjectToolStripMenuItem.Visible = False
                ElseIf cmbStatus.SelectedIndex.ToString = 1 Then
                    CloseProjectToolStripMenuItem.Visible = True
                Else
                    CloseProjectToolStripMenuItem.Visible = False
                End If

                SaveToolStripMenuItem.Visible = False
                EditProjectToolStripMenuItem.Visible = True
                CancelToolStripMenuItem.Visible = False
                CLOSEToolStripMenuItem.Visible = True
                CancelNEWToolStripMenuItem1.Visible = False

                loadHeaderGrid()
                'loadDetailGrid()
                checkGrid()

                connect.Close()

            Catch ex As Exception
                MessageBox.Show("Problem occured during Save. Please find the following problem as per error message below." & vbCrLf & vbCrLf & ex.ToString, "Error", MessageBoxButtons.OK)
            End Try

            connect.Close()
        End If

    End Sub

    Public Sub reload_cmbDB()
        Dim command5 As OdbcCommand
        Dim myAdapter5 As OdbcDataAdapter

        connect = New OdbcConnection(sqlConn)
        connect.Open()

        Try
            'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, concat(dbcode, ' - ', db_name1) as full FROM m_armaster where deleted = 0 ORDER BY dbcode;"
            Dim str_Customers_Show As String = "SELECT dbcode, db_name1, concat(dbcode, ' - ', db_name1) as full FROM m_armaster where dbcode <> '-' and deleted = 0 ORDER BY dbcode;"

            command5 = New OdbcCommand(str_Customers_Show, connect)

            myDataSet5 = New DataSet()
            myDataSet5.Tables.Clear()
            myAdapter5 = New OdbcDataAdapter()
            myAdapter5.SelectCommand = command5
            myAdapter5.Fill(myDataSet5, "Db")

            cmbDB.DataSource = myDataSet5.Tables("Db")
            cmbDB.DisplayMember = "full"
            cmbDB.ValueMember = "dbcode"
            If cmbDB.Enabled = True Then
                'cmbDB.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                'cmbDB.DropDownStyle = ComboBoxStyle.DropDown
                'cmbDB.AutoCompleteSource = AutoCompleteSource.ListItems
                cmbDB.SelectedIndex = -1
            Else
                'NOTHING
            End If

            cmbDB.Refresh()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        connect.Close()

    End Sub

    Public Sub reload_cmbProj()
        Dim command2 As OdbcCommand
        Dim myAdapter2 As OdbcDataAdapter
        Dim myDataSet2 As DataSet

        connect = New OdbcConnection(sqlConn)
        connect.Open()

        Try
            Dim str_PROJ_Show As String = "SELECT project_code FROM m_project_type where deleted = 0 ORDER BY project_code;"

            command2 = New OdbcCommand(str_PROJ_Show, connect)

            myDataSet2 = New DataSet()
            myDataSet2.Tables.Clear()
            myAdapter2 = New OdbcDataAdapter()
            myAdapter2.SelectCommand = command2
            myAdapter2.Fill(myDataSet2, "Proj")

            cmbProj.DataSource = myDataSet2.Tables("Proj")
            cmbProj.DisplayMember = "project_code"
            cmbProj.ValueMember = "project_code"
            cmbProj.DropDownStyle = ComboBoxStyle.DropDownList
            cmbProj.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()
    End Sub

    Private Sub diaPROJECTS_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        'Dim ref_no As String = ""
        'Dim iRowIndex As Double
        'iRowIndex = Convert.ToDouble(frmPROJECTS.GridView1.CurrentRow.Index)

        'Dim selectedrowCount As Double
        'selectedrowCount = frmPROJECTS.GridView1.SelectedRows.Count

        'Dim dgvRow As DataGridViewRow

        'For Each dgvRow In frmPROJECTS.GridView1.SelectedRows
        '    Dim ref_no1 As String = dgvRow.Cells("Ref No.").RowIndex
        '    ref_no = ref_no1.ToString
        'Next

        ''MsgBox(ref_no.tostring)

        'iRowIndex = Convert.ToDouble(frmPROJECTS.GridView1.CurrentRow.Index)
        'frmPROJECTS.refresh_gridview1()
        'frmPROJECTS.GridView1.Rows(iRowIndex).Selected = True

        'Dim result As Integer = MessageBox.Show("Refresh Projects List?", "ProManage System", MessageBoxButtons.YesNo)
        'If result = DialogResult.Yes Then
        'frmPROJECTS.refresh_gridview1()

        'ElseIf result = DialogResult.No Then
        ''Do Nothing
        'End If

    End Sub

    Private Sub diaPROJECTS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If dtpDate.Enabled = False Then

                CLOSEToolStripMenuItem_Click(Nothing, Nothing)

            Else
                'Do Nothing
            End If

        End If

    End Sub

    Public Sub diaPROJECTS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        validation1.Visible = False
        validation2.Visible = False

        Me.BackColor = Color.White
        Me.ForeColor = Color.Black
        MenuStrip1.BackColor = Color.Gainsboro
        MenuStrip1.ForeColor = Color.Black

        dtpDate.Format = DateTimePickerFormat.Custom

        dtpDate.CustomFormat = "dd/MM/yyyy"
        dtpDate.DropDownAlign = LeftRightAlignment.Right

        reload_cmbDB()
        txtEstHours.PromptChar = " "
        txtEstMinutes.PromptChar = " "
        txtProjPrice.PromptChar = " "

        If Label12.Text <> "" And Label13.Text <> "" Then
            Me.Text = "Project Entry Screen - " & Label12.Text.ToString.Trim.ToUpper & "-" & Label13.Text.ToString.Trim.ToUpper
            loadHeaderGrid()
            loadDetailGrid()
            checkGrid()
        Else
            Me.Text = "Project Entry Screen"
            reload_cmbProj()
        End If

        'loadHeaderGrid()
        'loadDetailGrid()

        'checkGrid()

    End Sub

    Public Sub deleteXML()
        'If File.Exists(myXML) = True Then
        '    File.Delete(myXML)
        'Else
        '    'do nothing
        'End If

    End Sub

    Public Sub loadHeaderGrid()

            Dim command1 As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            Dim myDataSet As DataSet

            connect = New OdbcConnection(sqlConn)
            Try

                Dim str_projs = "SELECT T.doc_type, T.doc_doc AS DoC, T.doc_ref AS ReF, T.doc_date as theDate, concat(T.doc_doc, '-', T.doc_ref) AS referenceNumber, T.dbcode as custCode, D.db_name1 as custName, concat(D.dbcode, ' - ', D.db_name1) as naMe1, " _
                                    & "T.project_code as projCode, P.project_name as projName, P.project_price as projPrice, T.desc as projDesc, " _
                                    & "T.ttl_hours as totalHours, T.ttl_minutes as totalMins, T.status as projStatus, T.deleted, T.cancel as docCancel, " _
                                    & "T.transtamp, T.closed, T.proj_price, T.est_hours, T.est_minutes " _
                                    & "FROM t_projects T, m_armaster D, m_project_type P " _
                                    & "WHERE T.dbcode = D.dbcode AND T.project_code = P.project_code " _
                                    & "AND T.doc_doc = '" & Label12.Text.ToString.Trim.ToUpper & "' AND T.doc_ref = '" & Label13.Text.ToString.Trim.ToUpper & "' ;"

                connect.Open()
                command1 = New OdbcCommand(str_projs, connect)

                myDataSet = New DataSet()
                myDataSet.Tables.Clear()
                myAdapter = New OdbcDataAdapter()
                myAdapter.SelectCommand = command1
                myAdapter.Fill(myDataSet, "Proj")

                txtDocCode.Text = myDataSet.Tables("Proj").Rows(0)(1).ToString
                txtDocRef.Text = myDataSet.Tables("Proj").Rows(0)(2).ToString
                dtpDate.Value = Convert.ToDateTime(myDataSet.Tables("Proj").Rows(0)(3).ToString)
                cmbDB.Text = myDataSet.Tables("Proj").Rows(0)(7).ToString
                lblDebtor.Text = myDataSet.Tables("Proj").Rows(0)(5).ToString
                txtRemarks.Text = myDataSet.Tables("Proj").Rows(0)(11).ToString
                cmbStatus.Text = myDataSet.Tables("Proj").Rows(0)(14).ToString

                Dim CLOSED As String = myDataSet.Tables("Proj").Rows(0)(18).ToString
                Dim cancelled As String = myDataSet.Tables("Proj").Rows(0)(16).ToString

                If CLOSED = "1" Then
                    lblComplete.Visible = True
                Else
                    lblComplete.Visible = False
                End If

                If cancelled = "1" Then
                    lblCancelled.Visible = True
                Else
                    lblCancelled.Visible = False
                End If

                Dim hr As Double = Convert.ToDouble(myDataSet.Tables("Proj").Rows(0)(12).ToString)
                Dim min As Double = Convert.ToDouble(myDataSet.Tables("Proj").Rows(0)(13).ToString)
                Dim totalhour As Double = hr
                Dim totalmin As Double
                totalmin = min

                Do While totalmin >= 60
                    totalmin = totalmin - 60
                    totalhour = totalhour + 1
                Loop

                Dim esthr As Double = Convert.ToDouble(myDataSet.Tables("Proj").Rows(0)(20).ToString)
                Dim estmin As Double = Convert.ToDouble(myDataSet.Tables("Proj").Rows(0)(21).ToString)
                txtEstHours.Text = esthr.ToString
                txtEstMinutes.Text = estmin.ToString

                txtProjPrice.Text = myDataSet.Tables("Proj").Rows(0)(19).ToString

                txtHours.Text = hr.ToString
                txtMins.Text = min.ToString
                txtTotalTime.Text = totalhour & ":" & totalmin.ToString("00")

                'START OF POPULATING COMBOBOX
                Dim command2 As OdbcCommand
                Dim myAdapter2 As OdbcDataAdapter
                Dim myDataSet2 As DataSet

                Dim str_PROJ_Show As String = "SELECT project_code FROM m_project_type where deleted = 0 ORDER BY project_code;"

                command2 = New OdbcCommand(str_PROJ_Show, connect)

                myDataSet2 = New DataSet()
                myDataSet2.Tables.Clear()
                myAdapter2 = New OdbcDataAdapter()
                myAdapter2.SelectCommand = command2
                myAdapter2.Fill(myDataSet2, "Proj")

                cmbProj.DataSource = myDataSet2.Tables("Proj")
                cmbProj.DisplayMember = "project_code"
                cmbProj.ValueMember = "project_code"
                cmbProj.DropDownStyle = ComboBoxStyle.DropDownList
                cmbProj.Text = myDataSet.Tables("Proj").Rows(0)(8).ToString
                lblProjectType.Text = myDataSet.Tables("Proj").Rows(0)(8).ToString
                'cmbProj.SelectedIndex = -1
                'END OF POPULATING PROJECT COMBO BOX

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

    End Sub

    Public Sub loadDetailGrid()

        Try
            'START OF GRIDVIEW POPULATING
            Dim command3 As OdbcCommand
            Dim myAdapter3 As OdbcDataAdapter
            Dim myDataSet3 As DataSet

            Dim str_show_history As String = "SELECT D.deleted AS Deleted, T.job_date as jobDate, T.job_code, T.job_number, " _
                                             & "concat(T.job_code, '-', T.job_number) as jobRef, T.username as uName, M.emp_name1 as eName, " _
                                             & "D.time_from as timefrom, D.time_to as timeto, D.create_date as CDate, D.edit_date as EDate, D.job_desc as jobDesc " _
                                             & "FROM t_projects P, t_jobsheet_trn T, t_jobsheet_detl D, m_employee M WHERE P.doc_doc = D.doc_doc " _
                                             & "AND P.doc_ref = D.doc_ref AND T.job_code = D.job_code AND T.job_number = D.job_number AND T.username = M.username " _
                                             & "AND P.doc_doc = '" & Label12.Text.ToString.Trim.ToUpper & "' and " _
                                             & "P.doc_ref = '" & Label13.Text.ToString.Trim.ToUpper & "' and D.deleted = 0;"

            command3 = New OdbcCommand(str_show_history, connect)

            myDataSet3 = New DataSet()
            myDataSet3.Tables.Clear()
            myAdapter3 = New OdbcDataAdapter()
            myAdapter3.SelectCommand = command3
            myAdapter3.Fill(myDataSet3, "Hist")

            Dim dtRetrievedData1 As DataTable = myDataSet3.Tables("Hist")

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "ProjTable"
            dtData.Columns.Add("X", GetType(Boolean))
            dtData.Columns.Add("Date")
            dtData.Columns.Add("Job No")
            dtData.Columns.Add("User")
            dtData.Columns.Add("Name")
            dtData.Columns.Add("Job Desc")
            dtData.Columns.Add("From")
            dtData.Columns.Add("To")

            For Each dtDataRows In dtRetrievedData1.Rows

                Dim del As Boolean = dtDataRows("Deleted").ToString().Trim()
                Dim dat As Date = dtDataRows("jobDate").ToString().Trim()
                Dim ref As String = dtDataRows("jobRef").ToString.Trim()
                Dim user As String = dtDataRows("uName").ToString.Trim()
                Dim userName As String = dtDataRows("eName").ToString.Trim()
                Dim desK As String = dtDataRows("jobDesc").ToString

                Dim timeFrom As String = Convert.ToDateTime((dtDataRows("timefrom").ToString.Trim())).ToString("HH:mm")
                Dim timeTo As String = Convert.ToDateTime((dtDataRows("timeto").ToString.Trim())).ToString("HH:mm")

                'dtData.Rows.Add(New Object() {del.ToString.Trim(), close.ToString.Trim(), dat.ToString("dd-MM-yyyy"), ref.ToString.Trim(), cust.ToString.Trim(), custName.ToString.Trim(), proj.ToString.Trim(), projName.ToString.Trim(), ttlHrs.ToString.Trim(), ttlMins.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim(), theStat.ToString.Trim(), descp.ToString.Trim()})
                dtData.Rows.Add(New Object() {del.ToString.Trim(), dat.ToString("dd-MM-yyyy"), ref.ToString.Trim(), user.ToString.Trim(), userName.ToString.Trim(), desK.ToString.Trim(), timeFrom.ToString.Trim(), timeTo.ToString.Trim()})
            Next

            finalDataSet.Tables.Add(dtData)

            GridView1.DataSource = finalDataSet.Tables(0)
            GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("X").Width = 19

            GridView1.Columns.Item("Date").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Date").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Date").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Date").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Date").Width = 70

            GridView1.Columns.Item("Job No").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Job No").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Job No").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Job No").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Job No").Width = 80

            GridView1.Columns.Item("User").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("User").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("User").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("User").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("User").Width = 50

            GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Name").Width = 200

            GridView1.Columns.Item("Job Desc").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Job Desc").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Job Desc").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Job Desc").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Job Desc").Width = 220
            GridView1.Columns.Item("Job Desc").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            GridView1.Columns.Item("Job Desc").Visible = True

            GridView1.Columns.Item("From").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("From").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("From").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("From").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            GridView1.Columns.Item("From").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("From").Width = 80

            GridView1.Columns.Item("To").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("To").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("To").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("To").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            GridView1.Columns.Item("To").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("To").Width = 80

            GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.NotSet
            GridView1.Refresh()

            checkGrid()

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()
    End Sub

    Public Sub checkGrid()
        If finalDataSet.Tables(0).Rows.Count = 0 Then
            'File.Delete(myXML)
            btnPrintPreview.Visible = False
        Else
            'File.Delete(myXML)
            finalDataSet.Tables(0).WriteXml(myXML)
            If EditProjectToolStripMenuItem.Visible = False Then
                btnPrintPreview.Visible = False
            Else
                btnPrintPreview.Visible = True
            End If

        End If
    End Sub

    Private Sub cmbDB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbDB.KeyDown
        If e.KeyCode = Keys.Enter Then
            validation1.Visible = False
            cmbName.DataSource = myDataSet5.Tables("Db")
            cmbName.DisplayMember = "db_name1"
            cmbName.DropDownStyle = ComboBoxStyle.Simple
            cmbName.Enabled = False
            cmbName.BackColor = Color.White
            cmbName.ForeColor = Color.Black
        End If

        If e.KeyCode = Keys.Tab Then
            validation1.Visible = False
            cmbName.DataSource = myDataSet5.Tables("Db")
            cmbName.DisplayMember = "db_name1"
            cmbName.DropDownStyle = ComboBoxStyle.Simple
            cmbName.Enabled = False
            cmbName.BackColor = Color.White
            cmbName.ForeColor = Color.Black
        End If
    End Sub

    Private Sub cmbDB_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDB.LostFocus
        If cmbDB.SelectedIndex = -1 Then
            validation1.Visible = True
            'cmbName.DropDownStyle = ComboBoxStyle.Simple
            'cmbName.Enabled = False
            'cmbName.BackColor = Color.White
            'cmbName.ForeColor = Color.Black
            'cmbName.Text = ""
            'frmCustomers.Close()
            'frmCustomers.ShowDialog()
        Else
            validation1.Visible = False
            'cmbName.DataSource = myDataSet5.Tables("Db")
            'cmbName.DisplayMember = "db_name1"
            'cmbName.DropDownStyle = ComboBoxStyle.Simple
            'cmbName.Enabled = False
            'cmbName.BackColor = Color.White
            'cmbName.ForeColor = Color.Black
        End If
    End Sub

    Private Sub cmbProj_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbProj.LostFocus
        If cmbProj.SelectedIndex = -1 Then
            validation2.Visible = True
        Else
            validation2.Visible = False
        End If
    End Sub

    'Private Sub cmbDB_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDB.SelectedIndexChanged
    '    If cmbDB.Text = "" Then
    '        validation1.Visible = True
    '        cmbName.DropDownStyle = ComboBoxStyle.Simple
    '        cmbName.Enabled = False
    '        cmbName.BackColor = Color.White
    '        cmbName.ForeColor = Color.Black
    '        cmbName.Text = ""
    '    Else
    '        validation1.Visible = False
    '        cmbName.DataSource = myDataSet5.Tables("Db")
    '        cmbName.DisplayMember = "db_name1"
    '        cmbName.DropDownStyle = ComboBoxStyle.Simple
    '        cmbName.Enabled = False
    '        cmbName.BackColor = Color.White
    '        cmbName.ForeColor = Color.Black
    '    End If
    'End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        diaCustomers.Close()
        diaCustomers.StartPosition = FormStartPosition.CenterScreen
        diaCustomers.ShowDialog()
    End Sub

    Private Sub btnAddProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddProject.Click
        diaProjType.Close()
        diaProjType.StartPosition = FormStartPosition.CenterScreen
        diaProjType.ShowDialog()
    End Sub

    Private Sub txtProjPrice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProjPrice.Click
        txtProjPrice.SelectAll()
    End Sub

    Private Sub txtProjPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        txtProjPrice.SelectAll()
    End Sub

    Private Sub txtEstHours_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEstHours.Click
        txtEstHours.SelectAll()
    End Sub

    Private Sub txtEstHours_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEstHours.GotFocus
        txtEstHours.SelectAll()
    End Sub

    Private Sub txtEstMinutes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtEstMinutes.SelectAll()
    End Sub

    Private Sub txtEstMinutes_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        txtEstMinutes.SelectAll()
    End Sub

    Private Sub GridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellContentDoubleClick

        Dim iRowIndex As Integer

        iRowIndex = GridView1.CurrentRow.Index

        Dim jobDesc As String = GridView1.CurrentRow.Cells("Job Desc").Value

        MessageBox.Show(jobDesc, "Job Description", MessageBoxButtons.OK)
    End Sub

    Private Sub EditProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditProjectToolStripMenuItem.Click

        'CloseProjectToolStripMenuItem.Visible = False
        CancelToolStripMenuItem.Visible = True
        btnPrintPreview.Visible = False

        Try

            If lblCancelled.Visible = True Then
                MessageBox.Show("Project already Cancelled. Edit not allowed.", "Error", MessageBoxButtons.OK)
                Exit Try
            Else
                If lblComplete.Visible = True Then
                    MessageBox.Show("Project already Closed. Edit not allowed.", "Error", MessageBoxButtons.OK)
                    Exit Try
                Else

                    EditProjectToolStripMenuItem.Visible = False
                    NewToolStripMenuItem.Visible = False

                    SaveToolStripMenuItem.Visible = True
                    CLOSEToolStripMenuItem.Visible = False
                    CancelToolStripMenuItem.Visible = True

                    txtDocCode.ReadOnly = False
                    txtDocRef.ReadOnly = False
                    txtRemarks.ReadOnly = False
                    txtHours.ReadOnly = False
                    txtMins.ReadOnly = False
                    txtEstHours.Enabled = True
                    txtEstHours.BackColor = Color.White
                    txtEstMinutes.Enabled = True
                    txtEstMinutes.BackColor = Color.White
                    txtProjPrice.Enabled = True
                    txtProjPrice.BackColor = Color.White

                    cmbDB.Enabled = True
                    'If cmbDB.Enabled = False Then
                    '    cmbDB.DropDownStyle = ComboBoxStyle.DropDownList
                    '    cmbDB.AutoCompleteMode = AutoCompleteMode.None
                    '    cmbDB.AutoCompleteSource = AutoCompleteSource.None
                    '    cmbDB.BackColor = Color.White
                    'Else
                    '    cmbDB.DropDownStyle = ComboBoxStyle.DropDownList
                    'End If

                    cmbStatus.Enabled = True
                    If cmbStatus.Enabled = False Then
                        cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList
                        cmbStatus.AutoCompleteMode = AutoCompleteMode.None
                        cmbStatus.AutoCompleteSource = AutoCompleteSource.None
                        cmbStatus.BackColor = Color.White
                    Else
                        cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList
                    End If

                    Dim theStatus As String = cmbStatus.SelectedItem.ToString

                    If theStatus = "In Progress" Then
                        CloseProjectToolStripMenuItem.Visible = False
                    ElseIf theStatus = "Completed" Then
                        CloseProjectToolStripMenuItem.Visible = True
                    End If

                    cmbProj.Enabled = True
                    If cmbProj.Enabled = False Then
                        cmbProj.DropDownStyle = ComboBoxStyle.DropDownList
                        cmbProj.AutoCompleteMode = AutoCompleteMode.None
                        cmbProj.AutoCompleteSource = AutoCompleteSource.None
                        cmbProj.BackColor = Color.White
                    Else
                        cmbProj.DropDownStyle = ComboBoxStyle.DropDownList
                    End If

                    dtpDate.Enabled = True
                    dtpDate.CalendarTitleBackColor = Color.White

                    btnAdd.Enabled = True
                    btnAddProject.Enabled = True

                    txtRemarks.TabStop = True

                    dtpDate.Focus()

                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error Occured. Please check error message below. " & vbCrLf & vbCrLf & ex.ToString, "Edit Project", MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub CloseProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseProjectToolStripMenuItem.Click

        Dim theDoc As String = Label12.Text.ToString
        Dim theRef As String = Label13.Text.ToString

        Try

            If lblComplete.Visible = True Then
                MessageBox.Show("Project already closed.", "Error", MessageBoxButtons.OK)
                Exit Try
            Else
                Dim result As Integer = MessageBox.Show("Confirm Close Project?", "Close Project", MessageBoxButtons.OKCancel)
                If result = DialogResult.OK Then
                    Dim updtcommand As OdbcCommand
                    Dim updtadapter As OdbcDataAdapter

                    connect = New OdbcConnection(sqlConn)
                    connect.Open()

                    Dim del_emp As String = "update t_projects set closed = 1 where doc_doc = '" & theDoc & "' and doc_ref = '" & theRef & "';"

                    updtcommand = New OdbcCommand(del_emp, connect)
                    updtadapter = New OdbcDataAdapter()

                    updtadapter.UpdateCommand = updtcommand
                    updtadapter.UpdateCommand.ExecuteNonQuery()

                    'DELETE THE JOB
                    Dim delcommand As OdbcCommand
                    Dim deladapter As OdbcDataAdapter

                    Dim del_cust As String = "update t_jobsheet_detl set closed = '1' where doc_doc = '" & theDoc & "' and doc_ref = '" & theRef & "';"

                    delcommand = New OdbcCommand(del_cust, connect)
                    deladapter = New OdbcDataAdapter()

                    deladapter.UpdateCommand = delcommand
                    deladapter.UpdateCommand.ExecuteNonQuery()
                    'END OF DELETE

                    frmPROJECTS.disableAll()
                    EditProjectToolStripMenuItem.Visible = False
                    CloseProjectToolStripMenuItem.Visible = False
                    CancelNEWToolStripMenuItem1.Visible = False
                    CLOSEToolStripMenuItem.Visible = True
                    lblComplete.Visible = True

                    'frmPROJECTS.refresh_gridview1()

                    MessageBox.Show("Project Successfully Closed.", "Close Project", MessageBoxButtons.OK)

                ElseIf result = DialogResult.Cancel Then
                    'DO NOTHING
                End If

            End If

        Catch ex As Exception
            MessageBox.Show("Error Occured. Please check error message below. " & vbCrLf & vbCrLf & ex.ToString, "Close Project", MessageBoxButtons.OK)
        End Try


    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CancelNEWToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelNEWToolStripMenuItem1.Click
        CancelNEWToolStripMenuItem1.Visible = False
        SaveToolStripMenuItem.Visible = False
        CLOSEToolStripMenuItem.Visible = True
        frmPROJECTS.disableAll()

        dtpDate.Value = Now()
        cmbDB.SelectedIndex = -1
        cmbProj.SelectedIndex = -1
        txtRemarks.Text = ""
        txtProjPrice.Text = ""
        txtEstHours.Text = ""
        txtEstMinutes.Text = ""
        validation1.Visible = False
        validation2.Visible = False

        deleteXML()
        Me.Close()
    End Sub
End Class