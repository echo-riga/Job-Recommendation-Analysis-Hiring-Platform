
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Forms.DataVisualization.Charting
Imports MySql.Data.MySqlClient
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Drawing.Layout
Imports PdfSharp.Pdf
Imports QRCoder
Imports ClosedXML.Excel
Imports PLPFACULTY.My
Public Class Form1

    Private Sub ClearInputs(container As Control)
        For Each ctrl As Control In container.Controls
            Select Case True
                Case TypeOf ctrl Is TextBox
                    DirectCast(ctrl, TextBox).Clear()

                Case TypeOf ctrl Is ComboBox
                    DirectCast(ctrl, ComboBox).SelectedIndex = -1

                Case TypeOf ctrl Is CheckBox
                    DirectCast(ctrl, CheckBox).Checked = False

                Case TypeOf ctrl Is RadioButton
                    DirectCast(ctrl, RadioButton).Checked = False

                Case TypeOf ctrl Is DateTimePicker
                    DirectCast(ctrl, DateTimePicker).Value = DateTime.Now

                Case ctrl.HasChildren
                    ' Recursively clear nested controls like panels or group boxes
                    ClearInputs(ctrl)
            End Select
        Next
    End Sub



    Private Sub toStudentSignUpPanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toStudentSignUpPanel.LinkClicked
        studentSignUpPanel.Location = New System.Drawing.Point(530, 0)
        TextBox1.Focus()
        ClearInputs(studentSignUpPanel)

        LoadSectionsIntoComboBox()

    End Sub
    Private Sub LoadSectionsIntoComboBox()
        Dim dt As DataTable = GetAllSections()

        dt.Columns.Add("DisplayText", GetType(String))
        For Each row As DataRow In dt.Rows
            row("DisplayText") = row("section").ToString()
        Next

        studentSectionBox.DataSource = dt
        studentSectionBox.DisplayMember = "DisplayText"
        studentSectionBox.ValueMember = "id"
        studentSectionBox.SelectedIndex = -1
    End Sub


    Private Sub toMainPortal_Click(sender As Object, e As EventArgs) Handles toMainPortal.Click
        studentSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        mainPortal.Location = New System.Drawing.Point(530, 0)

    End Sub

    Private Sub toStudentSignInPanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toStudentSignInPanel.LinkClicked
        studentSignInPanel.Location = New System.Drawing.Point(530, 0)
        studentSignUpPanel.Location = New System.Drawing.Point(1000, 10000)
        TextBox13.Focus()
    End Sub

    Private Sub toMainPortal2_Click(sender As Object, e As EventArgs) Handles toMainPortal2.Click
        professorSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        mainPortal.Location = New System.Drawing.Point(530, 0)

    End Sub

    Private Sub toProfessorSignUpPanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toProfessorSignUpPanel.LinkClicked
        professorSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        professorSignUpPanel.Location = New System.Drawing.Point(530, 0)
        professorLastNameInput.Focus()
    End Sub

    Private Sub toProfessorSignInPanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toProfessorSignInPanel.LinkClicked
        professorSignInPanel.Location = New System.Drawing.Point(530, 0)
        professorSignUpPanel.Location = New System.Drawing.Point(1000, 1000)
        ClearInputs(professorSignUpPanel)
        professorLastNameInput.Focus()

    End Sub

    Private Sub toStudentQrCodePanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toStudentQrCodePanel.LinkClicked
        studentSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        studentQrCodePanel.Location = New System.Drawing.Point(530, 0)
        studentNumberHolder.Focus()

    End Sub

    Private Sub toStudentForgotPanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toStudentForgotPanel.LinkClicked
        studentSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        studentForgotPanel.Location = New System.Drawing.Point(530, 0)
        studentForgotInput.Focus()
    End Sub

    Private Sub toStudentSignInPanel2_Click(sender As Object, e As EventArgs) Handles toStudentSignInPanel2.Click
        studentSignInPanel.Location = New System.Drawing.Point(530, 0)
        studentQrCodePanel.Location = New System.Drawing.Point(1000, 1000)
        TextBox13.Focus()
    End Sub

    Private Sub toStudentSignInPanel3_Click(sender As Object, e As EventArgs) Handles toStudentSignInPanel3.Click
        studentSignInPanel.Location = New System.Drawing.Point(530, 0)
        studentForgotPanel.Location = New System.Drawing.Point(1000, 1000)
        TextBox13.Focus()
    End Sub

    Private Sub toProfessorSignInPanel2_Click(sender As Object, e As EventArgs) Handles toProfessorSignInPanel2.Click
        professorSignInPanel.Location = New System.Drawing.Point(530, 0)
        professorQrCodePanel.Location = New System.Drawing.Point(1000, 1000)
        professorUsernameInput2.Focus()
    End Sub

    Private Sub toProfessorSignInPanel3_Click(sender As Object, e As EventArgs) Handles toProfessorSignInPanel3.Click
        professorForgotPanel.Location = New System.Drawing.Point(1000, 1000)
        professorSignInPanel.Location = New System.Drawing.Point(530, 0)
        professorUsernameInput2.Focus()
    End Sub

    Private Sub toProfessorQrCodePanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toProfessorQrCodePanel.LinkClicked
        professorSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        professorQrCodePanel.Location = New System.Drawing.Point(530, 0)
        professorUsernamePasswordHolder.Focus()
    End Sub

    Private Sub toProfessorForgotPanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toProfessorForgotPanel.LinkClicked
        professorForgotPanel.Location = New System.Drawing.Point(530, 0)
        professorSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        professorForgotInput.Focus()
    End Sub

    Private Sub toMainPortal3_Click(sender As Object, e As EventArgs) Handles toMainPortal3.Click
        adminSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        mainPortal.Location = New System.Drawing.Point(530, 0)
    End Sub

    Private Sub toAdminSignInPanel_Click(sender As Object, e As EventArgs) Handles toAdminSignInPanel.Click
        adminSignInPanel.Location = New System.Drawing.Point(530, 0)
        adminQrCodePanel.Location = New System.Drawing.Point(1000, 1000)
        adminUsernameInput.Focus()
    End Sub

    Private Sub toAdminSignInPanel2_Click(sender As Object, e As EventArgs) Handles toAdminSignInPanel2.Click
        adminSignInPanel.Location = New System.Drawing.Point(530, 0)
        adminForgotPanel.Location = New System.Drawing.Point(1000, 1000)
        adminUsernameInput.Focus()
    End Sub

    Private Sub toAdminSignInPanel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toAdminSignInPanel3.LinkClicked
        adminSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        adminQrCodePanel.Location = New System.Drawing.Point(530, 0)
        adminUsernamePasswordHolder.Focus()

    End Sub

    Private Sub toAdminQrCodePanel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles toAdminQrCodePanel.LinkClicked
        adminSignInPanel.Location = New System.Drawing.Point(1000, 1000)
        adminForgotPanel.Location = New System.Drawing.Point(530, 0)
        adminForgotInput.Focus()
    End Sub


    Private Sub goToAdminPortal_MouseEnter(sender As Object, e As EventArgs) Handles goToAdminPortal.MouseEnter
        PictureBox7.BackColor = System.Drawing.Color.FromArgb(228, 228, 228)
        Label64.BackColor = System.Drawing.Color.FromArgb(228, 228, 228)
    End Sub

    Private Sub goToAdminPortal_MouseLeave(sender As Object, e As EventArgs) Handles goToAdminPortal.MouseLeave
        PictureBox7.BackColor = System.Drawing.Color.White
        Label64.BackColor = System.Drawing.Color.White
    End Sub

    Private Sub goToProfessorPortal_MouseEnter(sender As Object, e As EventArgs) Handles goToProfessorPortal.MouseEnter
        PictureBox8.BackColor = System.Drawing.Color.FromArgb(228, 228, 228)
        Label65.BackColor = System.Drawing.Color.FromArgb(228, 228, 228)
    End Sub

    Private Sub goToProfessorPortal_MouseLeave(sender As Object, e As EventArgs) Handles goToProfessorPortal.MouseLeave
        PictureBox8.BackColor = System.Drawing.Color.White
        Label65.BackColor = System.Drawing.Color.White
    End Sub

    Private Sub goToStudentPortal_MouseEnter(sender As Object, e As EventArgs) Handles goToStudentPortal.MouseEnter
        PictureBox9.BackColor = System.Drawing.Color.FromArgb(228, 228, 228)
        Label66.BackColor = System.Drawing.Color.FromArgb(228, 228, 228)
    End Sub

    Private Sub goToStudentPortal_MouseLeave(sender As Object, e As EventArgs) Handles goToStudentPortal.MouseLeave
        PictureBox9.BackColor = System.Drawing.Color.White
        Label66.BackColor = System.Drawing.Color.White
    End Sub


    Private Sub goToAdminPortal_Click_1(sender As Object, e As EventArgs) Handles goToAdminPortal.Click
        adminSignInPanel.Location = New System.Drawing.Point(530, 0)
        mainPortal.Location = New System.Drawing.Point(1000, 1000)
        adminUsernameInput.Focus()
    End Sub

    Private Sub goToProfessorPortal_Click_1(sender As Object, e As EventArgs) Handles goToProfessorPortal.Click
        professorSignInPanel.Location = New System.Drawing.Point(530, 0)
        mainPortal.Location = New System.Drawing.Point(1000, 1000)
        professorUsernameInput2.Focus()
    End Sub

    Private Sub goToStudentPortal_Click_1(sender As Object, e As EventArgs) Handles goToStudentPortal.Click
        studentSignInPanel.Location = New System.Drawing.Point(530, 0)
        mainPortal.Location = New System.Drawing.Point(1000, 1000)
        TextBox13.Focus()
    End Sub
    Private Sub ForceUppercaseInput(sender As Object, e As EventArgs)
        Dim txtBox As TextBox = CType(sender, TextBox)
        Dim selectionStart As Integer = txtBox.SelectionStart
        txtBox.Text = txtBox.Text.ToUpper()
        txtBox.SelectionStart = selectionStart ' Keep cursor position
    End Sub
    Private Sub EnsureSettingsDefaults()
        If String.IsNullOrEmpty(My.Settings.DepartmentName) Then
            My.Settings.DepartmentName = "College of Computer Studies"
        End If

        If String.IsNullOrEmpty(My.Settings.systemTitle) Then
            My.Settings.systemTitle = "Faculty Consultation System"
        End If

        My.Settings.Save()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        If String.IsNullOrEmpty(dashboardYear.Text) Then
            dashboardYear.Text = Date.Now.Year.ToString()
        End If
        EnsureSettingsDefaults()

        AddHandler studentLastNameInput.TextChanged, AddressOf ForceUppercaseInput
        AddHandler studentFirstNameInput.TextChanged, AddressOf ForceUppercaseInput
        AddHandler studentMiddleInitialInput.TextChanged, AddressOf ForceUppercaseInput
        Dim currentFont As Font = studentView.ColumnHeadersDefaultCellStyle.Font

        ' Remove underline from the current style
        Dim currentStyle As System.Drawing.FontStyle = currentFont.Style
        Dim newStyle As System.Drawing.FontStyle = currentStyle And Not System.Drawing.FontStyle.Underline

        ' Create a new font using the name, size, and cleaned-up style
        Dim newFont As New Font(currentFont.Name, currentFont.Size, newStyle)

        ' Apply the new font to the column headers
        studentView.ColumnHeadersDefaultCellStyle.Font = newFont
        consultView.ColumnHeadersDefaultCellStyle.Font = newFont
        archiveView.ColumnHeadersDefaultCellStyle.Font = newFont
        AddHandler studentEmailInput.TextChanged, AddressOf ForceUppercaseInput
        LoadSystemTitle()
        Label64.ForeColor = ColorTranslator.FromHtml("#007800")
        Label65.ForeColor = ColorTranslator.FromHtml("#007800")
        Label66.ForeColor = ColorTranslator.FromHtml("#007800")

        SetPlaceholder()
        Try
            ' Step 1: Get the latest config date
            Dim dt As DataTable = ExecuteQuery("SELECT date FROM config ORDER BY id DESC LIMIT 1")

            If dt.Rows.Count > 0 Then
                Dim configDate As Date
                If Date.TryParse(dt.Rows(0)("date").ToString(), configDate) Then
                    Dim today As Date = Date.Today

                    Console.WriteLine("Config date (parsed): " & configDate.ToString("yyyy-MM-dd"))
                    Console.WriteLine("Today's date: " & today.ToString("yyyy-MM-dd"))
                    Console.WriteLine("Dates match: " & (configDate.Date = today.Date).ToString())

                    ' Compare only the date parts (ignore time)
                    If configDate.Date = today.Date Then
                        ' Check if already run today
                        Dim dtLog As DataTable = ExecuteQuery("SELECT COUNT(*) AS cnt FROM PromotionLog WHERE runDate = CURDATE()")

                        If Convert.ToInt32(dtLog.Rows(0)("cnt")) = 0 Then
                            ' Not yet run today -> run procedure
                            ExecuteNonQuery("CALL PromoteStudents()")

                            ' ✅ Insert a log so it won’t run again today
                            ExecuteNonQuery("INSERT INTO PromotionLog (runDate) VALUES (CURDATE())")

                            MessageBox.Show("Students promoted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            Console.WriteLine("Promotion already ran today, skipping...")
                        End If
                    End If

                Else
                    Console.WriteLine("Failed to parse config date.")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error during promotion check: " & ex.Message, "Promotion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Disconnect()
        End Try
    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs)

    End Sub
    Private Sub SetActiveAdminButton(activeButton As Button)
        ' Reset all buttons to DarkGreen
        Dim buttons As Button() = {
            toAdminAdminPanel,
        toAdminDashboardPanel,
        toAdminFormPanel,
        toAdminStudentPanel,
        toAdminProfessorPanel,
        toAdminReasonPanel,
        toAdminSectionPanel,
        toAdminDatePanel,
        toAdminArchivePanel,
        toAdminHidePanel
    }

        For Each btn In buttons
            btn.BackColor = Color.DarkGreen
        Next

        ' Set clicked/active button to Green
        activeButton.BackColor = Color.Green
    End Sub


    Private Sub hideAdminPanels()
        adminAdminPanel.Location = New System.Drawing.Point(1000, 1000)
        adminDashboardPanel.Location = New System.Drawing.Point(1000, 1000)
        adminFormPanel.Location = New System.Drawing.Point(1000, 1000)
        adminStudentPanel.Location = New System.Drawing.Point(1000, 1000)
        adminProfessorPanel.Location = New System.Drawing.Point(1000, 1000)
        adminSectionPanel.Location = New System.Drawing.Point(1000, 1000)
        adminDatePanel.Location = New System.Drawing.Point(1000, 1000)
        adminReasonPanel.Location = New System.Drawing.Point(1000, 1000)
        adminArchivePanel.Location = New System.Drawing.Point(1000, 1000)
        adminHidePanel.Location = New System.Drawing.Point(1000, 1000)
    End Sub

    Private Sub toAdminDashboardPanel_Click(sender As Object, e As EventArgs) Handles toAdminDashboardPanel.Click

        LoadChart3()
        LoadChart2()
        LoadCounts()
        hideAdminPanels()
        SetActiveAdminButton(toAdminDashboardPanel)
        adminDashboardPanel.Location = New System.Drawing.Point(250, 0)
    End Sub

    Private Sub toAdminFormPanel_Click(sender As Object, e As EventArgs) Handles toAdminFormPanel.Click
        ClearInputs(adminFormPanel)
        SetActiveAdminButton(toAdminFormPanel)
        LoadReports()
        LoadProfessorsToComboBox2()
        LoadProfessorsToComboBox3()
        LoadReasonsToComboBox2()
        hideAdminPanels()
        adminFormPanel.Location = New System.Drawing.Point(250, 0)
    End Sub

    Private Sub toAdminStudentPanel_Click(sender As Object, e As EventArgs) Handles toAdminStudentPanel.Click
        SetActiveAdminButton(toAdminStudentPanel)
        ClearInputs(adminStudentPanel)
        LoadSectionsIntoComboBox3()
        LoadSectionsIntoComboBox2()
        hideAdminPanels()
        adminStudentPanel.Location = New System.Drawing.Point(250, 0)
    End Sub

    Private Sub toAdminProfessorPanel_Click(sender As Object, e As EventArgs) Handles toAdminProfessorPanel.Click
        SetActiveAdminButton(toAdminProfessorPanel)
        LoadProfessorsToGrid()
        ClearInputs(adminProfessorPanel)
        hideAdminPanels()
        adminProfessorPanel.Location = New System.Drawing.Point(250, 0)
    End Sub

    Private Sub toAdminAdminPanel_Click(sender As Object, e As EventArgs) Handles toAdminAdminPanel.Click
        SetActiveAdminButton(toAdminAdminPanel)
        ClearInputs(adminAdminPanel)
        LoadAdminsToGrid()
        hideAdminPanels()
        adminAdminPanel.Location = New System.Drawing.Point(250, 0)
    End Sub

    Private Sub toAdminHidePanel_Click(sender As Object, e As EventArgs) Handles toAdminHidePanel.Click
        SetActiveAdminButton(toAdminHidePanel)
        ClearInputs(adminHidePanel)
        LoadHiddenProfessorsToGrid()
        hideAdminPanels()
        adminHidePanel.Location = New System.Drawing.Point(250, 0)
    End Sub

    Private Sub toAdminReasonPanel_Click(sender As Object, e As EventArgs) Handles toAdminReasonPanel.Click
        SetActiveAdminButton(toAdminReasonPanel)
        LoadReasonsToGrid()
        ClearInputs(adminReasonPanel)
        hideAdminPanels()
        adminReasonPanel.Location = New System.Drawing.Point(250, 0)
    End Sub

    Private Sub toAdminSectionPanel_Click(sender As Object, e As EventArgs) Handles toAdminSectionPanel.Click
        SetActiveAdminButton(toAdminSectionPanel)
        LoadSectionsToGrid()
        ClearInputs(adminSectionPanel)
        hideAdminPanels()
        adminSectionPanel.Location = New System.Drawing.Point(250, 0)
    End Sub

    Private Sub toAdminDatePanel_Click(sender As Object, e As EventArgs) Handles toAdminDatePanel.Click
        SetActiveAdminButton(toAdminDatePanel)
        LoadReasonsToFirstThirdReasonBox()
        LoadReasonsToFourthReasonBox()
        LoadDepartmentName()
        LoadSystemTitle()
        Dim path As String = My.Settings.LogoPath
        If System.IO.File.Exists(path) Then

            logoBox.Image = Image.FromFile(path)
            logoBox.Tag = path
        End If
        LoadCurrentGraduationDate()
        hideAdminPanels()
        adminDatePanel.Location = New System.Drawing.Point(250, 0)


        Dim savedReason As String = My.Settings.firstThirdReason
        If Not String.IsNullOrWhiteSpace(savedReason) Then
            For i As Integer = 0 To firstThirdReasonBox.Items.Count - 1
                Dim row As DataRowView = CType(firstThirdReasonBox.Items(i), DataRowView)
                If row("reason").ToString() = savedReason Then
                    firstThirdReasonBox.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        Dim savedReason1 As String = My.Settings.fourthReason
        If Not String.IsNullOrWhiteSpace(savedReason1) Then
            For i As Integer = 0 To fourthReasonBox.Items.Count - 1
                Dim row As DataRowView = CType(fourthReasonBox.Items(i), DataRowView)
                If row("reason").ToString() = savedReason1 Then
                    fourthReasonBox.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
    End Sub


    Private Sub adminSignOutBtn_Click(sender As Object, e As EventArgs) Handles adminSignOutBtn.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to sign out?", "Confirm Sign Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ' Hide admin dashboard and side panels
            adminDashboard.Location = New System.Drawing.Point(1000, 1000)
            sidePanel.Location = New System.Drawing.Point(0, 0)
            mainPortal.Location = New System.Drawing.Point(530, 0)
            ClearInputs(adminSignInPanel)
            ' Clear current admin session
            CurrentAdmin = Nothing
            LoadSystemTitle()
        End If
    End Sub

    Private Sub toAdminArchivePanel_Click(sender As Object, e As EventArgs) Handles toAdminArchivePanel.Click
        SetActiveAdminButton(toAdminArchivePanel)
        LoadSectionsIntoComboBox4()
        ClearInputs(adminArchivePanel)
        LoadArchiveYears()
        LoadArchives()
        hideAdminPanels()
        adminArchivePanel.Location = New System.Drawing.Point(250, 0)
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub studentHomeSignOutBtn_Click(sender As Object, e As EventArgs) Handles studentHomeSignOutBtn.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to sign out?", "Confirm Sign Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ' Hide student home and side panels
            studentHomePanel.Location = New System.Drawing.Point(1000, 1000)
            sidePanel.Location = New System.Drawing.Point(0, 0)
            studentSignInPanel.Location = New System.Drawing.Point(530, 0)
            ClearInputs(studentSignInPanel)
            ' Clear current student session
            CurrentStudent = Nothing
        End If
    End Sub

    Private Sub studentHomeSignOutBtn_MouseEnter(sender As Object, e As EventArgs) Handles studentHomeSignOutBtn.MouseEnter
        IconPictureBox17.BackColor = ColorTranslator.FromHtml("#007800") ' Dark green on hover
    End Sub

    Private Sub studentHomeSignOutBtn_MouseLeave(sender As Object, e As EventArgs) Handles studentHomeSignOutBtn.MouseLeave
        IconPictureBox17.BackColor = System.Drawing.Color.DarkGreen ' Or original color
    End Sub

    Private Sub professorSignOutBtn_MouseEnter(sender As Object, e As EventArgs) Handles professorSignOutBtn.MouseEnter
        IconPictureBox18.BackColor = System.Drawing.ColorTranslator.FromHtml("#007800") ' Dark green on hover
    End Sub

    Private Sub professorSignOutBtn_MouseLeave(sender As Object, e As EventArgs) Handles professorSignOutBtn.MouseLeave
        IconPictureBox18.BackColor = System.Drawing.Color.DarkGreen ' Or original color
    End Sub

    Private Sub professorGenerateReportBtn_MouseEnter(sender As Object, e As EventArgs) Handles professorGenerateReportBtn.MouseEnter
        IconPictureBox22.BackColor = System.Drawing.ColorTranslator.FromHtml("#007800") ' Dark green on hover
    End Sub

    Private Sub professorGenerateReportBtn_MouseLeave(sender As Object, e As EventArgs) Handles professorGenerateReportBtn.MouseLeave
        IconPictureBox22.BackColor = System.Drawing.Color.DarkGreen ' Or original color
    End Sub


    Private Sub LoadStatusIntoComboBox()
        Dim dt As New DataTable()
        dt.Columns.Add("Id", GetType(String))
        dt.Columns.Add("DisplayText", GetType(String))

        ' Add status options (lowercase values for DB)
        dt.Rows.Add("regular", "Regular")
        dt.Rows.Add("irregular", "Irregular")

        studentHomeStatusBox.DataSource = dt
        studentHomeStatusBox.DisplayMember = "DisplayText"
        studentHomeStatusBox.ValueMember = "Id"

        ' Default select the student's current status (lowercase)
        If Not String.IsNullOrEmpty(CurrentStudent.Status) Then
            studentHomeStatusBox.SelectedValue = CurrentStudent.Status.ToLower()

        Else
            studentHomeStatusBox.SelectedIndex = -1 ' nothing selected if null
        End If
    End Sub

    Private Sub LoadSectionsIntoComboBox7()
        Dim dt As DataTable = GetAllSections()

        ' Add DisplayText column to match section name in studentView
        dt.Columns.Add("DisplayText", GetType(String))
        For Each row As DataRow In dt.Rows
            row("DisplayText") = row("section").ToString()
        Next

        studentHomeSectionBox.DataSource = dt
        studentHomeSectionBox.DisplayMember = "DisplayText"
        studentHomeSectionBox.ValueMember = "id"

        ' Pre-select based on CurrentStudent.section (name like BSIT2A)
        If Not String.IsNullOrEmpty(CurrentStudent.Section) Then
            studentHomeSectionBox.SelectedIndex =
            studentHomeSectionBox.FindStringExact(CurrentStudent.Section.Trim())
        Else
            studentHomeSectionBox.SelectedIndex = -1
        End If
    End Sub



    Private Sub studentSignInBtn_Click(sender As Object, e As EventArgs) Handles studentSignInBtn.Click
        ' Construct student number from input textboxes
        Dim part1 As String = TextBox13.Text.Trim() & TextBox14.Text.Trim()
        Dim part2 As String = TextBox15.Text.Trim() & TextBox16.Text.Trim() & TextBox17.Text.Trim() & TextBox18.Text.Trim() & TextBox24.Text.Trim()
        Dim studentNumber As String = part1 & "-" & part2

        ' Validate complete input
        If part1.Length <> 2 OrElse part2.Length <> 5 Then
            MessageBox.Show("Please complete the student number correctly.", "Validation Error")
            Exit Sub
        End If

        ' Check student record from DB
        Dim student As StudentModel = GetStudentByNumber(studentNumber)

        If student IsNot Nothing Then

            CurrentStudent = student

            studentSignInPanel.Location = New System.Drawing.Point(1000, 1000)
            sidePanel.Location = New System.Drawing.Point(1000, 1000)
            studentHomePanel.Location = New System.Drawing.Point(0, 0)


            Dim fullName As String = CurrentStudent.FirstName

            ' Add middle initial if it exists
            If Not String.IsNullOrWhiteSpace(CurrentStudent.MiddleInitial) Then
                fullName &= " " & CurrentStudent.MiddleInitial & "."
            End If

            ' Add last name
            fullName &= " " & CurrentStudent.LastName

            ' Optionally add suffix
            If Not String.IsNullOrWhiteSpace(CurrentStudent.Suffix) Then
                fullName &= " " & CurrentStudent.Suffix
            End If
            ClearInputs(studentHomePanel)
            studentNameText.Text = fullName & "!"
            studentNumberText.Text = CurrentStudent.StudentNumber
            LoadSectionsIntoComboBox7()
            LoadStatusIntoComboBox()
            studentTimeInText.Text = "Time IN: " & DateTime.Now.ToString("hh:mm tt")

            LoadProfessorsToComboBox()
            LoadReasonsToComboBox()
            SetPlaceholder()
        Else
            MessageBox.Show("Student number does not exist.", "Sign-In Failed")
        End If
    End Sub


    Private Sub adminSignInBtn_Click(sender As Object, e As EventArgs) Handles adminSignInBtn.Click


        Dim username As String = adminUsernameInput.Text.Trim()
        Dim password As String = adminPasswordInput.Text.Trim()

        If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then
            MessageBox.Show("Please enter both username and password.", "Validation Error")
            Exit Sub
        End If

        Dim admin As AdminModel = LoginAdmin(username, password)

        If admin IsNot Nothing Then

            CurrentAdmin = admin

            adminSignInPanel.Location = New System.Drawing.Point(1000, 1000)
            sidePanel.Location = New System.Drawing.Point(1000, 1000)
            adminDashboard.Location = New System.Drawing.Point(0, 0)
            adminDashboardPanel.Location = New System.Drawing.Point(250, 0)

            LoadChart2()
            LoadCounts()
            LoadChart3()

            adminNameLabel.Text = CurrentAdmin.FirstName &
    If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".") &
    " " & CurrentAdmin.LastName &
    If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", ", " & CurrentAdmin.Suffix)

        Else
            MessageBox.Show("Invalid username or password.", "Login Failed")
        End If
    End Sub

    Private Sub professorSignInBtn_Click(sender As Object, e As EventArgs) Handles professorSignInBtn.Click
        Dim username As String = professorUsernameInput2.Text.Trim()
        Dim password As String = professorPasswordInput2.Text.Trim()

        If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then
            MessageBox.Show("Please enter both username and password.", "Validation Error")
            Exit Sub
        End If

        Dim prof As ProfessorModel = LoginProfessor(username, password)

        If prof IsNot Nothing Then
            CurrentProfessor = prof
            LoadSectionsIntoComboBox8()
            professorSignInPanel.Location = New System.Drawing.Point(1000, 1000)
            sidePanel.Location = New System.Drawing.Point(1000, 1000)
            professorHomePanel.Location = New System.Drawing.Point(0, 0)

            professorSurnameText.Text = "Welcome, Prof. " & CurrentProfessor.LastName & "!"
            professorDateText.Text = DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")
            LoadProfessorCounts()
            LoadConsultations()
        Else
            MessageBox.Show("Invalid username or password.", "Login Failed")
        End If
    End Sub



    Private Sub professorSignOutBtn_Click(sender As Object, e As EventArgs) Handles professorSignOutBtn.Click


        Dim result As DialogResult = MessageBox.Show("Are you sure you want to sign out?", "Confirm Sign Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ' Hide student home and side panels
            professorHomePanel.Location = New System.Drawing.Point(1000, 1000)
            sidePanel.Location = New System.Drawing.Point(0, 0)
            mainPortal.Location = New System.Drawing.Point(530, 0)

            ClearInputs(professorSignInPanel)
            ' Clear current student session
            CurrentProfessor = Nothing
        End If
    End Sub

    Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress, TextBox6.KeyPress, TextBox5.KeyPress, TextBox4.KeyPress, TextBox3.KeyPress, TextBox2.KeyPress, TextBox1.KeyPress _


        ' Allow only digits and backspace
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
            Return
        End If
    End Sub

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged, TextBox6.TextChanged, TextBox5.TextChanged, TextBox4.TextChanged, TextBox3.TextChanged, TextBox2.TextChanged, TextBox1.TextChanged


        Dim txt As TextBox = CType(sender, TextBox)
        If txt.Text.Length = 1 Then
            ' Move to next textbox
            Me.SelectNextControl(txt, True, True, True, True)
        End If
    End Sub

    Private Sub TextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown, TextBox6.KeyDown, TextBox5.KeyDown, TextBox4.KeyDown, TextBox3.KeyDown, TextBox2.KeyDown, TextBox1.KeyDown


        Dim txt As TextBox = CType(sender, TextBox)
        If e.KeyCode = Keys.Back AndAlso txt.SelectionStart = 0 Then
            ' Move to previous textbox
            Me.SelectNextControl(txt, False, True, True, True)
        End If
    End Sub

    ' Allow only digits and backspace
    Private Sub TextBox_KeyPress_StudentNumber2(sender As Object, e As KeyPressEventArgs) Handles TextBox24.KeyPress, TextBox18.KeyPress, TextBox17.KeyPress, TextBox16.KeyPress, TextBox15.KeyPress, TextBox14.KeyPress, TextBox13.KeyPress


        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    ' Auto-move to next TextBox when one digit is typed
    Private Sub TextBox_TextChanged_StudentNumber2(sender As Object, e As EventArgs) Handles TextBox24.TextChanged, TextBox18.TextChanged, TextBox17.TextChanged, TextBox16.TextChanged, TextBox15.TextChanged, TextBox14.TextChanged, TextBox13.TextChanged


        Dim txt As TextBox = CType(sender, TextBox)
        If txt.Text.Length = 1 Then
            Me.SelectNextControl(txt, True, True, True, True)
        End If
    End Sub

    ' Move back when pressing Backspace on empty TextBox
    Private Sub TextBox_KeyDown_StudentNumber2(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown, TextBox18.KeyDown, TextBox17.KeyDown, TextBox16.KeyDown, TextBox15.KeyDown, TextBox14.KeyDown, TextBox13.KeyDown


        Dim txt As TextBox = CType(sender, TextBox)
        If e.KeyCode = Keys.Back AndAlso txt.SelectionStart = 0 Then
            Me.SelectNextControl(txt, False, True, True, True)
        End If
    End Sub

    Private Sub studentSignUpBtn_Click(sender As Object, e As EventArgs) Handles studentSignUpBtn.Click
        ' Build the student number
        Dim part1 As String = TextBox1.Text.Trim() & TextBox2.Text.Trim()
        Dim part2 As String = TextBox3.Text.Trim() & TextBox4.Text.Trim() & TextBox5.Text.Trim() & TextBox6.Text.Trim() & TextBox7.Text.Trim()
        Dim studentNumber As String = part1 & "-" & part2

        ' Get input values
        Dim firstName As String = studentFirstNameInput.Text.Trim()
        Dim lastName As String = studentLastNameInput.Text.Trim()
        Dim middleInitial As String = studentMiddleInitialInput.Text.Trim()
        Dim suffix As String = studentSuffixBox.Text.Trim()
        Dim email As String = studentEmailInput.Text.Trim()
        Dim status As String = "regular"

        ' Validate section
        If studentSectionBox.SelectedIndex = -1 Then
            MessageBox.Show("Please select a section.", "Validation Error")
            Exit Sub
        End If

        Dim sectionId As Integer = Convert.ToInt32(studentSectionBox.SelectedValue)

        ' Validate input
        If Not System.Text.RegularExpressions.Regex.IsMatch(studentNumber, "^\d{2}-\d{5}$") OrElse
       String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) Then

            MessageBox.Show("Please enter valid and complete information.", "Validation Error")
            Exit Sub
        End If

        ' Try to insert the student
        Dim success As Boolean = InsertStudent(studentNumber, firstName, lastName, middleInitial, suffix, email, sectionId, status)

        If success Then
            MessageBox.Show("Student account successfully signed up!", "Success")
            ClearInputs(studentSignUpPanel)

            ' ✅ Send QR with logo in center
            Try
                Cursor.Current = Cursors.WaitCursor

                ' Generate QR
                Dim qrContent As String = $"[{studentNumber}]"
                Dim qrGenerator As New QRCoder.QRCodeGenerator()
                Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
                Dim qrCode = New QRCoder.QRCode(qrData)
                Dim qrImage = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

                ' Load and resize logo
                ' Use embedded resource (no need for path or File.Exists)
                Dim logo As Image = My.Resources.PLP

                ' Resize logo
                Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
                Dim resizedLogo As New Bitmap(logoSize, logoSize)
                Using g As Graphics = Graphics.FromImage(resizedLogo)
                    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(logo, 0, 0, logoSize, logoSize)
                End Using


                ' Overlay logo in center of QR
                Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
                Using g As Graphics = Graphics.FromImage(qrWithLogo)
                    g.Clear(Color.White)
                    g.DrawImage(qrImage, 0, 0)

                    Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                    Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                    Dim padding As Integer = 6
                    Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                    g.FillRectangle(Brushes.White, whiteRect)
                    g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
                End Using

                ' Convert to stream
                Dim ms As New MemoryStream()
                qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
                ms.Position = 0

                ' Email with QR
                Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
                Dim toAddress As New MailAddress(email)
                Dim fromPassword As String = "swlqbwgztcqbneuw"

                Dim message As New MailMessage(fromAddress, toAddress)
                message.Subject = "Your Student QR Code"
                message.Body = $"Hi {firstName}," & vbCrLf & vbCrLf &
                           "Attached is your student QR code." & vbCrLf &
                           "You may use it for login or authentication." & vbCrLf

                message.Attachments.Add(New Attachment(ms, "StudentQRCode.png", "image/png"))

                Dim smtp As New SmtpClient("smtp.gmail.com", 587)
                smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
                smtp.EnableSsl = True
                smtp.Send(message)

                MessageBox.Show("QR Code sent to student's email.", "QR Email Sent")
                studentSignInPanel.Location = New System.Drawing.Point(530, 0)
                studentSignUpPanel.Location = New System.Drawing.Point(1000, 10000)
                TextBox13.Focus()

            Catch ex As Exception
                MessageBox.Show("Failed to send QR email: " & ex.Message, "Email Error")
            Finally
                Cursor.Current = Cursors.Default
            End Try
        Else
            MessageBox.Show("Failed to sign up student account.", "Sign-Up Failed")
        End If
    End Sub


    Private Function IsValidName(name As String) As Boolean
        ' Allows letters, spaces, and optional trailing period
        Return System.Text.RegularExpressions.Regex.IsMatch(name, "^[A-Za-z\s]+\.?$")
    End Function


    Private Function IsValidEmail(email As String) As Boolean
        Dim isValid As Boolean = System.Text.RegularExpressions.Regex.IsMatch(
        email,
        "^[^@\s]+@plpasig\.edu\.ph$",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase
    )

        If Not isValid Then
            MessageBox.Show("❌ Invalid email format." & vbCrLf &
                        "✅ Only emails ending with @plpasig.edu.ph are allowed.",
                        "Email Validation Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
        End If

        Return isValid
    End Function

    Private Sub studentSubmitFormBtn_MouseEnter(sender As Object, e As EventArgs) Handles studentSubmitFormBtn.MouseEnter
        PictureBox11.BackColor = ColorTranslator.FromHtml("#007800") ' Green on hover
    End Sub

    Private Sub studentSubmitFormBtn_MouseLeave(sender As Object, e As EventArgs) Handles studentSubmitFormBtn.MouseLeave
        PictureBox11.BackColor = System.Drawing.Color.DarkGreen ' Or original color if different
    End Sub

    Private Sub professorSignUpBtn_Click(sender As Object, e As EventArgs) Handles professorSignUpBtn.Click
        ' Collect inputs
        Dim lastName As String = professorLastNameInput.Text.Trim()
        Dim firstName As String = professorFirstNameInput.Text.Trim()
        Dim middleInitial As String = professorMiddleInitialInput.Text.Trim()
        Dim suffix As String = professorSuffixBox.Text.Trim()
        Dim email As String = professorEmailInput.Text.Trim()
        Dim username As String = professorUsernameInput.Text.Trim()
        Dim password As String = professorPasswordInput.Text.Trim()

        ' Basic combined validation
        If String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       String.IsNullOrWhiteSpace(username) OrElse
       String.IsNullOrWhiteSpace(password) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       Not IsValidEmail(email) Then

            MessageBox.Show("Please enter valid and complete information.", "Validation Error")
            Exit Sub
        End If

        ' Insert to DB
        Dim success As Boolean = InsertProfessor(lastName, firstName, middleInitial, suffix, email, username, password)

        If success Then
            MessageBox.Show("Professor account successfully signed up!", "Success")
            ClearInputs(professorSignUpPanel)

            ' ✅ Generate and send QR code
            Try
                Cursor.Current = Cursors.WaitCursor

                ' QR code content
                Dim qrContent As String = $"[{username}][{password}]"
                Dim qrGenerator As New QRCoder.QRCodeGenerator()
                Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
                Dim qrCode = New QRCoder.QRCode(qrData)
                Dim qrImage As Bitmap = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

                ' Use embedded resource (no need for path or File.Exists)
                Dim logo As Image = My.Resources.PLP

                ' Resize logo
                Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
                Dim resizedLogo As New Bitmap(logoSize, logoSize)
                Using g As Graphics = Graphics.FromImage(resizedLogo)
                    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(logo, 0, 0, logoSize, logoSize)
                End Using


                ' Merge QR and logo
                Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
                Using g As Graphics = Graphics.FromImage(qrWithLogo)
                    g.Clear(Color.White)
                    g.DrawImage(qrImage, 0, 0)

                    Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                    Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                    Dim padding As Integer = 6
                    Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                    g.FillRectangle(Brushes.White, whiteRect)
                    g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
                End Using

                ' Convert to MemoryStream
                Dim ms As New MemoryStream()
                qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
                ms.Position = 0

                ' Send Email
                Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
                Dim toAddress As New MailAddress(email)
                Dim fromPassword As String = "swlqbwgztcqbneuw"

                Dim message As New MailMessage(fromAddress, toAddress)
                message.Subject = "Your Professor QR Code"
                message.Body = $"Hi {firstName}," & vbCrLf & vbCrLf &
                        "Attached is your QR code containing your login credentials." & vbCrLf &
                        "Please scan it for convenient login." & vbCrLf & vbCrLf &
                        "Thank you."

                message.Attachments.Add(New Attachment(ms, "ProfessorQRCode.png", "image/png"))

                Dim smtp As New SmtpClient("smtp.gmail.com", 587)
                smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
                smtp.EnableSsl = True
                smtp.Send(message)

                MessageBox.Show("QR code sent to professor's email.", "Email Sent")
            Catch ex As Exception
                MessageBox.Show("Failed to send QR email: " & ex.Message, "Email Error")
            Finally
                Cursor.Current = Cursors.Default
            End Try

        Else
            MessageBox.Show("Failed to sign up professor account.", "Sign-Up Failed")
        End If
    End Sub


    Private Sub LoadProfessorsToComboBox()
        ' Get all professors from database
        Dim dt As DataTable = GetAllProfessors()

        ' Add computed column for formatted full name
        dt.Columns.Add("FullName", GetType(String))

        For Each row As DataRow In dt.Rows
            Dim lastName As String = row("last_name").ToString().Trim()
            Dim firstName As String = row("first_name").ToString().Trim()
            Dim mi As String = If(String.IsNullOrWhiteSpace(row("middle_initial").ToString()), "", " " & row("middle_initial").ToString() & ".")
            Dim suffix As String = If(String.IsNullOrWhiteSpace(row("suffix").ToString()), "", " " & row("suffix").ToString())

            ' Format: LastName, FirstName M. Suffix
            row("FullName") = $"{lastName}, {firstName}{mi}{suffix}".Trim()
        Next

        ' Bind to ComboBox
        professorBox.DataSource = dt
        professorBox.DisplayMember = "FullName" ' What user sees
        professorBox.ValueMember = "id"         ' Hidden value for FK
        professorBox.SelectedIndex = -1         ' No selection initially
    End Sub


    Private Sub LoadReasonsToComboBox()
        Dim dt As DataTable = GetAllReasons()

        reasonBox.DataSource = dt
        reasonBox.DisplayMember = "reason"
        reasonBox.ValueMember = "id"
        reasonBox.SelectedIndex = -1

        ' Check student year level
        If CurrentStudent IsNot Nothing AndAlso CurrentStudent.Section.Length >= 5 Then
            Dim yearChar As Char = CurrentStudent.Section(4)

            ' Toggle graduationCap visibility
            graduationCap.Visible = (yearChar = "4"c)

            ' Determine reason setting
            Dim targetReason As String = If(yearChar = "4"c, My.Settings.fourthReason, My.Settings.firstThirdReason)

            ' If "Empty", skip selection
            If targetReason = "Empty" Then
                reasonBox.SelectedIndex = -1
                Exit Sub
            End If

            ' Try to find and select the matching reason
            For i As Integer = 0 To reasonBox.Items.Count - 1
                Dim row As DataRowView = CType(reasonBox.Items(i), DataRowView)
                If row("reason").ToString() = targetReason Then
                    reasonBox.SelectedIndex = i
                    Exit For
                End If
            Next
        Else
            graduationCap.Visible = False
        End If
    End Sub



    Private Sub reasonBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles reasonBox.SelectedIndexChanged
        Dim selectedRow As DataRowView = TryCast(reasonBox.SelectedItem, DataRowView)

        If selectedRow IsNot Nothing Then
            Dim isSpecial As Boolean = Convert.ToBoolean(selectedRow("is_special"))
            professorBox.Enabled = Not isSpecial ' Disable if special is true
        End If
    End Sub
    Private Sub studentSubmitFormBtn_Click(sender As Object, e As EventArgs) Handles studentSubmitFormBtn.Click
        ' Validate reason selection
        If reasonBox.SelectedIndex = -1 OrElse reasonBox.SelectedValue Is Nothing Then
            MessageBox.Show("Please select a reason for consultation.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim selectedReasonId As Object = reasonBox.SelectedValue

        ' Get the is_special value from the selected reason
        Dim isSpecial As Boolean = Convert.ToBoolean(CType(reasonBox.SelectedItem, DataRowView)("is_special"))

        ' Only require professor if NOT special
        Dim selectedProfessorId As Object = Nothing
        If Not isSpecial Then
            If professorBox.SelectedIndex = -1 OrElse professorBox.SelectedValue Is Nothing Then
                MessageBox.Show("Please select a professor (required for this reason).", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            selectedProfessorId = professorBox.SelectedValue
        End If

        ' Get student message (optional)
        Dim studentMessage As String = studentMessageInput.Text.Trim()

        ' If the message is still the placeholder, treat it as empty
        If studentMessage = "Describe your concern" Then
            studentMessage = ""
        End If


        ' Parse Time In from textbox (already assumed valid)
        Dim timePart As String = studentTimeInText.Text.Replace("Time IN: ", "").Trim()
        Dim studentTimeIn As TimeSpan = DateTime.Parse(timePart).TimeOfDay


        ' Get today's date
        Dim consultationDate As Date = Date.Today

        Dim result = MessageBox.Show("Are you sure you want to submit this consultation request?", "Confirm Submission", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result <> DialogResult.Yes Then Exit Sub

        Dim studentTimeOut As TimeSpan = DateTime.Now.TimeOfDay

        ' Get student ID
        Dim studentId As Integer = CurrentStudent.Id

        Report.InsertReport(studentId, selectedReasonId, studentMessage, consultationDate, studentTimeIn, studentTimeOut, selectedProfessorId)

        ' --- After successful insert, check for section/status changes ---
        ' Compare by name but update by ID
        Dim newSectionName As String = studentHomeSectionBox.Text.Trim()
        Dim newSectionId As Integer = Convert.ToInt32(studentHomeSectionBox.SelectedValue)



        ' display text like "BSIT2A"
        Dim newStatus As String = studentHomeStatusBox.SelectedValue?.ToString().Trim().ToLower() ' "regular" / "irregular"

        Dim updates As New List(Of String)

        ' Compare section
        If Not String.IsNullOrEmpty(newSectionName) AndAlso CurrentStudent.Section <> newSectionName Then
            updates.Add($"section = {newSectionId}")
            CurrentStudent.Section = newSectionName ' keep readable name in memory
        End If

        ' Compare status
        If Not String.IsNullOrEmpty(newStatus) AndAlso CurrentStudent.Status <> newStatus Then
            updates.Add($"status = '{MySqlHelper.EscapeString(newStatus)}'")
            CurrentStudent.Status = newStatus ' update in memory
        End If

        ' Apply update if needed
        If updates.Count > 0 Then
            Dim updateSql As String = $"UPDATE students SET {String.Join(", ", updates)} WHERE id = {studentId};"
            ExecuteNonQuery(updateSql)
            MessageBox.Show("Your profile information (section/status) was updated.", "Update Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' --- Clear form + UI reset ---
        MessageBox.Show("Consultation report submitted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ClearInputs(studentHomePanel)
        ClearInputs(studentSignInPanel)
        studentHomePanel.Location = New System.Drawing.Point(1000, 1000)
        sidePanel.Location = New System.Drawing.Point(0, 0)
        studentSignInPanel.Location = New System.Drawing.Point(530, 0)
    End Sub


    Private Sub studentNameSearch_TextChanged(sender As Object, e As EventArgs)
        LoadReports()
    End Sub

    Private Sub professorNameSearch_TextChanged(sender As Object, e As EventArgs)
        LoadReports()
    End Sub
    Private Sub LoadReports()
        Dim fromDate As Date = reportFromDate.Value.Date
        Dim toDate As Date = reportToDate.Value.Date

        ' Get selected professor ID from ComboBox (0 = All Professors)
        Dim professorId As Integer = 0
        If professorBox2.SelectedIndex >= 0 Then
            professorId = Convert.ToInt32(professorBox2.SelectedValue)
        End If

        ' Pass professorId to the function
        Dim dt As DataTable = GetFormattedReports(fromDate, toDate, professorId)

        reportView.Rows.Clear()

        For Each row As DataRow In dt.Rows
            reportView.Rows.Add(
            If(row("student_number") Is DBNull.Value, "", row("student_number")),
            If(row("student_name") Is DBNull.Value, "", row("student_name")),
            If(row("section") Is DBNull.Value, "", row("section")),
            If(row("professor_name") Is DBNull.Value, "", row("professor_name")),
            If(row("reason") Is DBNull.Value, "", row("reason")),
            If(row("message") Is DBNull.Value, "", row("message")),
            If(row("consultation_date") Is DBNull.Value, "", row("consultation_date")),
            If(row("time_in") Is DBNull.Value, "", row("time_in")),
            If(row("time_out") Is DBNull.Value, "", row("time_out")),
            If(row("id") Is DBNull.Value, 0, row("id"))
        )
        Next
    End Sub




    Private Sub reportStudentNumberInput_TextChanged(sender As Object, e As EventArgs) Handles reportStudentNumberInput.TextChanged
        Dim txtBox As TextBox = CType(sender, TextBox)
        Dim cursorPosition As Integer = txtBox.SelectionStart

        ' Keep only numeric characters (remove letters, symbols, etc.)
        Dim digitsOnly As String = New String(txtBox.Text.Where(AddressOf Char.IsDigit).ToArray())

        ' Format: insert dash after 2 digits (e.g., 20-1234)
        If digitsOnly.Length > 2 Then
            txtBox.Text = digitsOnly.Substring(0, 2) & "-" & digitsOnly.Substring(2)
        Else
            txtBox.Text = digitsOnly
        End If

        ' Reset cursor position
        If cursorPosition <= 2 Then
            txtBox.SelectionStart = cursorPosition
        Else
            txtBox.SelectionStart = Math.Min(txtBox.Text.Length, cursorPosition + 1)
        End If
    End Sub

    Private Sub reportProfessorBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles reportProfessorBox.SelectedIndexChanged

    End Sub

    Private Sub LoadProfessorsToComboBox2()
        Dim dt As DataTable = GetAllProfessors()

        ' Add computed column for formatted full name (e.g., Lastname, Firstname M. Suffix)
        dt.Columns.Add("FullName", GetType(String))
        For Each row As DataRow In dt.Rows
            Dim lastName As String = row("last_name").ToString()
            Dim firstName As String = row("first_name").ToString()
            Dim mi As String = If(String.IsNullOrWhiteSpace(row("middle_initial").ToString()), "", " " & row("middle_initial") & ".")
            Dim suffix As String = If(String.IsNullOrWhiteSpace(row("suffix").ToString()), "", " " & row("suffix").ToString())

            row("FullName") = $"{lastName}, {firstName}{mi}{suffix}".Trim()
        Next

        reportProfessorBox.DataSource = dt
        reportProfessorBox.DisplayMember = "FullName" ' what user sees
        reportProfessorBox.ValueMember = "id"         ' internal value (used for FK)
        reportProfessorBox.SelectedIndex = -1         ' no selection initially
    End Sub
    Private Sub LoadProfessorsToComboBox3()
        Dim dt As DataTable = GetAllProfessors()

        ' Add computed column for formatted full name (Lastname, Firstname M. Suffix)
        dt.Columns.Add("FullName", GetType(String))
        For Each row As DataRow In dt.Rows
            Dim lastName As String = row("last_name").ToString()
            Dim firstName As String = row("first_name").ToString()
            Dim mi As String = If(String.IsNullOrWhiteSpace(row("middle_initial").ToString()), "", " " & row("middle_initial") & ".")
            Dim suffix As String = If(String.IsNullOrWhiteSpace(row("suffix").ToString()), "", " " & row("suffix").ToString())
            row("FullName") = $"{lastName}, {firstName}{mi}{suffix}".Trim()
        Next

        ' Add extra "Special Reasons" option
        Dim specialRow As DataRow = dt.NewRow()
        specialRow("id") = 99
        specialRow("FullName") = "Special Reasons"
        dt.Rows.Add(specialRow)

        ' Set DisplayMember and ValueMember BEFORE DataSource
        professorBox2.DisplayMember = "FullName"
        professorBox2.ValueMember = "id"
        professorBox2.DataSource = dt
        professorBox2.SelectedIndex = -1 ' no selection initially
    End Sub

    Private Sub LoadReasonsToComboBox2()
        Dim dt As DataTable = GetAllReasons() ' From your Reason module

        reportReasonBox.DataSource = dt
        reportReasonBox.DisplayMember = "reason"    ' what the user sees
        reportReasonBox.ValueMember = "id"          ' internal hidden FK id



        reportReasonBox.SelectedIndex = -1 ' ← No item selected
    End Sub
    Private Sub reportReasonBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles reportReasonBox.SelectedIndexChanged
        Dim selectedRow As DataRowView = TryCast(reportReasonBox.SelectedItem, DataRowView)

        If selectedRow IsNot Nothing Then
            Dim isSpecial As Boolean = Convert.ToBoolean(selectedRow("is_special"))
            reportProfessorBox.Enabled = Not isSpecial ' Disable if special is true
        End If
    End Sub

    Private Sub FormatAndValidate24HourTime(ByRef txtBox As TextBox)
        If String.IsNullOrWhiteSpace(txtBox.Text) Then Exit Sub

        Dim originalText As String = txtBox.Text
        Dim originalCursor As Integer = txtBox.SelectionStart

        ' Extract digits only
        Dim digits As String = New String(originalText.Where(Function(c) Char.IsDigit(c)).ToArray())

        If digits.Length > 4 Then digits = digits.Substring(0, 4)

        ' Apply formatting (HH:mm)
        Dim formatted As String = digits
        If digits.Length >= 3 Then
            formatted = digits.Substring(0, 2) & ":" & digits.Substring(2)
        End If

        ' Adjust the cursor position
        Dim newCursor As Integer = originalCursor

        ' If the user just typed the 3rd digit (i.e., adding ':'), move cursor right
        If digits.Length = 3 AndAlso Not originalText.Contains(":") AndAlso originalCursor >= 2 Then
            newCursor += 1
        End If

        txtBox.Text = formatted
        txtBox.SelectionStart = Math.Min(newCursor, txtBox.Text.Length)

        ' Validate only when format is fully formed (e.g., "23:45")
        If txtBox.Text.Length = 5 Then
            Dim parsedTime As DateTime
            If Not DateTime.TryParseExact(txtBox.Text, "HH:mm", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, parsedTime) Then
                MessageBox.Show("Invalid time. Please enter a valid 24-hour time (e.g., 23:45).", "Time Format Error")
                txtBox.Clear()
            End If
        End If
    End Sub


    Private Sub reportTimeInInput_TextChanged(sender As Object, e As EventArgs) Handles reportTimeInInput.TextChanged
        FormatAndValidate24HourTime(reportTimeInInput)
    End Sub

    Private Sub reportTimeOutInput_TextChanged(sender As Object, e As EventArgs) Handles reportTimeOutInput.TextChanged
        FormatAndValidate24HourTime(reportTimeOutInput)
    End Sub

    Private Sub formAddBtn_Click(sender As Object, e As EventArgs) Handles formAddBtn.Click
        ' Validate student number
        If String.IsNullOrWhiteSpace(reportStudentNumberInput.Text) Then
            MessageBox.Show("Please enter a student number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim studentNumber As String = reportStudentNumberInput.Text.Trim()
        Dim student As StudentModel = GetStudentByNumber(studentNumber)

        If student Is Nothing Then
            MessageBox.Show("Student number not found in the database.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim studentId As Integer = student.Id


        ' Validate reason
        If reportReasonBox.SelectedIndex = -1 OrElse reportReasonBox.SelectedValue Is Nothing Then
            MessageBox.Show("Please select a reason.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim selectedReasonId As Object = reportReasonBox.SelectedValue

        ' Check if reason is special
        Dim isSpecial As Boolean = Convert.ToBoolean(CType(reportReasonBox.SelectedItem, DataRowView)("is_special"))

        ' Validate professor only if not special
        Dim selectedProfessorId As Object = Nothing
        If Not isSpecial Then
            If reportProfessorBox.SelectedIndex = -1 OrElse reportProfessorBox.SelectedValue Is Nothing Then
                MessageBox.Show("Please select a professor.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            selectedProfessorId = reportProfessorBox.SelectedValue
        End If

        ' Validate date
        If reportDateInput.Value = Nothing Then
            MessageBox.Show("Please select a date.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Dim consultationDate As Date = reportDateInput.Value.Date

        ' Validate and parse time in
        Dim timeIn As TimeSpan
        If Not TimeSpan.TryParse(reportTimeInInput.Text.Trim(), timeIn) Then
            MessageBox.Show("Please enter a valid Time In (e.g., 14:00).", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Validate and parse time out
        Dim timeOut As TimeSpan
        If Not TimeSpan.TryParse(reportTimeOutInput.Text.Trim(), timeOut) Then
            MessageBox.Show("Please enter a valid Time Out (e.g., 15:30).", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Get optional message
        Dim message As String = reportMessageInput.Text.Trim()

        ' Confirm submission
        Dim confirm = MessageBox.Show("Are you sure you want to add this report?", "Confirm Add", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirm <> DialogResult.Yes Then Exit Sub

        ' Perform insert
        InsertReport(studentId, selectedReasonId, message, consultationDate, timeIn, timeOut, selectedProfessorId)

        MessageBox.Show("Report successfully added.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        LoadReports()
        ClearInputs(adminFormPanel) ' ← replace with your input container
    End Sub


    Private selectedReportId As Integer = -1

    Private Sub reportView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles reportView.CellContentClick
        If e.RowIndex < 0 Then Exit Sub

        Dim row As DataGridViewRow = reportView.Rows(e.RowIndex)

        ' Get hidden Report ID
        selectedReportId = Convert.ToInt32(row.Cells("reportIdCol").Value)

        ' Student Number
        reportStudentNumberInput.Text = row.Cells("studentNumberCol").Value.ToString()

        ' Reason - match by display text
        Dim reasonText As String = row.Cells("reasonCol").Value.ToString().Trim()
        For i As Integer = 0 To reportReasonBox.Items.Count - 1
            If reportReasonBox.GetItemText(reportReasonBox.Items(i)).Trim() = reasonText Then
                reportReasonBox.SelectedIndex = i
                Exit For
            End If
        Next

        ' Professor - match by display text
        Dim professorText As String = row.Cells("professorNameCol").Value.ToString().Trim()
        If String.IsNullOrEmpty(professorText) Then
            reportProfessorBox.SelectedIndex = -1 ' clear selection
        Else
            For i As Integer = 0 To reportProfessorBox.Items.Count - 1
                If reportProfessorBox.GetItemText(reportProfessorBox.Items(i)).Trim() = professorText Then
                    reportProfessorBox.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        ' Message
        reportMessageInput.Text = row.Cells("messageCol").Value.ToString()

        ' Date
        Dim consultationDate As Date
        If Date.TryParse(row.Cells("dateCol").Value.ToString(), consultationDate) Then
            reportDateInput.Value = consultationDate
        End If

        ' Time In and Time Out
        reportTimeInInput.Text = row.Cells("timeInCol").Value.ToString()
        reportTimeOutInput.Text = row.Cells("timeOutCol").Value.ToString()
    End Sub


    Private Sub formUpdateBtn_Click(sender As Object, e As EventArgs) Handles formUpdateBtn.Click
        ' Validate reason selection
        Dim selectedRow As DataRowView = TryCast(reportReasonBox.SelectedItem, DataRowView)
        If selectedRow Is Nothing Then
            MessageBox.Show("Please select a valid reason.", "Validation Error")
            Exit Sub
        End If

        ' Get reason and special flag
        Dim reasonId As Integer = Convert.ToInt32(selectedRow("id"))
        Dim isSpecial As Boolean = Convert.ToBoolean(selectedRow("is_special"))

        ' Get other inputs
        Dim message As String = reportMessageInput.Text.Trim()
        Dim consultationDate As Date = reportDateInput.Value.Date
        Dim studentNumber As String = reportStudentNumberInput.Text.Trim()

        ' Validate student number
        If String.IsNullOrWhiteSpace(studentNumber) Then
            MessageBox.Show("Student number cannot be empty.", "Validation Error")
            Exit Sub
        End If

        If Not DoesStudentExist(studentNumber) Then
            MessageBox.Show("The entered student number does not exist in the database.", "Invalid Student")
            Exit Sub
        End If

        ' Parse TimeIn/TimeOut
        Dim timeIn As TimeSpan
        Dim timeOut As TimeSpan
        If Not TimeSpan.TryParse(reportTimeInInput.Text, timeIn) OrElse Not TimeSpan.TryParse(reportTimeOutInput.Text, timeOut) Then
            MessageBox.Show("Please enter valid time values.", "Validation Error")
            Exit Sub
        End If

        ' Decide professorId
        Dim professorId As Integer? = Nothing
        If Not isSpecial Then
            If reportProfessorBox.SelectedIndex >= 0 AndAlso IsNumeric(reportProfessorBox.SelectedValue) Then
                professorId = Convert.ToInt32(reportProfessorBox.SelectedValue)
            End If
        Else
            reportProfessorBox.SelectedIndex = -1
            reportProfessorBox.Text = ""
        End If

        ' Update the report
        UpdateReport(selectedReportId,
                 reasonId,
                 message,
                 consultationDate,
                 timeIn,
                 timeOut,
                 professorId,
                 studentNumber)

        ' Refresh and confirm
        LoadReports()
        MessageBox.Show("Report updated successfully.", "Success")
    End Sub


    Private Sub formDeleteBtn_Click(sender As Object, e As EventArgs) Handles formDeleteBtn.Click
        If selectedReportId <= 0 Then
            MessageBox.Show("Please select a report to delete.", "No Report Selected")
            Exit Sub
        End If

        Dim confirmResult As DialogResult = MessageBox.Show(
            "Are you sure you want to delete this report?",
            "Confirm Delete",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        )

        If confirmResult = DialogResult.Yes Then
            Try
                DeleteReport(selectedReportId)
                LoadReports() ' Refresh the table after deletion
                MessageBox.Show("Report deleted successfully.", "Success")
                ' Optionally clear form fields here
            Catch ex As Exception
                MessageBox.Show("An error occurred while deleting the report: " & ex.Message, "Error")
            End Try
        End If
    End Sub

    Private Sub reportFromDate_ValueChanged(sender As Object, e As EventArgs) Handles reportFromDate.ValueChanged
        ' Refresh the report view when the "from" date is changed
        LoadReports()
    End Sub

    Private Sub reportToDate_ValueChanged(sender As Object, e As EventArgs) Handles reportToDate.ValueChanged
        ' Refresh the report view when the "to" date is changed
        LoadReports()
    End Sub

    Private Sub LoadSectionsIntoComboBox2()
        Dim dt As DataTable = GetAllSections()

        ' Optional: add a display column if you want custom format
        dt.Columns.Add("DisplayText", GetType(String))
        For Each row As DataRow In dt.Rows
            row("DisplayText") = row("section").ToString()
        Next

        studentSectionSearch.DataSource = dt
        studentSectionSearch.DisplayMember = "DisplayText"
        studentSectionSearch.ValueMember = "id" ' this keeps the hidden ID

        studentSectionSearch.SelectedIndex = -1 ' ← No item selected
    End Sub

    Private Sub LoadSectionsIntoComboBox3()
        Dim dt As DataTable = GetAllSections()

        ' Add DisplayText column to match section name in studentView
        dt.Columns.Add("DisplayText", GetType(String))
        For Each row As DataRow In dt.Rows
            row("DisplayText") = row("section").ToString()
        Next

        sectionBox.DataSource = dt
        sectionBox.DisplayMember = "DisplayText"
        sectionBox.ValueMember = "id"
        sectionBox.SelectedIndex = -1
    End Sub


    Private Sub LoadFilteredStudents()
        Dim lastNameFilter As String = studentSurnameSearch.Text.Trim()
        Dim selectedSectionId As Integer? = Nothing

        ' Make sure the SelectedValue is an Integer (not a DataRowView)
        If studentSectionSearch.SelectedIndex <> -1 AndAlso TypeOf studentSectionSearch.SelectedValue Is Integer Then
            selectedSectionId = CInt(studentSectionSearch.SelectedValue)
        End If

        Dim studentList = SearchStudentsByLastNameAndSection(lastNameFilter, selectedSectionId)

        studentView.Rows.Clear()
        For Each stu In studentList
            studentView.Rows.Add(
            stu.StudentNumber,
            stu.FirstName,
            stu.LastName,
            stu.MiddleInitial,
            stu.Suffix,
            stu.Section,
            stu.Email,
            stu.Status,
            stu.Id ' This should match the last column (studentId)
        )
        Next
    End Sub

    Private Sub studentSurnameSearch_TextChanged(sender As Object, e As EventArgs) Handles studentSurnameSearch.TextChanged
        LoadFilteredStudents()
    End Sub

    Private Sub studentSectionSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles studentSectionSearch.SelectedIndexChanged
        LoadFilteredStudents()
    End Sub

    Private selectedStudentId As Integer = -1
    Dim studentIdList As New List(Of Integer)
    Private Sub studentView_SelectionChanged(sender As Object, e As EventArgs) Handles studentView.SelectionChanged
        ' Update the global list
        studentIdList = studentView.SelectedRows _
        .Cast(Of DataGridViewRow)() _
        .Where(Function(r) Not r.IsNewRow) _
        .Select(Function(r) Convert.ToInt32(r.Cells("studentId").Value)) _
        .ToList()

        ' Optional: display IDs for debugging
        Console.WriteLine("Selected IDs: " & String.Join(", ", studentIdList))
    End Sub


    Private Sub studentView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles studentView.CellContentClick
        If e.RowIndex >= 0 AndAlso e.RowIndex < studentView.Rows.Count Then
            Dim row As DataGridViewRow = studentView.Rows(e.RowIndex)

            ' Capture ID
            selectedStudentId = Convert.ToInt32(row.Cells("studentId").Value)
            originalStudentNumber = studentView.CurrentRow.Cells("studentNumber").Value.ToString()

            ' Safe text assignments
            studentNumberInput.Text = If(IsDBNull(row.Cells("studentNumber").Value), "", row.Cells("studentNumber").Value.ToString())
            firstNameInput.Text = If(IsDBNull(row.Cells("firstName").Value), "", row.Cells("firstName").Value.ToString())
            lastNameInput.Text = If(IsDBNull(row.Cells("lastName").Value), "", row.Cells("lastName").Value.ToString())
            middleInitialInput.Text = If(IsDBNull(row.Cells("middleInitial").Value), "", row.Cells("middleInitial").Value.ToString())
            emailInput.Text = If(IsDBNull(row.Cells("email").Value), "", row.Cells("email").Value.ToString())

            ' Suffix
            Dim suffixValue As String = If(IsDBNull(row.Cells("suffix").Value), "", row.Cells("suffix").Value.ToString().Trim())
            If String.IsNullOrEmpty(suffixValue) OrElse Not suffixBox.Items.Contains(suffixValue) Then
                suffixBox.SelectedIndex = -1
            Else
                suffixBox.SelectedItem = suffixValue
            End If

            ' Section
            Dim sectionName As String = If(IsDBNull(row.Cells("section").Value), "", row.Cells("section").Value.ToString().Trim())
            Dim matchedSectionIndex As Integer = -1
            For i As Integer = 0 To sectionBox.Items.Count - 1
                If sectionBox.GetItemText(sectionBox.Items(i)).Trim() = sectionName Then
                    matchedSectionIndex = i
                    Exit For
                End If
            Next
            sectionBox.SelectedIndex = matchedSectionIndex

            ' Status
            Dim statusValue As String = If(IsDBNull(row.Cells("status").Value), "", row.Cells("status").Value.ToString())
            If String.IsNullOrEmpty(statusValue) OrElse Not statusBox.Items.Contains(statusValue) Then
                statusBox.SelectedIndex = -1
            Else
                statusBox.SelectedItem = statusValue
            End If
        End If
    End Sub
    Private Sub archiveBtn_Click(sender As Object, e As EventArgs) Handles archiveBtn.Click
        ' Check if single student is selected or multiple students are selected
        Dim studentIds As New List(Of Integer)()

        ' Check for multiple selection in DataGridView
        If studentView.SelectedRows.Count > 1 Then
            ' Multiple students selected
            For Each row As DataGridViewRow In studentView.SelectedRows
                If row.Cells("studentId").Value IsNot Nothing Then
                    studentIds.Add(CInt(row.Cells("studentId").Value))
                End If
            Next
        ElseIf selectedStudentId > 0 Then
            ' Single student selected (legacy method)
            studentIds.Add(selectedStudentId)
        Else
            MessageBox.Show("Please select at least one student to archive.", "No Selection")
            Exit Sub
        End If

        ' Confirm action
        Dim message As String = If(studentIds.Count > 1,
                              $"Are you sure you want to archive {studentIds.Count} students?",
                              "Are you sure you want to archive this student?")

        If MessageBox.Show(message, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Exit Sub
        End If

        ' Get status (for single student, use the form inputs; for multiple, use a default or prompt)
        ' For multiple students, we won't change the status, just set isGraduate = 1
        Dim status As String = Nothing ' Use Nothing to indicate we don't want to change status
        If studentIds.Count = 1 Then
            ' For single student, use the form input status
            status = statusBox.Text.Trim()
        End If
        ' Archive the students
        Dim success As Boolean = MarkStudentsAsGraduated(studentIds, status)

        If success Then
            MessageBox.Show($"{studentIds.Count} student(s) archived successfully!", "Success")
            ClearInputs(adminStudentPanel)
            LoadFilteredStudents()          ' Refresh the grid
            selectedStudentId = -1          ' Reset selected ID
        Else
            MessageBox.Show("Failed to archive student(s).", "Error")
        End If
    End Sub



    Private Sub studentNumberInput_TextChanged(sender As Object, e As EventArgs) Handles studentNumberInput.TextChanged
        Dim txtBox As TextBox = CType(sender, TextBox)
        Dim cursorPosition As Integer = txtBox.SelectionStart

        ' Keep only numeric characters (remove letters, symbols, etc.)
        Dim digitsOnly As String = New String(txtBox.Text.Where(AddressOf Char.IsDigit).ToArray())

        ' Format: insert dash after 2 digits (e.g., 20-1234)
        If digitsOnly.Length > 2 Then
            txtBox.Text = digitsOnly.Substring(0, 2) & "-" & digitsOnly.Substring(2)
        Else
            txtBox.Text = digitsOnly
        End If

        ' Reset cursor position
        If cursorPosition <= 2 Then
            txtBox.SelectionStart = cursorPosition
        Else
            txtBox.SelectionStart = Math.Min(txtBox.Text.Length, cursorPosition + 1)
        End If
    End Sub

    Private Sub studentAddBtn_Click(sender As Object, e As EventArgs) Handles studentAddBtn.Click
        ' Get input values
        Dim studentNumber As String = studentNumberInput.Text.Trim()
        Dim firstName As String = firstNameInput.Text.Trim()
        Dim lastName As String = lastNameInput.Text.Trim()
        Dim middleInitial As String = middleInitialInput.Text.Trim()
        Dim suffix As String = suffixBox.Text.Trim()
        Dim email As String = emailInput.Text.Trim()
        Dim status As String = statusBox.Text.Trim()

        ' Validate section
        If sectionBox.SelectedIndex = -1 Then
            MessageBox.Show("Please select a section.", "Validation Error")
            Exit Sub
        End If

        Dim sectionId As Integer = Convert.ToInt32(sectionBox.SelectedValue)

        ' Validate input
        If Not System.Text.RegularExpressions.Regex.IsMatch(studentNumber, "^\d{2}-\d{5}$") OrElse
       String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) Then

            MessageBox.Show("Please enter valid and complete information.", "Validation Error")
            Exit Sub
        End If

        ' Insert student
        Dim success As Boolean = InsertStudent(studentNumber, firstName, lastName, middleInitial, suffix, email, sectionId, status)

        If success Then
            MessageBox.Show("Student successfully added!", "Success")
            ClearInputs(adminStudentPanel)
            LoadFilteredStudents()

            ' ✅ QR Content Format: [StudentNumber]
            Dim qrContent As String = $"[{studentNumber}]"

            Try
                Cursor.Current = Cursors.WaitCursor

                ' ✅ Generate QR code
                Dim qrGenerator As New QRCodeGenerator()
                Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCodeGenerator.ECCLevel.H)
                Dim qrCode = New QRCode(qrData)
                Dim qrImage = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

                ' Use embedded resource (no need for path or File.Exists)
                Dim logo As Image = My.Resources.PLP

                ' Resize logo
                Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
                Dim resizedLogo As New Bitmap(logoSize, logoSize)
                Using g As Graphics = Graphics.FromImage(resizedLogo)
                    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(logo, 0, 0, logoSize, logoSize)
                End Using


                ' ✅ Combine QR and logo
                Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
                Using g As Graphics = Graphics.FromImage(qrWithLogo)
                    g.Clear(Color.White)
                    g.DrawImage(qrImage, 0, 0)

                    ' White box behind logo
                    Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                    Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                    Dim padding As Integer = 6
                    Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                    g.FillRectangle(Brushes.White, whiteRect)

                    g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
                End Using

                ' ✅ Save QR to stream
                Dim ms As New MemoryStream()
                qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
                ms.Position = 0

                ' ✅ Send Email
                Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
                Dim toAddress As New MailAddress(email)
                Dim fromPassword As String = "swlqbwgztcqbneuw"

                Dim message As New MailMessage(fromAddress, toAddress)
                message.Subject = "Your Student QR Code"
                message.Body = $"Hi {firstName}," & vbCrLf & vbCrLf &
                           "Attached is your student QR code." & vbCrLf &
                           "You may use it for login or authentication." & vbCrLf

                message.Attachments.Add(New Attachment(ms, "StudentQR.png", "image/png"))

                Dim smtp As New SmtpClient("smtp.gmail.com", 587)
                smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
                smtp.EnableSsl = True
                smtp.Send(message)

                MessageBox.Show("QR Code sent to student's email.", "QR Email Sent")

            Catch ex As Exception
                MessageBox.Show("Failed to send QR email: " & ex.Message, "QR Email Error")
            Finally
                Cursor.Current = Cursors.Default
            End Try
        Else
            MessageBox.Show("Failed to add student.", "Error")
        End If
    End Sub

    Private originalStudentNumber As String = ""

    Private Sub studentUpdateBtn_Click(sender As Object, e As EventArgs) Handles studentUpdateBtn.Click
        ' Get input values
        Dim studentNumber As String = studentNumberInput.Text.Trim()
        Dim firstName As String = firstNameInput.Text.Trim()
        Dim lastName As String = lastNameInput.Text.Trim()
        Dim middleInitial As String = middleInitialInput.Text.Trim()
        Dim suffix As String = suffixBox.Text.Trim()
        Dim email As String = emailInput.Text.Trim()
        Dim status As String = statusBox.Text.Trim()

        If sectionBox.SelectedIndex = -1 Then
            MessageBox.Show("Please select a section.", "Validation Error")
            Exit Sub
        End If

        Dim sectionId As Integer = Convert.ToInt32(sectionBox.SelectedValue)

        ' Validate inputs
        If Not System.Text.RegularExpressions.Regex.IsMatch(studentNumber, "^\d{2}-\d{5}$") OrElse
       String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) Then

            MessageBox.Show("Please enter valid and complete information.", "Validation Error")
            Exit Sub
        End If

        ' Compare with original student number
        Dim studentNumberChanged As Boolean = (studentNumber <> originalStudentNumber)

        ' Update student using original student number
        Dim success As Boolean = UpdateStudent(originalStudentNumber, studentNumber, firstName, lastName, middleInitial, suffix, email, sectionId, status)

        If success Then
            MessageBox.Show("Student successfully updated!", "Success")
            ClearInputs(adminStudentPanel)
            LoadFilteredStudents()

            If studentNumberChanged Then
                Try
                    Cursor.Current = Cursors.WaitCursor

                    ' ✅ QR Code Content
                    Dim qrContent As String = $"[{studentNumber}]"
                    Dim qrGenerator As New QRCoder.QRCodeGenerator()
                    Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
                    Dim qrCode = New QRCoder.QRCode(qrData)
                    Dim qrImage = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

                    ' ✅ Load and resize logo
                    ' Use embedded resource (no need for path or File.Exists)
                    Dim logo As Image = My.Resources.PLP

                    ' Resize logo
                    Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
                    Dim resizedLogo As New Bitmap(logoSize, logoSize)
                    Using g As Graphics = Graphics.FromImage(resizedLogo)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        g.DrawImage(logo, 0, 0, logoSize, logoSize)
                    End Using


                    ' ✅ Combine QR and logo
                    Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
                    Using g As Graphics = Graphics.FromImage(qrWithLogo)
                        g.Clear(Color.White)
                        g.DrawImage(qrImage, 0, 0)

                        Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                        Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                        Dim padding As Integer = 6
                        Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                        g.FillRectangle(Brushes.White, whiteRect)

                        g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
                    End Using

                    ' ✅ Save QR to stream
                    Dim ms As New MemoryStream()
                    qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
                    ms.Position = 0

                    ' ✅ Email
                    Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
                    Dim toAddress As New MailAddress(email)
                    Dim fromPassword As String = "swlqbwgztcqbneuw"

                    Dim message As New MailMessage(fromAddress, toAddress)
                    message.Subject = "Updated Student QR Code"
                    message.Body = $"Hi {firstName}," & vbCrLf & vbCrLf &
                               "Your updated student QR code is attached." & vbCrLf &
                               "Use this QR to log in or verify your identity."

                    message.Attachments.Add(New Attachment(ms, "UpdatedStudentQR.png", "image/png"))

                    Dim smtp As New SmtpClient("smtp.gmail.com", 587)
                    smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
                    smtp.EnableSsl = True
                    smtp.Send(message)

                    MessageBox.Show("Updated QR code sent to student's email.", "Email Sent")

                Catch ex As Exception
                    MessageBox.Show("Failed to send QR email: " & ex.Message, "Email Error")
                Finally
                    Cursor.Current = Cursors.Default
                End Try
            End If

        Else
            MessageBox.Show("Failed to update student.", "Error")
        End If
    End Sub


    Private Sub studentDeleteBtn_Click(sender As Object, e As EventArgs) Handles studentDeleteBtn.Click
        ' Check if a student is selected
        If selectedStudentId = -1 Then
            MessageBox.Show("Please select a student to delete.", "Validation Error")
            Exit Sub
        End If

        ' Confirm deletion
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this student?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = DialogResult.Yes Then
            Dim success As Boolean = DeleteStudent(selectedStudentId)

            If success Then
                MessageBox.Show("Student deleted successfully.", "Success")
                ClearInputs(adminStudentPanel) ' Clear input fields
                selectedStudentId = -1          ' Reset selected ID
                LoadFilteredStudents()
            Else
                MessageBox.Show("Failed to delete student.", "Error")
            End If
        End If
    End Sub
    Private Sub professorSurnameSearch_TextChanged(sender As Object, e As EventArgs) Handles professorSurnameSearch.TextChanged
        LoadProfessorsToGrid()
    End Sub

    Private Sub LoadProfessorsToGrid()
        Dim surnameFilter As String = professorSurnameSearch.Text.Trim()
        Dim dt As DataTable = GetFilteredProfessors(surnameFilter)

        professorView.Rows.Clear()

        For Each row As DataRow In dt.Rows
            professorView.Rows.Add(
                row("first_name").ToString(),
                row("last_name").ToString(),
                row("middle_initial").ToString(),
                row("suffix").ToString(),
                row("email").ToString(),
                row("username").ToString(),
                row("password").ToString(),
                row("id").ToString() ' professorIdCol
            )
        Next
    End Sub

    Private selectedProfessorId As Integer
    Private Sub professorView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles professorView.CellContentClick
        If e.RowIndex >= 0 AndAlso e.RowIndex < professorView.Rows.Count Then
            Dim row As DataGridViewRow = professorView.Rows(e.RowIndex)

            ' Assign values from each column safely
            professorFirstName.Text = If(IsDBNull(row.Cells("firstNameCol").Value), "", row.Cells("firstNameCol").Value.ToString())
            professorLastName.Text = If(IsDBNull(row.Cells("lastNameCol").Value), "", row.Cells("lastNameCol").Value.ToString())
            professorMiddleInitial.Text = If(IsDBNull(row.Cells("middleInitialCol").Value), "", row.Cells("middleInitialCol").Value.ToString())

            ' Handle suffix safely
            Dim suffixValue As String = If(IsDBNull(row.Cells("suffixCol").Value), "", row.Cells("suffixCol").Value.ToString().Trim())
            If String.IsNullOrEmpty(suffixValue) OrElse Not professorSuffix.Items.Contains(suffixValue) Then
                professorSuffix.SelectedIndex = -1
            Else
                professorSuffix.SelectedItem = suffixValue
            End If

            professorEmail.Text = If(IsDBNull(row.Cells("emailCol").Value), "", row.Cells("emailCol").Value.ToString())
            professorUsername.Text = If(IsDBNull(row.Cells("usernameCol").Value), "", row.Cells("usernameCol").Value.ToString())
            professorPassword.Text = If(IsDBNull(row.Cells("passwordCol").Value), "", row.Cells("passwordCol").Value.ToString())

            ' Store the professor ID
            selectedProfessorId = Convert.ToInt32(row.Cells("professorIdCol").Value)
        End If
    End Sub


    Private Sub professorAddBtn_Click(sender As Object, e As EventArgs) Handles professorAddBtn.Click
        ' Get input values
        Dim lastName As String = professorLastName.Text.Trim()
        Dim firstName As String = professorFirstName.Text.Trim()
        Dim middleInitial As String = professorMiddleInitial.Text.Trim()
        Dim suffix As String = If(professorSuffix.SelectedItem IsNot Nothing, professorSuffix.SelectedItem.ToString().Trim(), "")
        Dim email As String = professorEmail.Text.Trim()
        Dim username As String = professorUsername.Text.Trim()
        Dim password As String = professorPassword.Text.Trim()

        ' Validate input
        If String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       String.IsNullOrWhiteSpace(username) OrElse
       String.IsNullOrWhiteSpace(password) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) Then

            MessageBox.Show("Please enter valid and complete professor information.", "Validation Error")
            Exit Sub
        End If

        ' Insert professor
        Dim success As Boolean = InsertProfessor(lastName, firstName, middleInitial, suffix, email, username, password)

        If success Then
            MessageBox.Show("Professor successfully added!", "Success")
            ClearInputs(adminProfessorPanel)
            LoadProfessorsToGrid()

            Try
                Cursor.Current = Cursors.WaitCursor

                ' QR code content
                Dim qrContent As String = $"[{username}][{password}]"
                Dim qrGenerator As New QRCoder.QRCodeGenerator()
                Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
                Dim qrCode = New QRCoder.QRCode(qrData)
                Dim qrImage As Bitmap = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

                ' Logo
                ' Use embedded resource (no need for path or File.Exists)
                Dim logo As Image = My.Resources.PLP

                ' Resize logo
                Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
                Dim resizedLogo As New Bitmap(logoSize, logoSize)
                Using g As Graphics = Graphics.FromImage(resizedLogo)
                    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(logo, 0, 0, logoSize, logoSize)
                End Using


                ' Merge QR and logo
                Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
                Using g As Graphics = Graphics.FromImage(qrWithLogo)
                    g.Clear(Color.White)
                    g.DrawImage(qrImage, 0, 0)

                    Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                    Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                    Dim padding As Integer = 6
                    Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                    g.FillRectangle(Brushes.White, whiteRect)
                    g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
                End Using

                ' Convert image to stream
                Dim ms As New MemoryStream()
                qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
                ms.Position = 0

                ' Send email
                Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
                Dim toAddress As New MailAddress(email)
                Dim fromPassword As String = "swlqbwgztcqbneuw"

                Dim message As New MailMessage(fromAddress, toAddress)
                message.Subject = "Your Professor QR Code"
                message.Body = $"Hi {firstName}," & vbCrLf & vbCrLf &
                           "Attached is your QR code containing your login credentials." & vbCrLf &
                           "Please scan it for convenient login." & vbCrLf & vbCrLf &
                           "Thank you."

                message.Attachments.Add(New Attachment(ms, "ProfessorQRCode.png", "image/png"))

                Dim smtp As New SmtpClient("smtp.gmail.com", 587)
                smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
                smtp.EnableSsl = True
                smtp.Send(message)

                MessageBox.Show("QR code sent to professor's email.", "Email Sent")
            Catch ex As Exception
                MessageBox.Show("Failed to send QR email: " & ex.Message, "Email Error")
            Finally
                Cursor.Current = Cursors.Default
            End Try
        Else
            MessageBox.Show("Failed to add professor.", "Error")
        End If
    End Sub

    Private Sub professorUpdateBtn_Click(sender As Object, e As EventArgs) Handles professorUpdateBtn.Click
        ' ✅ Validate selection
        If selectedProfessorId = -1 Then
            MessageBox.Show("No professor selected for update.", "Validation Error")
            Exit Sub
        End If

        ' ✅ Retrieve inputs
        Dim firstName As String = professorFirstName.Text.Trim()
        Dim lastName As String = professorLastName.Text.Trim()
        Dim middleInitial As String = professorMiddleInitial.Text.Trim()
        Dim suffix As String = If(professorSuffix.SelectedItem IsNot Nothing, professorSuffix.SelectedItem.ToString(), "")
        Dim email As String = professorEmail.Text.Trim()
        Dim newUsername As String = professorUsername.Text.Trim()
        Dim newPassword As String = professorPassword.Text.Trim()

        ' ✅ Validate inputs
        If String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       String.IsNullOrWhiteSpace(newUsername) OrElse
       String.IsNullOrWhiteSpace(newPassword) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) Then

            MessageBox.Show("Please enter valid and complete professor information.", "Validation Error")
            Exit Sub
        End If

        ' ✅ Fetch old credentials
        Dim oldData As DataRow = GetProfessorById(selectedProfessorId)
        If oldData Is Nothing Then
            MessageBox.Show("Professor not found in database.", "Error")
            Exit Sub
        End If

        Dim oldUsername As String = oldData("username").ToString().Trim()
        Dim oldPassword As String = oldData("password").ToString().Trim()

        Dim credentialsChanged As Boolean =
        (newUsername.Trim().ToLower() <> oldUsername.ToLower()) OrElse
        (newPassword.Trim() <> oldPassword)

        ' ✅ Update record
        Dim success As Boolean = UpdateProfessor(selectedProfessorId, lastName, firstName, middleInitial, suffix, email, newUsername, newPassword)

        If success Then
            MessageBox.Show("Professor successfully updated.", "Success")
            ClearInputs(adminProfessorPanel)
            LoadProfessorsToGrid()
            selectedProfessorId = -1

            If credentialsChanged Then
                Try
                    Cursor.Current = Cursors.WaitCursor

                    ' QR generation
                    Dim qrContent As String = $"[{newUsername}][{newPassword}]"
                    Dim qrGenerator As New QRCoder.QRCodeGenerator()
                    Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
                    Dim qrCode = New QRCoder.QRCode(qrData)
                    Dim qrImage = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

                    ' Logo setup
                    ' Use embedded resource (no need for path or File.Exists)
                    Dim logo As Image = My.Resources.PLP

                    ' Resize logo
                    Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
                    Dim resizedLogo As New Bitmap(logoSize, logoSize)
                    Using g As Graphics = Graphics.FromImage(resizedLogo)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        g.DrawImage(logo, 0, 0, logoSize, logoSize)
                    End Using


                    ' Merge logo and QR
                    Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
                    Using g As Graphics = Graphics.FromImage(qrWithLogo)
                        g.Clear(Color.White)
                        g.DrawImage(qrImage, 0, 0)

                        Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                        Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                        Dim padding As Integer = 6
                        Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                        g.FillRectangle(Brushes.White, whiteRect)
                        g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
                    End Using

                    ' Convert QR to stream
                    Dim ms As New MemoryStream()
                    qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
                    ms.Position = 0

                    ' Email configuration
                    Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
                    Dim toAddress As New MailAddress(email)
                    Dim fromPassword As String = "swlqbwgztcqbneuw"

                    Dim message As New MailMessage(fromAddress, toAddress)
                    message.Subject = "Updated Professor QR Code"
                    message.Body = $"Hello {firstName}," & vbCrLf & vbCrLf &
                               "Your updated login QR code is attached." & vbCrLf &
                               "Scan it to log in using your new credentials."
                    message.Attachments.Add(New Attachment(ms, "UpdatedProfessorQR.png", "image/png"))

                    Dim smtp As New SmtpClient("smtp.gmail.com", 587)
                    smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
                    smtp.EnableSsl = True
                    smtp.Send(message)

                    MessageBox.Show("Updated QR code sent via email.", "Email Sent")
                Catch ex As Exception
                    MessageBox.Show("Failed to send QR email: " & ex.Message, "Email Error")
                Finally
                    Cursor.Current = Cursors.Default
                End Try
            End If
        Else
            MessageBox.Show("Failed to update professor.", "Error")
        End If
    End Sub

    Private Sub professorDeleteBtn_Click(sender As Object, e As EventArgs)
        If selectedProfessorId = -1 Then
            MessageBox.Show("Please select a professor to delete.", "Warning")
            Exit Sub
        End If

        Dim confirm As DialogResult = MessageBox.Show("Are you sure you want to delete this professor?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirm = DialogResult.Yes Then
            If DeleteProfessor(selectedProfessorId) Then
                MessageBox.Show("Professor successfully deleted!", "Success")
                ClearInputs(adminProfessorPanel) ' Optional: clear form fields
                LoadProfessorsToGrid()       ' Optional: reload GridView after deletion
                selectedProfessorId = -1         ' Reset the selection
            Else
                MessageBox.Show("Failed to delete professor.", "Error")
            End If
        End If
    End Sub
    Private Sub LoadAdminsToGrid()
        Dim surnameFilter As String = adminSurnameSearch.Text.Trim()
        Dim adminList As List(Of AdminModel) = GetAdminsByLastName(surnameFilter)

        adminView.Rows.Clear()

        For Each admin As AdminModel In adminList
            adminView.Rows.Add(admin.FirstName,
                               admin.LastName,
                               admin.MiddleInitial,
                               admin.Suffix,
                               admin.Email,
                               admin.Username,
                               admin.Password,
                               admin.Id)
        Next
    End Sub
    Private Sub adminSurnameSearch_TextChanged(sender As Object, e As EventArgs) Handles adminSurnameSearch.TextChanged
        LoadAdminsToGrid()
    End Sub

    ' Declare this at the form level
    Private selectedAdminId As Integer = -1

    Private Sub adminView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles adminView.CellContentClick
        If e.RowIndex >= 0 AndAlso e.RowIndex < adminView.Rows.Count Then
            Dim row As DataGridViewRow = adminView.Rows(e.RowIndex)

            adminFirstNameInput.Text = row.Cells("adminFirstName").Value.ToString()
            adminLastNameInput.Text = row.Cells("adminLastName").Value.ToString()
            adminMiddleInitialInput.Text = row.Cells("adminMiddleInitial").Value.ToString()

            ' Handle suffix safely
            Dim suffixValue As String = row.Cells("adminSuffix").Value?.ToString().Trim()
            If String.IsNullOrEmpty(suffixValue) OrElse Not adminSuffixBox.Items.Contains(suffixValue) Then
                adminSuffixBox.SelectedIndex = -1 ' Clear selection
            Else
                adminSuffixBox.SelectedItem = suffixValue
            End If

            adminEmailInput.Text = row.Cells("adminEmail").Value.ToString()
            adminUsernameInput2.Text = row.Cells("adminUsername").Value.ToString()
            adminPasswordInput2.Text = row.Cells("adminPassword").Value.ToString()

            selectedAdminId = Convert.ToInt32(row.Cells("adminId").Value)
        End If
    End Sub

    Private Sub adminAddBtn_Click(sender As Object, e As EventArgs) Handles adminAddBtn.Click
        ' Get input values
        Dim firstName As String = adminFirstNameInput.Text.Trim()
        Dim lastName As String = adminLastNameInput.Text.Trim()
        Dim middleInitial As String = adminMiddleInitialInput.Text.Trim()
        Dim suffix As String = If(adminSuffixBox.SelectedItem IsNot Nothing, adminSuffixBox.SelectedItem.ToString().Trim(), "")
        Dim email As String = adminEmailInput.Text.Trim()
        Dim username As String = adminUsernameInput2.Text.Trim()
        Dim password As String = adminPasswordInput2.Text.Trim()

        ' Validate input
        If String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       String.IsNullOrWhiteSpace(username) OrElse
       String.IsNullOrWhiteSpace(password) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) Then

            MessageBox.Show("Please enter valid and complete admin information.", "Validation Error")
            Exit Sub
        End If

        ' Insert admin
        Dim success As Boolean = InsertAdmin(firstName, lastName, middleInitial, suffix, email, username, password)

        If success Then
            MessageBox.Show("Admin successfully added!", "Success")
            ClearInputs(adminAdminPanel)
            LoadAdminsToGrid()

            Try
                Cursor.Current = Cursors.WaitCursor

                ' ✅ Step 1: Create QR content
                Dim qrContent As String = $"[{username}][{password}]"
                Dim qrGenerator As New QRCodeGenerator()
                Dim qrData As QRCodeData = qrGenerator.CreateQrCode(qrContent, QRCodeGenerator.ECCLevel.H)
                Dim qrCode As New QRCode(qrData)
                Dim qrImage As Bitmap = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

                ' Use embedded resource (no need for path or File.Exists)
                Dim logo As Image = My.Resources.PLP

                ' Resize logo
                Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
                Dim resizedLogo As New Bitmap(logoSize, logoSize)
                Using g As Graphics = Graphics.FromImage(resizedLogo)
                    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(logo, 0, 0, logoSize, logoSize)
                End Using

                ' ✅ Step 3: Merge logo with QR code
                Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
                Using g As Graphics = Graphics.FromImage(qrWithLogo)
                    g.Clear(Color.White)
                    g.DrawImage(qrImage, 0, 0)

                    ' Draw white background behind logo
                    Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                    Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                    Dim padding As Integer = 6
                    Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                    g.FillRectangle(Brushes.White, whiteRect)

                    ' Draw logo
                    g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
                End Using

                ' ✅ Step 4: Convert to stream
                Dim ms As New MemoryStream()
                qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
                ms.Position = 0

                ' ✅ Step 5: Email QR
                Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
                Dim toAddress As New MailAddress(email)
                Dim fromPassword As String = "swlqbwgztcqbneuw"

                Dim message As New MailMessage(fromAddress, toAddress)
                message.Subject = "Your Admin QR Code Login"
                message.Body = $"Hello {firstName}," & vbCrLf & vbCrLf &
                           "Attached is your QR code containing your login credentials." & vbCrLf &
                           "You can scan this QR to log in instead of typing your username and password."

                message.Attachments.Add(New Attachment(ms, "AdminLoginQRCode.png", "image/png"))

                Dim smtp As New SmtpClient("smtp.gmail.com", 587)
                smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
                smtp.EnableSsl = True
                smtp.Send(message)

                MessageBox.Show("QR Code sent to admin's email.", "Email Sent")

            Catch ex As Exception
                MessageBox.Show("Failed to send email: " & ex.Message, "Email Error")
            Finally
                Cursor.Current = Cursors.Default
            End Try
        Else
            MessageBox.Show("Failed to add admin.", "Error")
        End If
    End Sub


    Private Sub adminUpdateBtn_Click(sender As Object, e As EventArgs) Handles adminUpdateBtn.Click
        ' ✅ Retrieve input values
        Dim firstName As String = adminFirstNameInput.Text.Trim()
        Dim lastName As String = adminLastNameInput.Text.Trim()
        Dim middleInitial As String = adminMiddleInitialInput.Text.Trim()
        Dim suffix As String = If(adminSuffixBox.SelectedItem IsNot Nothing, adminSuffixBox.SelectedItem.ToString().Trim(), "")
        Dim email As String = adminEmailInput.Text.Trim()
        Dim newUsername As String = adminUsernameInput2.Text.Trim()
        Dim newPassword As String = adminPasswordInput2.Text.Trim()

        ' ✅ Make sure an admin is selected
        If selectedAdminId = -1 Then
            MessageBox.Show("Please select an admin to update.", "No Selection")
            Exit Sub
        End If

        ' ✅ Validate inputs
        If String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       String.IsNullOrWhiteSpace(newUsername) OrElse
       String.IsNullOrWhiteSpace(newPassword) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) Then

            MessageBox.Show("Please enter valid and complete admin information.", "Validation Error")
            Exit Sub
        End If

        ' ✅ Fetch existing record for comparison
        Dim oldData As DataRow = GetAdminById(selectedAdminId)
        If oldData Is Nothing Then
            MessageBox.Show("Admin not found in database.", "Error")
            Exit Sub
        End If

        Dim oldUsername As String = oldData("username").ToString().Trim()
        Dim oldPassword As String = oldData("password").ToString().Trim()

        Dim credentialsChanged As Boolean =
        (newUsername.Trim().ToLower() <> oldUsername.ToLower()) OrElse
        (newPassword.Trim() <> oldPassword)

        ' ✅ Update admin record
        Dim success As Boolean = UpdateAdmin(selectedAdminId, firstName, lastName, middleInitial, suffix, email, newUsername, newPassword)

        If success Then
            MessageBox.Show("Admin successfully updated!", "Success")
            ClearInputs(adminAdminPanel)
            LoadAdminsToGrid()
            selectedAdminId = -1

            ' ✅ Only send QR if credentials changed
            If credentialsChanged Then
                Try
                    Cursor.Current = Cursors.WaitCursor

                    ' QR content
                    Dim qrContent As String = $"[{newUsername}][{newPassword}]"
                    Dim qrGenerator As New QRCoder.QRCodeGenerator()
                    Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
                    Dim qrCode = New QRCoder.QRCode(qrData)
                    Dim qrImage = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

                    ' Use embedded resource (no need for path or File.Exists)
                    Dim logo As Image = My.Resources.PLP

                    ' Resize logo
                    Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
                    Dim resizedLogo As New Bitmap(logoSize, logoSize)
                    Using g As Graphics = Graphics.FromImage(resizedLogo)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        g.DrawImage(logo, 0, 0, logoSize, logoSize)
                    End Using


                    ' Combine QR and logo
                    Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
                    Using g As Graphics = Graphics.FromImage(qrWithLogo)
                        g.Clear(Color.White)
                        g.DrawImage(qrImage, 0, 0)

                        ' White background
                        Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                        Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                        Dim padding As Integer = 6
                        Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                        g.FillRectangle(Brushes.White, whiteRect)

                        ' Logo
                        g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
                    End Using

                    ' Stream for email
                    Dim ms As New MemoryStream()
                    qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
                    ms.Position = 0

                    ' Email setup
                    Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
                    Dim toAddress As New MailAddress(email)
                    Dim fromPassword As String = "swlqbwgztcqbneuw"


                    Dim message As New MailMessage(fromAddress, toAddress)
                    message.Subject = "Updated Admin QR Code"
                    message.Body = $"Hello {firstName}," & vbCrLf & vbCrLf &
                               "Your updated login QR code is attached." & vbCrLf &
                               "Scan it to log in using your new credentials."

                    message.Attachments.Add(New Attachment(ms, "UpdatedAdminQR.png", "image/png"))

                    Dim smtp As New SmtpClient("smtp.gmail.com", 587)
                    smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
                    smtp.EnableSsl = True
                    smtp.Send(message)

                    MessageBox.Show("Updated QR code sent via email.", "Email Sent")
                Catch ex As Exception
                    MessageBox.Show("Failed to send QR email: " & ex.Message, "Email Error")
                Finally
                    Cursor.Current = Cursors.Default
                End Try
            End If
        Else
            MessageBox.Show("Failed to update admin.", "Error")
        End If
    End Sub


    Private Sub adminDeleteBtn_Click(sender As Object, e As EventArgs) Handles adminDeleteBtn.Click
        ' Check if an admin is selected
        If selectedAdminId <= 0 Then
            MessageBox.Show("Please select an admin to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Confirm deletion
        Dim confirm As DialogResult = MessageBox.Show("Are you sure you want to delete this admin?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If confirm = DialogResult.Yes Then
            ' Call the delete function
            If DeleteAdmin(selectedAdminId) Then
                MessageBox.Show("Admin deleted successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' Optionally, refresh the DataGridView here (e.g., LoadAdmins())
                LoadAdminsToGrid() ' This should reload the DataGridView with updated records

                ' Clear form fields
                adminFirstNameInput.Clear()
                adminLastNameInput.Clear()
                adminMiddleInitialInput.Clear()
                adminSuffixBox.SelectedIndex = -1
                adminEmailInput.Clear()
                adminUsernameInput2.Clear()
                adminPasswordInput2.Clear()

                ' Reset selected ID
                selectedAdminId = -1
            Else
                MessageBox.Show("Failed to delete the admin.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub LoadReasonsToGrid()
        Dim keyword As String = reasonSearch.Text.Trim()
        Dim isSpecial As Nullable(Of Boolean) = Nothing

        If reasonSpecialBox.SelectedItem IsNot Nothing Then
            Select Case reasonSpecialBox.SelectedItem.ToString()
                Case "Special"
                    isSpecial = True
                Case "Normal"
                    isSpecial = False
            End Select

        End If

        Dim dt As DataTable = GetReasons(keyword, isSpecial)

        ' Prevent automatic column generation
        reasonView.AutoGenerateColumns = False
        reasonView.Rows.Clear()

        ' Manually add rows using your existing columns
        For Each row As DataRow In dt.Rows
            reasonView.Rows.Add(
      row("reason"),
      If(row("is_special") IsNot DBNull.Value AndAlso Convert.ToBoolean(row("is_special")), "Special", "Normal"),
      row("id")
  )

        Next
    End Sub
    Private Sub reasonSearch_TextChanged(sender As Object, e As EventArgs) Handles reasonSearch.TextChanged
        LoadReasonsToGrid()
    End Sub

    Private Sub reasonSpecialBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles reasonSpecialBox.SelectedIndexChanged
        LoadReasonsToGrid()
    End Sub

    Private selectedReasonId As Integer = -1
    Private Sub reasonView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles reasonView.CellContentClick
        ' Ensure the click is on a valid row (not header or outside range)
        If e.RowIndex >= 0 AndAlso e.RowIndex < reasonView.Rows.Count Then
            Dim row As DataGridViewRow = reasonView.Rows(e.RowIndex)

            ' Get values from the cells
            Dim reasonText As String = row.Cells("reason").Value.ToString()
            Dim specialText As String = row.Cells("special").Value.ToString()
            Dim reasonIdValue As Object = row.Cells("reasonId").Value

            ' Safely store the ID
            If reasonIdValue IsNot Nothing AndAlso IsNumeric(reasonIdValue) Then
                selectedReasonId = Convert.ToInt32(reasonIdValue)
            Else
                selectedReasonId = -1
            End If

            ' Set values to inputs
            reasonInput.Text = reasonText

            If specialText = "Special" Then
                specialBox.SelectedItem = "Special"
            ElseIf specialText = "Normal" Then
                specialBox.SelectedItem = "Normal"
            Else
                specialBox.SelectedIndex = -1
            End If
        End If
    End Sub

    Private Sub reasonAddBtn_Click(sender As Object, e As EventArgs) Handles reasonAddBtn.Click
        Dim reason As String = reasonInput.Text.Trim()
        Dim isSpecial As Boolean = (specialBox.SelectedItem?.ToString().ToLower() = "special")

        If reason = "" OrElse specialBox.SelectedItem Is Nothing Then
            MessageBox.Show("Please fill in both fields.", "Validation")
            Return
        End If

        InsertReason(reason, isSpecial)
        LoadReasonsToGrid()
        ClearInputs(adminReasonPanel)
        MessageBox.Show("Reason added.")
    End Sub
    Private Sub reasonUpdateBtn_Click(sender As Object, e As EventArgs) Handles reasonUpdateBtn.Click
        If selectedReasonId = -1 Then
            MessageBox.Show("Select a reason to update.")
            Return
        End If

        Dim reason As String = reasonInput.Text.Trim()
        Dim isSpecial As Boolean = (specialBox.SelectedItem?.ToString().ToLower() = "special")

        If reason = "" OrElse specialBox.SelectedItem Is Nothing Then
            MessageBox.Show("Please fill in both fields.", "Validation")
            Return
        End If

        UpdateReason(selectedReasonId, reason, isSpecial)
        LoadReasonsToGrid()
        ClearInputs(adminReasonPanel)
        MessageBox.Show("Reason updated.")
    End Sub
    Private Sub reasonDeleteBtn_Click(sender As Object, e As EventArgs) Handles reasonDeleteBtn.Click
        If selectedReasonId = -1 Then
            MessageBox.Show("Select a reason to delete.")
            Return
        End If

        Dim result = MessageBox.Show("Are you sure you want to delete this reason?", "Confirm", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            DeleteReason(selectedReasonId)
            LoadReasonsToGrid()
            ClearInputs(adminReasonPanel)
            MessageBox.Show("Reason deleted.")
        End If
    End Sub

    Private Sub LoadSectionsToGrid()
        Dim keyword As String = sectionSearch.Text.Trim()
        Dim dt As DataTable = GetSections(keyword)

        sectionView.AutoGenerateColumns = False
        sectionView.Rows.Clear()

        For Each row As DataRow In dt.Rows
            sectionView.Rows.Add(row("section"), row("id")) ' Adjust if your column names differ
        Next
    End Sub
    Private Sub sectionSearch_TextChanged(sender As Object, e As EventArgs) Handles sectionSearch.TextChanged
        LoadSectionsToGrid()
    End Sub

    Private selectedSectionId As Integer = -1 ' Declare this at the form level

    Private Sub sectionView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles sectionView.CellContentClick
        If e.RowIndex >= 0 AndAlso e.RowIndex < sectionView.Rows.Count Then
            Dim row As DataGridViewRow = sectionView.Rows(e.RowIndex)

            ' Updated column names to match your DataGridView
            If Not IsDBNull(row.Cells("sectionId").Value) Then
                selectedSectionId = Convert.ToInt32(row.Cells("sectionId").Value)
                sectionInput.Text = row.Cells("sectionName").Value.ToString()
            End If
        End If
    End Sub

    ' ADD Section
    Private Sub sectionAddBtn_Click(sender As Object, e As EventArgs) Handles sectionAddBtn.Click
        Dim name As String = sectionInput.Text.Trim()

        If String.IsNullOrWhiteSpace(name) Then
            MessageBox.Show("Please enter a section name.", "Validation Error")
            Return
        End If

        InsertSection(name)
        MessageBox.Show("Section added successfully.", "Success")
        LoadSectionsToGrid()
        sectionInput.Clear()
    End Sub

    ' UPDATE Section
    Private Sub sectionUpdateBtn_Click(sender As Object, e As EventArgs) Handles sectionUpdateBtn.Click
        Dim name As String = sectionInput.Text.Trim()

        If selectedSectionId <= 0 Then
            MessageBox.Show("Please select a section to update.", "Validation Error")
            Return
        End If

        If String.IsNullOrWhiteSpace(name) Then
            MessageBox.Show("Please enter a new section name.", "Validation Error")
            Return
        End If

        UpdateSection(selectedSectionId, name)
        MessageBox.Show("Section updated successfully.", "Success")
        LoadSectionsToGrid()
        sectionInput.Clear()
        selectedSectionId = -1
    End Sub

    ' DELETE Section
    Private Sub sectionDeleteBtn_Click(sender As Object, e As EventArgs) Handles sectionDeleteBtn.Click
        If selectedSectionId <= 0 Then
            MessageBox.Show("Please select a section to delete.", "Validation Error")
            Return
        End If

        Dim confirm As DialogResult = MessageBox.Show("Are you sure you want to delete this section?", "Confirm Delete", MessageBoxButtons.YesNo)

        If confirm = DialogResult.Yes Then
            DeleteSection(selectedSectionId)
            MessageBox.Show("Section deleted successfully.", "Success")
            LoadSectionsToGrid()
            sectionInput.Clear()
            selectedSectionId = -1
        End If
    End Sub


    Private Sub setGraduationBtn_Click(sender As Object, e As EventArgs) Handles setGraduationBtn.Click
        Try
            Dim selectedDate As Date = graduationDate.Value
            SetGraduationDate(selectedDate)
            MessageBox.Show("Graduation date updated successfully!", "Success")
            LoadCurrentGraduationDate()
        Catch ex As Exception
            MessageBox.Show("Failed to update graduation date: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub LoadCurrentGraduationDate()
        Try
            Dim gradDate As Date = GetLatestGraduationDate()
            currentGraduationText.Text = "Current Graduation Date: " & gradDate.ToString("MMMM dd, yyyy")
        Catch ex As Exception
            currentGraduationText.Text = "Not Set"
        End Try
    End Sub

    ' Call this method to load/reload archive data
    Private Sub LoadArchives()
        Dim surnameFilter As String = archiveSearchSurname.Text.Trim()
        Dim yearFilter As String = If(adminSearchYear.SelectedItem IsNot Nothing, adminSearchYear.SelectedItem.ToString(), "")

        ' Get filtered data
        Dim dt As DataTable = Archive.SearchArchives(surnameFilter, yearFilter)

        ' Clear existing rows
        archiveView.Rows.Clear()

        ' Populate each row manually
        For Each row As DataRow In dt.Rows
            archiveView.Rows.Add(
        row("student_number"),
        row("first_name"),
        row("last_name"),
        row("middle_initial"),
        row("suffix"),
        row("section_name"), ' already joined in SQL
        row("email"),
        row("status"),
        row("year_graduated"),
        row("id")
    )
        Next
    End Sub

    Private Sub archiveSearchSurname_TextChanged(sender As Object, e As EventArgs) Handles archiveSearchSurname.TextChanged
        LoadArchives()
    End Sub

    Private Sub adminSearchYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles adminSearchYear.SelectedIndexChanged
        LoadArchives()
    End Sub
    Private Sub LoadArchiveYears()
        adminSearchYear.Items.Clear()

        Dim dt As DataTable = ExecuteQuery("SELECT DISTINCT year_graduated FROM students WHERE isGraduate = 1 ORDER BY year_graduated DESC")

        For Each row As DataRow In dt.Rows
            adminSearchYear.Items.Add(row("year_graduated").ToString())
        Next

        adminSearchYear.SelectedIndex = -1 ' optional: nothing selected by default
    End Sub

    Private Sub LoadCounts()
        ' Admin count
        Dim adminDt As DataTable = ExecuteQuery("SELECT COUNT(*) AS total FROM admins")
        If adminDt.Rows.Count > 0 Then
            adminCount.Text = adminDt.Rows(0)("total").ToString()
        End If

        ' Professor count
        Dim profDt As DataTable = ExecuteQuery("SELECT COUNT(*) AS total FROM professors")
        If profDt.Rows.Count > 0 Then
            professorCount.Text = profDt.Rows(0)("total").ToString()
        End If

        ' Archive count - UPDATED QUERY
        Dim archiveDt As DataTable = ExecuteQuery("SELECT COUNT(*) AS total FROM students WHERE isGraduate = 1")
        If archiveDt.Rows.Count > 0 Then
            archiveCount.Text = archiveDt.Rows(0)("total").ToString()
        End If

        ' Student count - You might also want to update this to show only active students
        Dim studentDt As DataTable = ExecuteQuery("SELECT COUNT(*) AS total FROM students WHERE isGraduate = 0")
        If studentDt.Rows.Count > 0 Then
            studentCount.Text = studentDt.Rows(0)("total").ToString()
        End If
    End Sub

    Private Sub LoadChart2()
        ' Clear existing chart data
        Chart2.Series.Clear()
        Chart2.Titles.Clear()

        ' Create new doughnut chart series
        Dim series As New DataVisualization.Charting.Series("Reasons")
        series.ChartType = DataVisualization.Charting.SeriesChartType.Doughnut
        series.IsValueShownAsLabel = True
        series.LabelForeColor = System.Drawing.Color.White

        Try
            Connect()

            ' Get year from dashboardYear label
            Dim selectedYear As Integer
            If Not Integer.TryParse(dashboardYear.Text.Trim(), selectedYear) Then
                MessageBox.Show("Invalid year format in dashboardYear.")
                Exit Sub
            End If

            ' Define date range
            Dim fromDate As String = $"{selectedYear}-01-01"
            Dim toDate As String = $"{selectedYear}-12-31"

            ' SQL query with date filtering
            Dim query As String = "SELECT r.reason, COUNT(rep.reason_id) AS count " &
                      "FROM reports rep " &
                      "JOIN reasons r ON rep.reason_id = r.id " &
                      "WHERE rep.consultation_date BETWEEN @fromDate AND @toDate " &
                      "GROUP BY rep.reason_id"


            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@fromDate", fromDate)
                cmd.Parameters.AddWithValue("@toDate", toDate)

                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim reason As String = reader("reason").ToString()
                        Dim count As Integer = Convert.ToInt32(reader("count"))
                        series.Points.AddXY(reason, count)
                    End While
                End Using
            End Using

            ' Add series and title to chart
            Chart2.Series.Add(series)

            Dim title As New DataVisualization.Charting.Title()
            title.Text = $"Consultation Reasons - {selectedYear}"
            title.Font = New System.Drawing.Font("Segoe UI", 16, System.Drawing.FontStyle.Bold)
            title.ForeColor = System.Drawing.Color.FromArgb(40, 40, 40)
            title.Alignment = ContentAlignment.TopCenter
            Chart2.Titles.Add(title)

        Catch ex As Exception
            MessageBox.Show("Failed to load chart: " & ex.Message, "Chart Error")
        End Try
    End Sub

    Private Sub LoadChart3()
        ' Clear chart
        Chart3.Series.Clear()
        Chart3.Titles.Clear()

        ' Get year from dashboardYear textbox
        Dim selectedYear As Integer
        If Not Integer.TryParse(dashboardYear.Text, selectedYear) Then
            MessageBox.Show("Invalid year selected.", "Input Error")
            Exit Sub
        End If

        Try
            Connect()
            If conn.State <> ConnectionState.Open Then
                MessageBox.Show("Database connection failed", "Connection Error")
                Exit Sub
            End If

            ' Define date range
            Dim fromDate As DateTime = New DateTime(selectedYear, 1, 1)
            Dim toDate As DateTime = New DateTime(selectedYear, 12, 31)

            ' Dictionary to count total consultations per month
            Dim monthlyCounts As New Dictionary(Of Integer, Integer)()
            For m As Integer = 1 To 12
                monthlyCounts(m) = 0 ' Default to zero
            Next

            ' Query to get all consultation dates and count grouped by month
            Dim query As String = "
        SELECT MONTH(consultation_date) AS month, COUNT(*) AS count
        FROM reports
        WHERE consultation_date BETWEEN @fromDate AND @toDate
        GROUP BY MONTH(consultation_date);
        "

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@fromDate", fromDate.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@toDate", toDate.ToString("yyyy-MM-dd"))

                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim month As Integer = Convert.ToInt32(reader("month"))
                        Dim count As Integer = Convert.ToInt32(reader("count"))
                        monthlyCounts(month) = count
                    End While
                End Using
            End Using

            ' Create Area chart series
            Dim series As New DataVisualization.Charting.Series("Consultations (" & selectedYear & ")")
            series.ChartType = DataVisualization.Charting.SeriesChartType.Area
            series.IsValueShownAsLabel = False
            series.BorderWidth = 2
            series.Color = System.Drawing.Color.MediumSeaGreen

            ' Add 12 points for each month
            For m As Integer = 1 To 12
                Dim label As String = New Date(selectedYear, m, 1).ToString("MMM")
                series.Points.AddXY(label, monthlyCounts(m))
            Next

            ' Add to chart
            Chart3.Series.Add(series)

            ' Configure chart area
            Dim chartArea As DataVisualization.Charting.ChartArea = Chart3.ChartAreas(0)
            chartArea.AxisX.Title = "Month"
            chartArea.AxisY.Title = "Consultation Count"
            chartArea.AxisX.Interval = 1
            chartArea.AxisY.Minimum = 0

            ' Add title
            Dim title As New DataVisualization.Charting.Title(
            "Consultations (" & selectedYear & ")", Docking.Top,
            New System.Drawing.Font("Segoe UI", 14, System.Drawing.FontStyle.Bold),
            System.Drawing.Color.FromArgb(40, 40, 40)
        )
            Chart3.Titles.Add(title)

            ' Refresh chart
            Chart3.Refresh()

        Catch ex As Exception
            MessageBox.Show("Failed to load chart: " & ex.Message, "Chart Error")
        Finally
            If conn IsNot Nothing AndAlso conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub



    Private Sub LoadProfessorCounts()
        Connect()

        Try
            Dim todayCountVal As Integer = 0
            Dim weekCountVal As Integer = 0
            Dim uniqueStudentCount As Integer = 0

            ' Total consultations today
            Dim queryToday As String = "
            SELECT COUNT(*) 
            FROM reports 
            WHERE professor_id = @profId 
              AND consultation_date = CURDATE();"

            ' Total consultations this week
            Dim queryWeek As String = "
            SELECT COUNT(*) 
            FROM reports 
            WHERE professor_id = @profId 
              AND YEARWEEK(consultation_date, 1) = YEARWEEK(CURDATE(), 1);"

            ' Total unique students consulted
            Dim queryUnique As String = "
            SELECT COUNT(DISTINCT student_id) 
            FROM reports 
            WHERE professor_id = @profId;"

            ' Execute Today
            Using cmd As New MySqlCommand(queryToday, conn)
                cmd.Parameters.AddWithValue("@profId", CurrentProfessor.Id)
                todayCountVal = Convert.ToInt32(cmd.ExecuteScalar())
            End Using

            ' Execute Week
            Using cmd As New MySqlCommand(queryWeek, conn)
                cmd.Parameters.AddWithValue("@profId", CurrentProfessor.Id)
                weekCountVal = Convert.ToInt32(cmd.ExecuteScalar())
            End Using

            ' Execute Unique Student Count
            Using cmd As New MySqlCommand(queryUnique, conn)
                cmd.Parameters.AddWithValue("@profId", CurrentProfessor.Id)
                uniqueStudentCount = Convert.ToInt32(cmd.ExecuteScalar())
            End Using

            ' Set label texts
            todayCount.Text = todayCountVal.ToString()
            weekCount.Text = weekCountVal.ToString()
            studentConsultCount.Text = uniqueStudentCount.ToString()

        Catch ex As Exception
            MessageBox.Show("Error loading professor counts: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub LoadConsultations()
        Dim fromDate As Date = consultFromDate.Value.Date
        Dim toDate As Date = consultToDate.Value.Date

        ' Safe read of section selection
        Dim selectedSectionId As Integer? = Nothing
        If professorSectionBox.SelectedIndex > 0 AndAlso professorSectionBox.SelectedValue IsNot Nothing Then
            Dim val = professorSectionBox.SelectedValue
            If IsNumeric(val) Then
                selectedSectionId = Convert.ToInt32(val)
            End If
        End If

        ' Fetch filtered reports
        Dim dt As DataTable = GetFormattedReportsByProfessor(CurrentProfessor.Id, fromDate, toDate, selectedSectionId)

        ' Clear grid
        consultView.Rows.Clear()

        ' Load data
        For Each row As DataRow In dt.Rows
            consultView.Rows.Add(
            row("student_number").ToString(),
            row("student_name").ToString(),
            row("section").ToString(),
            row("reason").ToString(),
            row("message").ToString(),
            row("consultation_date").ToString(),
            row("time_in").ToString(),
            row("time_out").ToString()
        )
        Next
    End Sub

    Private Sub LoadSectionsIntoComboBox8()
        Dim dt As DataTable = GetAllSections()

        dt.Columns.Add("DisplayText", GetType(String))
        For Each row As DataRow In dt.Rows
            row("DisplayText") = row("section").ToString()
        Next

        ' Add "-- All Sections --" as default item
        Dim allRow As DataRow = dt.NewRow()
        allRow("id") = -1
        allRow("DisplayText") = "-- All Sections --"
        dt.Rows.InsertAt(allRow, 0)

        professorSectionBox.DataSource = dt
        professorSectionBox.DisplayMember = "DisplayText"
        professorSectionBox.ValueMember = "id"
        professorSectionBox.SelectedIndex = 0 ' Select "-- All Sections --" by default
    End Sub


    Private Sub professorSectionBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles professorSectionBox.SelectedIndexChanged
        LoadConsultations()
    End Sub
    Private Sub consultFromDate_ValueChanged(sender As Object, e As EventArgs) Handles consultFromDate.ValueChanged
        LoadConsultations()
    End Sub
    Private Sub consultToDate_ValueChanged(sender As Object, e As EventArgs) Handles consultToDate.ValueChanged
        LoadConsultations()
    End Sub

    Public Sub GenerateProfessorReportPDF(prof As ProfessorModel, fromDate As Date, toDate As Date)
        Dim profId As Integer = prof.Id
        Dim dt As DataTable = GetFormattedReportsByProfessor(profId, fromDate, toDate)

        If dt.Rows.Count = 0 Then
            MessageBox.Show("No records found for the selected date range.", "No Data")
            Return
        End If

        Dim reasonCounts As New Dictionary(Of String, Integer)()
        Dim dailyCounts As New Dictionary(Of Date, Integer)()

        ' Initialize daily counts with 0 for each day between fromDate and toDate
        Dim curDate As Date = fromDate
        While curDate <= toDate
            dailyCounts(curDate) = 0
            curDate = curDate.AddDays(1)
        End While

        Try
            Connect()

            ' Reason counts
            Dim query As String = "SELECT r.reason, COUNT(rep.reason_id) AS count " &
                              "FROM reports rep " &
                              "JOIN reasons r ON rep.reason_id = r.id " &
                              "WHERE rep.professor_id = @profId " &
                              "AND rep.consultation_date BETWEEN @fromDate AND @toDate " &
                              "GROUP BY rep.reason_id"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@profId", profId)
                cmd.Parameters.AddWithValue("@fromDate", fromDate.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@toDate", toDate.ToString("yyyy-MM-dd"))
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        reasonCounts(reader("reason").ToString()) = Convert.ToInt32(reader("count"))
                    End While
                End Using
            End Using

            ' Daily consultation counts
            Dim dailyQuery As String = "SELECT consultation_date, COUNT(*) AS count FROM reports WHERE professor_id = @profId AND consultation_date BETWEEN @fromDate AND @toDate GROUP BY consultation_date"
            Using cmd As New MySqlCommand(dailyQuery, conn)
                cmd.Parameters.AddWithValue("@profId", profId)
                cmd.Parameters.AddWithValue("@fromDate", fromDate.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@toDate", toDate.ToString("yyyy-MM-dd"))
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim dateKey As Date = Convert.ToDateTime(reader("consultation_date"))
                        dailyCounts(dateKey) = Convert.ToInt32(reader("count"))
                    End While
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Failed to fetch reason distribution: " & ex.Message, "DB Error")
        End Try

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
        saveFileDialog.FileName = $"ProfessorReport_{fromDate:yyyyMMdd}_{toDate:yyyyMMdd}.pdf"

        If saveFileDialog.ShowDialog() <> DialogResult.OK Then Exit Sub

        Try
            Dim document As New PdfDocument()
            document.Info.Title = "Consultation Report"

            ' Setup fonts
            Dim fontTitle As New XFont("Segoe UI", 16, XFontStyle.Bold)
            Dim fontSubTitle As New XFont("Segoe UI", 9, XFontStyle.Regular)
            Dim fontHeader As New XFont("Segoe UI", 9, XFontStyle.Bold)
            Dim fontRow As New XFont("Segoe UI", 8, XFontStyle.Regular)
            Dim fontSmall As New XFont("Segoe UI", 7, XFontStyle.Regular)

            ' === Load left logo from relative path ===
            Dim leftLogoPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Resources\plplogo.png")
            Dim leftLogo As XImage = Nothing
            If File.Exists(leftLogoPath) Then
                leftLogo = XImage.FromFile(leftLogoPath)
            End If

            ' === Load right logo from settings ===
            Dim rightLogoPath As String = My.Settings.LogoPath
            Dim rightLogo As XImage = Nothing
            If File.Exists(rightLogoPath) Then
                rightLogo = XImage.FromFile(rightLogoPath)
            End If

            ' Full professor name (if used)
            Dim fullProfName As String = prof.FirstName & " " & prof.MiddleInitial & ". " & prof.LastName &
                                         If(String.IsNullOrWhiteSpace(prof.Suffix), "", ", " & prof.Suffix)

            ' Draw header
            Dim DrawHeader = Sub(hdrGfx As XGraphics, hdrPage As PdfPage)
                                 ' Background rectangle
                                 hdrGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(245, 245, 245)), 0, 0, hdrPage.Width.Point, 80)

                                 ' Top-left text: small, single line
                                 Dim xLeft As Double = 20
                                 Dim yTop As Double = 5
                                 Dim smallFont As New XFont("Segoe UI", 7, XFontStyle.Regular)
                                 Dim mi As String = If(String.IsNullOrWhiteSpace(CurrentProfessor.MiddleInitial), "", " " & CurrentProfessor.MiddleInitial & ".")
                                 Dim suffix As String = If(String.IsNullOrWhiteSpace(CurrentProfessor.Suffix), "", " " & CurrentProfessor.Suffix)
                                 Dim fullName As String = $"{CurrentProfessor.FirstName}{mi} {CurrentProfessor.LastName}{suffix}".Trim()

                                 Dim generatedText As String = $"Report generated by Prof. {fullName} as of {DateTime.Now:MMMM dd, yyyy, 'at' hh:mm tt}"

                                 hdrGfx.DrawString(generatedText, smallFont, XBrushes.Black, New XRect(xLeft, yTop, 400, 10), XStringFormats.TopLeft)



                                 ' === LOGOS (same height, keep aspect ratio) ===
                                 Dim targetHeight As Double = 50

                                 ' Left logo
                                 If leftLogo IsNot Nothing Then
                                     Dim aspectLeft As Double = leftLogo.PixelWidth / leftLogo.PixelHeight
                                     Dim newWidthLeft As Double = targetHeight * aspectLeft
                                     hdrGfx.DrawImage(leftLogo, xLeft, yTop + 15, newWidthLeft, targetHeight)
                                 End If

                                 ' Right logo
                                 If rightLogo IsNot Nothing Then
                                     Dim aspectRight As Double = rightLogo.PixelWidth / rightLogo.PixelHeight
                                     Dim newWidthRight As Double = targetHeight * aspectRight
                                     hdrGfx.DrawImage(rightLogo, hdrPage.Width.Point - newWidthRight - 20, yTop + 15, newWidthRight, targetHeight)
                                 End If

                                 ' Centered text below logos
                                 Dim centerX As Double = hdrPage.Width.Point / 2
                                 Dim yCenter As Double = yTop + 20
                                 hdrGfx.DrawString("Consultation Report", fontTitle, XBrushes.Black, New XRect(centerX - 150, yCenter, 300, 20), XStringFormats.TopCenter)
                                 hdrGfx.DrawString("Pamantasan ng Lungsod ng Pasig", fontSubTitle, XBrushes.Black, New XRect(centerX - 150, yCenter + 20, 300, 15), XStringFormats.TopCenter)
                                 hdrGfx.DrawString(My.Settings.DepartmentName, fontSubTitle, XBrushes.Black, New XRect(centerX - 150, yCenter + 35, 300, 15), XStringFormats.TopCenter)
                             End Sub

            ' === PAGE 1 ===
            Dim chartPage As PdfPage = document.AddPage()
            chartPage.Size = PageSize.A4
            Dim chartGfx As XGraphics = XGraphics.FromPdfPage(chartPage)
            DrawHeader(chartGfx, chartPage)

            Dim chartY As Double = 80
            Dim totalConsultations As Integer = dt.Rows.Count
            Dim uniqueStudents As Integer = dt.AsEnumerable().Select(Function(r) r.Field(Of String)("student_number")).Distinct().Count()

            ' === Summary Boxes ===
            Dim cardFontTitle As New XFont("Segoe UI", 9, XFontStyle.Bold)
            Dim cardFontValue As New XFont("Segoe UI", 14, XFontStyle.Bold)
            Dim cardWidth As Double = 240

            Dim cardHeight As Double = 50
            Dim cardY As Double = chartY + 20
            Dim cardGap As Double = 30
            Dim startX As Double = 30


            chartGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(230, 248, 255)), startX, cardY, cardWidth, cardHeight)
            chartGfx.DrawRectangle(XPens.DeepSkyBlue, startX, cardY, cardWidth, cardHeight)
            chartGfx.DrawString("Total Consultations", cardFontTitle, XBrushes.Black, New XPoint(startX + 10, cardY + 15))
            chartGfx.DrawString(totalConsultations.ToString(), cardFontValue, XBrushes.Black, New XPoint(startX + 10, cardY + 38))

            Dim secondCardX As Double = startX + cardWidth + cardGap

            chartGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(230, 255, 240)), secondCardX, cardY, cardWidth, cardHeight)
            chartGfx.DrawRectangle(XPens.MediumSeaGreen, secondCardX, cardY, cardWidth, cardHeight)
            chartGfx.DrawString("Unique Students", cardFontTitle, XBrushes.Black, New XPoint(secondCardX + 10, cardY + 15))
            chartGfx.DrawString(uniqueStudents.ToString(), cardFontValue, XBrushes.Black, New XPoint(secondCardX + 10, cardY + 38))

            ' === Bar Chart (left-aligned with HSV dynamic colors) ===
            Dim barYStart As Double = cardY + cardHeight + 40
            Dim barBottomY As Double = barYStart ' To track where the bar chart ends

            If reasonCounts.Count > 0 Then
                Dim total As Integer = reasonCounts.Values.Sum()
                Dim currentY As Double = barYStart
                Dim barXStart As Double = startX
                Dim maxWidth As Double = 400

                Dim barHeight As Double = 14
                Dim spacing As Double = 6
                Dim colorIndex As Integer = 0
                Dim maxValue As Integer = reasonCounts.Values.Max()

                chartGfx.DrawString("Reason Distribution", fontHeader, XBrushes.Black, New XPoint(barXStart, currentY - 15))

                ' Draw bars
                For Each pair In reasonCounts
                    Dim barWidth As Double = (pair.Value / maxValue) * maxWidth
                    Dim brush As XBrush = GetColorByIndex(colorIndex)

                    chartGfx.DrawRectangle(brush, barXStart, currentY, barWidth, barHeight)
                    currentY += barHeight + spacing
                    colorIndex += 1
                Next

                ' Draw legend
                Dim legendX As Double = barXStart + maxWidth + 20
                Dim legendY As Double = barYStart
                colorIndex = 0

                For Each pair In reasonCounts
                    Dim brush As XBrush = GetColorByIndex(colorIndex)
                    chartGfx.DrawRectangle(brush, legendX, legendY, 10, 10)
                    chartGfx.DrawString(pair.Key & " (" & pair.Value & ")", fontSmall, XBrushes.Black, New XPoint(legendX + 14, legendY + 9))
                    legendY += 14
                    colorIndex += 1
                Next

                barBottomY = Math.Max(currentY, legendY)
            End If












            ' === Prepare dailyCounts with ALL dates from fromDate to toDate ===
            Dim lineCounts As New Dictionary(Of Date, Integer)()

            Dim currentDate As Date = fromDate
            While currentDate <= toDate
                dailyCounts(currentDate) = 0
                currentDate = currentDate.AddDays(1)
            End While

            ' Populate actual consultation counts
            For Each row As DataRow In dt.Rows
                Dim dateKey As Date = CDate(row("consultation_date"))
                If dailyCounts.ContainsKey(dateKey) Then
                    dailyCounts(dateKey) += 1
                End If
            Next

            ' === Line Chart Drawing ===
            Dim maxLineVal As Integer = If(dailyCounts.Values.Count > 0, dailyCounts.Values.Max(), 1)
            If maxLineVal = 0 Then maxLineVal = 1

            Dim chartHeight As Double = 100
            Dim chartWidth As Double = 550
            Dim chartXStart As Double = startX
            Dim chartYStart As Double = barBottomY + 40

            Dim pointsPerChart As Integer = 20


            Dim allDates = dailyCounts.Keys.OrderBy(Function(d) d).ToList()
            Dim totalChunks As Integer = Math.Ceiling(allDates.Count / pointsPerChart)

            For chunkIndex As Integer = 0 To totalChunks - 1
                ' Check if current Y exceeds page height limit
                ' Check if current Y exceeds page height limit
                If chartYStart + chartHeight + 30 > chartPage.Height Then
                    ' Add new page
                    chartPage = document.AddPage()
                    chartPage.Size = PageSize.A4
                    chartGfx = XGraphics.FromPdfPage(chartPage)
                    DrawHeader(chartGfx, chartPage)

                    ' Reset Y position on new page
                    chartYStart = 110
                End If


                Dim chunkDates = allDates.Skip(chunkIndex * pointsPerChart).Take(pointsPerChart).ToList()
                If chunkDates.Count = 0 Then Exit For

                chartGfx.DrawString("Consultations Per Day (Part " & (chunkIndex + 1).ToString() & ")", fontHeader, XBrushes.Black, New XPoint(chartXStart, chartYStart - 15))

                Dim pointGap As Double = chartWidth / Math.Max(1, chunkDates.Count - 1)

                ' Draw Y-axis grid and values
                Dim yStep As Integer = Math.Max(1, Math.Ceiling(maxLineVal / 5))
                For i As Integer = 0 To maxLineVal Step yStep
                    Dim yVal As Double = chartYStart + chartHeight - (i / maxLineVal * chartHeight)
                    chartGfx.DrawLine(XPens.LightGray, chartXStart, yVal, chartXStart + chartWidth, yVal)
                    chartGfx.DrawString(i.ToString(), fontSmall, XBrushes.Black, New XPoint(chartXStart - 20, yVal - 4))
                Next

                ' Draw chart axis
                chartGfx.DrawLine(XPens.Black, chartXStart, chartYStart, chartXStart, chartYStart + chartHeight)
                chartGfx.DrawLine(XPens.Black, chartXStart, chartYStart + chartHeight, chartXStart + chartWidth, chartYStart + chartHeight)

                ' Draw data points and lines
                Dim prevX As Double = 0
                Dim prevY As Double = 0

                For index As Integer = 0 To chunkDates.Count - 1
                    Dim dateKey = chunkDates(index)
                    Dim value = dailyCounts(dateKey)
                    Dim x = chartXStart + index * pointGap
                    Dim y = chartYStart + chartHeight - (value / maxLineVal * chartHeight)

                    chartGfx.DrawEllipse(XBrushes.DarkBlue, x - 2, y - 2, 4, 4)
                    If index > 0 Then
                        chartGfx.DrawLine(XPens.DarkBlue, prevX, prevY, x, y)
                    End If

                    ' Always draw every date label
                    chartGfx.DrawString(dateKey.ToString("MMM d"), fontSmall, XBrushes.Black,
                            New XRect(x - 15, chartYStart + chartHeight + 2, 30, 10), XStringFormats.TopCenter)

                    prevX = x
                    prevY = y
                Next

                ' Advance Y position for next row of line graph
                chartYStart += chartHeight + 60
            Next


            ' Page 2: Consultation table
            Dim page As PdfPage = document.AddPage()
            page.Size = PageSize.A4
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            Dim tf As New XTextFormatter(gfx)
            DrawHeader(gfx, page)

            Dim yPoint As Double = 90 ' <- Slight margin before the table
            Dim marginLeft As Integer = 10
            Dim lineHeight As Double = 12
            Dim colWidths() As Integer = {60, 90, 40, 70, 160, 60, 50, 50}
            Dim headers() As String = {"Student No.", "Name", "Section", "Reason", "Message", "Date", "Time In", "Time Out"}

            Dim xPos As Double = marginLeft
            For i = 0 To headers.Length - 1
                gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopLeft)
                xPos += colWidths(i)
            Next
            yPoint += lineHeight + 6

            Dim rowColorToggle As Boolean = False

            For Each row As DataRow In dt.Rows
                xPos = marginLeft
                Dim maxRowHeight As Double = 0
                Dim rowData() As String = {
                row("student_number").ToString(),
                row("student_name").ToString(),
                row("section").ToString(),
                row("reason").ToString(),
                ForceWrap(row("message").ToString()),
                row("consultation_date").ToString(),
                row("time_in").ToString(),
                row("time_out").ToString()
            }

                For i = 0 To rowData.Length - 1
                    Dim layoutRect As New XRect(0, 0, colWidths(i) - 4, Double.MaxValue)
                    Dim dummyGfx As XGraphics = XGraphics.CreateMeasureContext(New XSize(colWidths(i) - 4, Double.MaxValue), XGraphicsUnit.Point, XPageDirection.Downwards)
                    Dim size As XSize = dummyGfx.MeasureString(rowData(i), fontRow)
                    Dim linesNeeded As Integer = Math.Ceiling(size.Width / layoutRect.Width)
                    Dim heightNeeded As Double = linesNeeded * lineHeight
                    If heightNeeded > maxRowHeight Then maxRowHeight = heightNeeded
                Next

                If maxRowHeight < lineHeight * 2 Then maxRowHeight = lineHeight * 2

                If rowColorToggle Then
                    gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(240, 240, 240)), marginLeft, yPoint, colWidths.Sum(), maxRowHeight)
                End If

                xPos = marginLeft
                For i = 0 To rowData.Length - 1
                    tf.DrawString(rowData(i), fontRow, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, maxRowHeight), XStringFormats.TopLeft)
                    gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), maxRowHeight)
                    xPos += colWidths(i)
                Next

                yPoint += maxRowHeight
                rowColorToggle = Not rowColorToggle

                If yPoint > page.Height.Point - 40 Then
                    page = document.AddPage()
                    page.Size = PageSize.A4
                    gfx = XGraphics.FromPdfPage(page)
                    tf = New XTextFormatter(gfx)
                    DrawHeader(gfx, page)

                    yPoint = 90 ' Maintain top margin
                    xPos = marginLeft
                    For i = 0 To headers.Length - 1
                        gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                        gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopLeft)
                        xPos += colWidths(i)
                    Next
                    yPoint += lineHeight + 6
                End If
            Next

            document.Save(saveFileDialog.FileName)
            MessageBox.Show("PDF Report generated successfully!", "Success")

        Catch ex As Exception
            MessageBox.Show("Failed to generate PDF: " & ex.Message, "PDF Error")
        End Try
    End Sub
    Private Function GetColorByIndex(index As Integer) As XBrush
        Dim hue As Double = (index * 137.508) Mod 360 ' Golden angle to spread hues
        Dim color As Color = ColorFromHSV(hue, 0.6, 0.85)
        Return New XSolidBrush(XColor.FromArgb(color.R, color.G, color.B))
    End Function

    Private Function ColorFromHSV(hue As Double, saturation As Double, value As Double) As Color
        Dim hi As Integer = CInt(Math.Floor(hue / 60)) Mod 6
        Dim f As Double = hue / 60 - Math.Floor(hue / 60)

        value *= 255
        Dim v As Integer = CInt(value)
        Dim p As Integer = CInt(value * (1 - saturation))
        Dim q As Integer = CInt(value * (1 - f * saturation))
        Dim t As Integer = CInt(value * (1 - (1 - f) * saturation))

        Select Case hi
            Case 0 : Return Color.FromArgb(255, v, t, p)
            Case 1 : Return Color.FromArgb(255, q, v, p)
            Case 2 : Return Color.FromArgb(255, p, v, t)
            Case 3 : Return Color.FromArgb(255, p, q, v)
            Case 4 : Return Color.FromArgb(255, t, p, v)
            Case Else : Return Color.FromArgb(255, v, p, q)
        End Select
    End Function

    Private Function ForceWrap(text As String, Optional interval As Integer = 30) As String
        Dim result As New System.Text.StringBuilder()
        For i = 0 To text.Length - 1
            result.Append(text(i))
            If (i + 1) Mod interval = 0 Then
                result.Append(" ")
            End If
        Next
        Return result.ToString()
    End Function



    Private Sub professorGenerateReportBtn_Click(sender As Object, e As EventArgs) Handles professorGenerateReportBtn.Click
        ' Validate professor and date range
        If CurrentProfessor Is Nothing Then
            MessageBox.Show("Professor data is not loaded.", "Missing Data")
            Return
        End If

        Dim fromDate As Date = consultFromDate.Value.Date
        Dim toDate As Date = consultToDate.Value.Date

        If fromDate > toDate Then
            MessageBox.Show("Start date cannot be after end date.", "Invalid Date Range")
            Return
        End If

        ' Confirmation dialog before generating
        Dim msg As String = $"Generate report for the date range:" & vbCrLf &
                        $"{fromDate:MMMM dd, yyyy} to {toDate:MMMM dd, yyyy}?" & vbCrLf &
                        "Do you want to proceed?"
        Dim result As DialogResult = MessageBox.Show(msg, "Confirm Report Generation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.No Then
            Return ' Cancel if user selects No
        End If

        ' Generate PDF report
        GenerateProfessorReportPDF(CurrentProfessor, fromDate, toDate)
    End Sub

    ' Placeholder text logic
    Private Sub studentMessageInput_Enter(sender As Object, e As EventArgs) Handles studentMessageInput.Enter
        If studentMessageInput.Text = "Describe your concern" Then
            studentMessageInput.Text = ""
            studentMessageInput.ForeColor = Color.Black
        End If
    End Sub

    Private Sub studentMessageInput_Leave(sender As Object, e As EventArgs) Handles studentMessageInput.Leave
        If String.IsNullOrWhiteSpace(studentMessageInput.Text) Then
            SetPlaceholder()
        Else
            studentMessageInput.ForeColor = Color.Black ' Ensure input text is black when valid
        End If
    End Sub

    Private Sub SetPlaceholder()
        studentMessageInput.Text = "Describe your concern"
        studentMessageInput.ForeColor = Color.Black
    End Sub


    Private Sub studentMessageInput_TextChanged(sender As Object, e As EventArgs) Handles studentMessageInput.TextChanged

    End Sub

    Private Sub studentNumberHolder_TextChanged(sender As Object, e As EventArgs) Handles studentNumberHolder.TextChanged
        ' Only proceed if the panel is currently in the visible position (530, 0)
        If studentQrCodePanel.Location <> New System.Drawing.Point(530, 0) Then
            Return
        End If

        Dim rawInput As String = studentNumberHolder.Text.Trim()

        ' Extract content inside [ ]
        Dim extractedStudentNumber As String = ExtractBetweenBrackets(rawInput)

        If Not String.IsNullOrEmpty(extractedStudentNumber) Then
            Dim student As StudentModel = GetStudentByNumber(extractedStudentNumber)

            If student IsNot Nothing Then
                CurrentStudent = student

                studentQrCodePanel.Location = New System.Drawing.Point(1000, 1000)
                sidePanel.Location = New System.Drawing.Point(1000, 1000)
                studentHomePanel.Location = New System.Drawing.Point(0, 0)

                Dim fullName As String = CurrentStudent.FirstName

                ' Add middle initial if it exists
                If Not String.IsNullOrWhiteSpace(CurrentStudent.MiddleInitial) Then
                    fullName &= " " & CurrentStudent.MiddleInitial & "."
                End If

                ' Add last name
                fullName &= " " & CurrentStudent.LastName

                ' Optionally add suffix
                If Not String.IsNullOrWhiteSpace(CurrentStudent.Suffix) Then
                    fullName &= " " & CurrentStudent.Suffix
                End If

                studentNameText.Text = fullName & "!"
                studentNumberText.Text = CurrentStudent.StudentNumber

                studentTimeInText.Text = "Time IN: " & DateTime.Now.ToString("hh:mm tt")

                ClearInputs(studentHomePanel)
                LoadProfessorsToComboBox()
                LoadReasonsToComboBox()
                studentNumberHolder.Clear()
                SetPlaceholder()
            Else
                studentNumberHolder.Clear()
                MessageBox.Show("Student number does not exist.", "Sign-In Failed")
            End If
        End If
    End Sub


    ' Helper function
    Private Function ExtractBetweenBrackets(input As String) As String
        Dim start As Integer = input.IndexOf("[")
        Dim endPos As Integer = input.IndexOf("]")

        If start <> -1 AndAlso endPos > start Then
            Return input.Substring(start + 1, endPos - start - 1)
        End If

        Return String.Empty
    End Function


    Private Sub arhiveStudentNumberInput_TextChanged(sender As Object, e As EventArgs) Handles archiveStudentNumberInput.TextChanged
        Dim txtBox As TextBox = CType(sender, TextBox)
        Dim cursorPosition As Integer = txtBox.SelectionStart

        ' Keep only numeric characters (remove letters, symbols, etc.)
        Dim digitsOnly As String = New String(txtBox.Text.Where(AddressOf Char.IsDigit).ToArray())

        ' Format: insert dash after 2 digits (e.g., 20-1234)
        If digitsOnly.Length > 2 Then
            txtBox.Text = digitsOnly.Substring(0, 2) & "-" & digitsOnly.Substring(2)
        Else
            txtBox.Text = digitsOnly
        End If

        ' Reset cursor position
        If cursorPosition <= 2 Then
            txtBox.SelectionStart = cursorPosition
        Else
            txtBox.SelectionStart = Math.Min(txtBox.Text.Length, cursorPosition + 1)
        End If
    End Sub
    Private Sub LoadSectionsIntoComboBox4()
        Dim dt As DataTable = GetAllSections()

        dt.Columns.Add("DisplayText", GetType(String))
        For Each row As DataRow In dt.Rows
            row("DisplayText") = row("section").ToString()
        Next

        archiveSectionBox.DataSource = dt
        archiveSectionBox.DisplayMember = "DisplayText"
        archiveSectionBox.ValueMember = "id"
        archiveSectionBox.SelectedIndex = -1
    End Sub
    Private selectedArchiveId As Integer = -1
    Private selectedArchiveIds As New List(Of Integer)()
    Private Sub archiveView_SelectionChanged(sender As Object, e As EventArgs) Handles archiveView.SelectionChanged
        ' Update the global list of selected archive IDs
        selectedArchiveIds = archiveView.SelectedRows _
        .Cast(Of DataGridViewRow)() _
        .Where(Function(r) Not r.IsNewRow) _
        .Select(Function(r) Convert.ToInt32(r.Cells("id").Value)) _
        .ToList()

        ' Optional: display IDs for debugging
        Console.WriteLine("Selected Archive IDs: " & String.Join(", ", selectedArchiveIds))
    End Sub

    Private Sub archiveView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles archiveView.CellContentClick
        If e.RowIndex >= 0 AndAlso e.RowIndex < archiveView.Rows.Count Then
            Dim row As DataGridViewRow = archiveView.Rows(e.RowIndex)

            ' Capture ID
            selectedArchiveId = Convert.ToInt32(row.Cells("id").Value)

            ' Assign values safely to inputs
            archiveStudentNumberInput.Text = If(IsDBNull(row.Cells("archiveStudentNumber").Value), "", row.Cells("archiveStudentNumber").Value.ToString())
            archiveFirstNameInput.Text = If(IsDBNull(row.Cells("archiveFirstName").Value), "", row.Cells("archiveFirstName").Value.ToString())
            archiveLastNameInput.Text = If(IsDBNull(row.Cells("archiveLastName").Value), "", row.Cells("archiveLastName").Value.ToString())
            archiveMiddleInitialInput.Text = If(IsDBNull(row.Cells("archiveMiddleInitial").Value), "", row.Cells("archiveMiddleInitial").Value.ToString())
            archiveEmailInput.Text = If(IsDBNull(row.Cells("archiveEmail").Value), "", row.Cells("archiveEmail").Value.ToString())
            archiveYearInput.Text = If(IsDBNull(row.Cells("archiveYear").Value), "", row.Cells("archiveYear").Value.ToString())

            ' Suffix
            Dim suffixValue As String = If(IsDBNull(row.Cells("archiveSuffix").Value), "", row.Cells("archiveSuffix").Value.ToString().Trim())
            If String.IsNullOrEmpty(suffixValue) OrElse Not archiveSuffixBox.Items.Contains(suffixValue) Then
                archiveSuffixBox.SelectedIndex = -1
            Else
                archiveSuffixBox.SelectedItem = suffixValue
            End If

            ' Section (matching section name to archiveSectionBox)
            Dim sectionName As String = If(IsDBNull(row.Cells("archiveSection").Value), "", row.Cells("archiveSection").Value.ToString().Trim())
            Dim matchedSectionIndex As Integer = -1

            For i As Integer = 0 To archiveSectionBox.Items.Count - 1
                If archiveSectionBox.GetItemText(archiveSectionBox.Items(i)).Trim().Equals(sectionName, StringComparison.OrdinalIgnoreCase) Then
                    matchedSectionIndex = i
                    Exit For
                End If
            Next

            archiveSectionBox.SelectedIndex = matchedSectionIndex

            ' Status (match manually like in student view)
            Dim statusValue As String = If(IsDBNull(row.Cells("archiveStatus").Value), "", row.Cells("archiveStatus").Value.ToString().Trim())
            Dim matchedStatusIndex As Integer = -1
            For i As Integer = 0 To archiveStatusBox.Items.Count - 1
                If archiveStatusBox.Items(i).ToString().Trim().ToLower() = statusValue.ToLower() Then
                    matchedStatusIndex = i
                    Exit For
                End If
            Next
            archiveStatusBox.SelectedIndex = matchedStatusIndex
        End If
    End Sub

    Private Sub unarchiveBtn_Click(sender As Object, e As EventArgs) Handles unarchiveBtn.Click
        ' Check if single student is selected or multiple students are selected
        Dim studentIds As New List(Of Integer)()

        ' Check for multiple selection in DataGridView
        If archiveView.SelectedRows.Count > 1 Then
            ' Multiple students selected
            For Each row As DataGridViewRow In archiveView.SelectedRows
                If row.Cells("id").Value IsNot Nothing Then
                    studentIds.Add(CInt(row.Cells("id").Value))
                End If
            Next
        ElseIf selectedArchiveId > 0 Then
            ' Single student selected (legacy method)
            studentIds.Add(selectedArchiveId)
        Else
            MessageBox.Show("Please select at least one student to unarchive.", "No Selection")
            Exit Sub
        End If

        ' Validate section for single student (for multiple, we'll use their existing sections)
        If studentIds.Count = 1 AndAlso archiveSectionBox.SelectedIndex = -1 Then
            MessageBox.Show("Please select a valid section.", "Validation Error")
            Exit Sub
        End If

        Dim message As String = If(studentIds.Count > 1,
                              $"Are you sure you want to unarchive {studentIds.Count} students?",
                              "Are you sure you want to unarchive this student?")

        If MessageBox.Show(message, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Exit Sub
        End If

        ' For single student, get the updated values from form
        ' For multiple students, we'll just set isGraduate = 0 without changing other fields
        Dim success As Boolean = False

        If studentIds.Count = 1 Then
            ' Get input values for single student
            Dim studentNumber As String = archiveStudentNumberInput.Text.Trim()
            Dim firstName As String = archiveFirstNameInput.Text.Trim()
            Dim lastName As String = archiveLastNameInput.Text.Trim()
            Dim middleInitial As String = archiveMiddleInitialInput.Text.Trim()
            Dim suffix As String = archiveSuffixBox.Text.Trim()
            Dim email As String = archiveEmailInput.Text.Trim()
            Dim status As String = archiveStatusBox.Text.Trim()
            Dim sectionId As Integer = Convert.ToInt32(archiveSectionBox.SelectedValue)

            ' Input validation for single student
            If Not System.Text.RegularExpressions.Regex.IsMatch(studentNumber, "^\d{2}-\d{5}$") OrElse
           String.IsNullOrWhiteSpace(firstName) OrElse
           String.IsNullOrWhiteSpace(lastName) OrElse
           String.IsNullOrWhiteSpace(email) OrElse
           Not IsValidName(firstName) OrElse
           Not IsValidName(lastName) OrElse
           (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
           (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
           Not IsValidEmail(email) Then

                MessageBox.Show("Please enter valid and complete information.", "Validation Error")
                Exit Sub
            End If

            ' Update single student with new values
            success = UnarchiveStudents(studentIds, firstName, lastName, middleInitial, suffix, email, sectionId, status)
        Else
            ' For multiple students, just set isGraduate = 0 without changing other fields
            success = UnarchiveStudents(studentIds)
        End If

        If success Then
            MessageBox.Show($"{studentIds.Count} student(s) successfully unarchived!", "Success")
            ClearInputs(adminArchivePanel)
            LoadArchives()
            selectedArchiveId = -1
        Else
            MessageBox.Show("Failed to unarchive student(s).", "Error")
        End If
    End Sub

    Private Sub archiveYearInput_TextChanged(sender As Object, e As EventArgs) Handles archiveYearInput.TextChanged
        Dim raw As String = archiveYearInput.Text
        Dim digitsOnly As String = New String(raw.Where(Function(c) Char.IsDigit(c)).ToArray())

        ' Limit to 8 digits total (for 2 full years)
        If digitsOnly.Length > 8 Then
            digitsOnly = digitsOnly.Substring(0, 8)
        End If

        ' Reformat to YYYY-YYYY if possible
        If digitsOnly.Length >= 5 Then
            archiveYearInput.Text = digitsOnly.Substring(0, 4) & "-" & digitsOnly.Substring(4)
        Else
            archiveYearInput.Text = digitsOnly
        End If

        ' Move the caret to the end so it doesn't jump back
        archiveYearInput.SelectionStart = archiveYearInput.Text.Length
    End Sub

    Private Sub archiveAddBtn_Click(sender As Object, e As EventArgs) Handles archiveAddBtn.Click
        ' Get input values
        Dim studentNumber As String = archiveStudentNumberInput.Text.Trim()
        Dim firstName As String = archiveFirstNameInput.Text.Trim()
        Dim lastName As String = archiveLastNameInput.Text.Trim()
        Dim middleInitial As String = archiveMiddleInitialInput.Text.Trim()
        Dim suffix As String = archiveSuffixBox.Text.Trim()
        Dim email As String = archiveEmailInput.Text.Trim()
        Dim status As String = archiveStatusBox.Text.Trim()
        Dim yearGraduated As String = archiveYearInput.Text.Trim()

        ' Validate section
        If archiveSectionBox.SelectedIndex = -1 Then
            MessageBox.Show("Please select a section.", "Validation Error")
            Exit Sub
        End If

        Dim sectionId As Integer = Convert.ToInt32(archiveSectionBox.SelectedValue)

        ' Input validation
        If Not System.Text.RegularExpressions.Regex.IsMatch(studentNumber, "^\d{2}-\d{5}$") OrElse
       String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) OrElse
       Not System.Text.RegularExpressions.Regex.IsMatch(yearGraduated, "^\d{4}-\d{4}$") Then

            MessageBox.Show("Please enter valid and complete information.", "Validation Error")
            Exit Sub
        End If

        ' Insert as graduated student
        Dim success As Boolean = InsertGraduatedStudent(studentNumber, firstName, lastName, middleInitial, suffix, email, sectionId, status, yearGraduated)

        If success Then
            MessageBox.Show("Graduated student added successfully!", "Success")
            ClearInputs(adminArchivePanel) ' Replace with your actual panel
            LoadArchives()
        Else
            MessageBox.Show("Failed to add graduated student.", "Error")
        End If
    End Sub
    Private Sub archiveUpdateBtn_Click(sender As Object, e As EventArgs) Handles archiveUpdateBtn.Click
        If selectedArchiveId <= 0 Then
            MessageBox.Show("Please select an archive entry to update.", "No Selection")
            Exit Sub
        End If

        ' Get values from form
        Dim studentNumber As String = archiveStudentNumberInput.Text.Trim()
        Dim firstName As String = archiveFirstNameInput.Text.Trim()
        Dim lastName As String = archiveLastNameInput.Text.Trim()
        Dim middleInitial As String = archiveMiddleInitialInput.Text.Trim()
        Dim suffix As String = archiveSuffixBox.Text.Trim()
        Dim email As String = archiveEmailInput.Text.Trim()
        Dim status As String = archiveStatusBox.Text.Trim()
        Dim yearGraduated As String = archiveYearInput.Text.Trim()

        If archiveSectionBox.SelectedIndex = -1 Then
            MessageBox.Show("Please select a section.", "Validation Error")
            Exit Sub
        End If

        Dim sectionId As Integer = Convert.ToInt32(archiveSectionBox.SelectedValue)

        ' Input validation
        If Not System.Text.RegularExpressions.Regex.IsMatch(studentNumber, "^\d{2}-\d{5}$") OrElse
       String.IsNullOrWhiteSpace(firstName) OrElse
       String.IsNullOrWhiteSpace(lastName) OrElse
       String.IsNullOrWhiteSpace(email) OrElse
       Not IsValidName(firstName) OrElse
       Not IsValidName(lastName) OrElse
       (Not String.IsNullOrWhiteSpace(middleInitial) AndAlso Not IsValidName(middleInitial)) OrElse
       (Not String.IsNullOrWhiteSpace(suffix) AndAlso Not IsValidName(suffix)) OrElse
       Not IsValidEmail(email) OrElse
       Not System.Text.RegularExpressions.Regex.IsMatch(yearGraduated, "^\d{4}-\d{4}$") Then

            MessageBox.Show("Please enter valid and complete information.", "Validation Error")
            Exit Sub
        End If

        ' Perform update
        Dim success As Boolean = UpdateGraduatedStudent(selectedArchiveId, studentNumber, firstName, lastName, middleInitial, suffix, email, sectionId, status, yearGraduated)

        If success Then
            MessageBox.Show("Graduated student successfully updated!", "Success")
            ClearInputs(adminArchivePanel)
            LoadArchives()
            selectedArchiveId = -1
        Else
            MessageBox.Show("Failed to update graduated student.", "Error")
        End If
    End Sub

    Public Function DeleteGraduatedStudent(id As Integer) As Boolean
        Try
            Connect()

            Dim query As String = "DELETE FROM students WHERE id = @id AND isGraduate = 1"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", id)
                Return cmd.ExecuteNonQuery() > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Delete graduated student failed: " & ex.Message, "DB Error")
            Return False
        Finally
            Disconnect()
        End Try
    End Function

    Private Sub archiveDeleteBtn_Click(sender As Object, e As EventArgs) Handles archiveDeleteBtn.Click
        If selectedArchiveId <= 0 Then
            MessageBox.Show("Please select an archived student to delete.", "No Selection")
            Exit Sub
        End If

        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this graduated student?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If result = DialogResult.Yes Then
            Dim success As Boolean = DeleteGraduatedStudent(selectedArchiveId)

            If success Then
                MessageBox.Show("Graduated student deleted successfully!", "Success")
                ClearInputs(adminArchivePanel)
                LoadArchives()
                selectedArchiveId = -1
            Else
                MessageBox.Show("Failed to delete graduated student.", "Error")
            End If
        End If
    End Sub

    Private Sub studentForgotBtn_Click(sender As Object, e As EventArgs) Handles studentForgotBtn.Click
        Dim email As String = studentForgotInput.Text.Trim()

        ' ✅ Validate email input
        If String.IsNullOrWhiteSpace(email) OrElse Not IsValidEmail(email) Then
            MessageBox.Show("Please enter a valid email address.", "Validation Error")
            Exit Sub
        End If

        ' ✅ Find the student by email
        Dim student As StudentModel = GetStudentByEmail(email)
        If student Is Nothing Then
            MessageBox.Show("No student found with that email.", "Not Found")
            Exit Sub
        End If

        Try
            Cursor.Current = Cursors.WaitCursor

            ' ✅ Generate QR Code with ECC Level H (High)
            Dim qrContent As String = $"[{student.StudentNumber}]"
            Dim qrGenerator As New QRCoder.QRCodeGenerator()
            Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
            Dim qrCode = New QRCoder.QRCode(qrData)
            Dim qrImage As Bitmap = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

            ' Use embedded resource (no need for path or File.Exists)
            Dim logo As Image = My.Resources.PLP

            ' Resize logo
            Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
            Dim resizedLogo As New Bitmap(logoSize, logoSize)
            Using g As Graphics = Graphics.FromImage(resizedLogo)
                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(logo, 0, 0, logoSize, logoSize)
            End Using


            ' ✅ Combine QR and Logo (with white background under logo)
            Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
            Using g As Graphics = Graphics.FromImage(qrWithLogo)
                g.Clear(Color.White)
                g.DrawImage(qrImage, 0, 0)

                ' Draw white rectangle "safe zone"
                Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                Dim padding As Integer = 6
                Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                g.FillRectangle(Brushes.White, whiteRect)

                ' Draw resized logo
                g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
            End Using

            ' ✅ Convert QR to stream
            Dim ms As New MemoryStream()
            qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
            ms.Position = 0

            ' ✅ Email setup
            Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
            Dim toAddress As New MailAddress(student.Email)
            Dim fromPassword As String = "swlqbwgztcqbneuw"

            Dim message As New MailMessage(fromAddress, toAddress)
            message.Subject = "Your Student QR Code"
            message.Body = $"Hi {student.FirstName}," & vbCrLf & vbCrLf &
               $"This is your assigned student number: {student.StudentNumber}" & vbCrLf & vbCrLf &
               "For your convenience, we’ve also provided a QR code as an alternative method to log in or authenticate yourself within the system." & vbCrLf &
               "You can scan it during login instead of typing your student number." & vbCrLf & vbCrLf &
               "If you did not request this or have questions, please contact the faculty immediately." & vbCrLf & vbCrLf &
               "Thank you," & vbCrLf

            message.Attachments.Add(New Attachment(ms, "StudentQRCode.png", "image/png"))

            Dim smtp As New SmtpClient("smtp.gmail.com", 587)
            smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
            smtp.EnableSsl = True
            smtp.Send(message)
            ClearInputs(studentForgotPanel)
            MessageBox.Show("Student number and QR code sent successfully.", "Email Sent")

        Catch ex As Exception
            MessageBox.Show("Failed to send email: " & ex.Message, "Email Error")
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub adminForgotBtn_Click(sender As Object, e As EventArgs) Handles adminForgotBtn.Click
        Dim email As String = adminForgotInput.Text.Trim()

        ' ✅ Validate email input
        If String.IsNullOrWhiteSpace(email) OrElse Not IsValidEmail(email) Then
            MessageBox.Show("Please enter a valid email address.", "Validation Error")
            Exit Sub
        End If

        ' ✅ Find admin
        Dim admin As AdminModel = GetAdminByEmail(email)
        If admin Is Nothing Then
            MessageBox.Show("No admin found with that email.", "Not Found")
            Exit Sub
        End If

        Try
            Cursor.Current = Cursors.WaitCursor

            ' ✅ Generate QR content: "username|password"
            Dim qrContent As String = $"[{admin.Username}][{admin.Password}]"


            ' ✅ Generate QR with ECC level H
            Dim qrGenerator As New QRCoder.QRCodeGenerator()
            Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
            Dim qrCode = New QRCoder.QRCode(qrData)
            Dim qrImage As Bitmap = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

            ' Use embedded resource (no need for path or File.Exists)
            Dim logo As Image = My.Resources.PLP

            ' Resize logo
            Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
            Dim resizedLogo As New Bitmap(logoSize, logoSize)
            Using g As Graphics = Graphics.FromImage(resizedLogo)
                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(logo, 0, 0, logoSize, logoSize)
            End Using


            ' ✅ Combine QR and logo
            Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
            Using g As Graphics = Graphics.FromImage(qrWithLogo)
                g.Clear(Color.White)
                g.DrawImage(qrImage, 0, 0)

                Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                Dim padding As Integer = 6
                Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                g.FillRectangle(Brushes.White, whiteRect)
                g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
            End Using

            ' ✅ Convert QR to stream
            Dim ms As New MemoryStream()
            qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
            ms.Position = 0

            ' ✅ Email setup
            Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
            Dim toAddress As New MailAddress(admin.Email)
            Dim fromPassword As String = "swlqbwgztcqbneuw"

            Dim message As New MailMessage(fromAddress, toAddress)
            message.Subject = "Your Admin Account Login QR Code"

            message.Body = $"Hi {admin.FirstName}," & vbCrLf & vbCrLf &
                           "Here are your login credentials:" & vbCrLf &
                           $"Username: {admin.Username}" & vbCrLf &
                           $"Password: {admin.Password}" & vbCrLf & vbCrLf &
                           "For your convenience, we’ve also attached a QR code with these credentials." & vbCrLf &
                           "You may use it as an alternative method to log in." & vbCrLf & vbCrLf &
                           "Please keep this information private and secure. If you did not request this, contact the faculty immediately." & vbCrLf & vbCrLf &
                           "Thank you," & vbCrLf

            message.Attachments.Add(New Attachment(ms, "AdminCredentialsQR.png", "image/png"))

            Dim smtp As New SmtpClient("smtp.gmail.com", 587)
            smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
            smtp.EnableSsl = True
            smtp.Send(message)

            ClearInputs(adminForgotPanel)
            MessageBox.Show("QR code and login credentials sent successfully.", "Email Sent")

        Catch ex As Exception
            MessageBox.Show("Failed to send email: " & ex.Message, "Email Error")
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub professorForgotBtn_Click(sender As Object, e As EventArgs) Handles professorForgotBtn.Click
        Dim email As String = professorForgotInput.Text.Trim()

        ' ✅ Validate input
        If String.IsNullOrWhiteSpace(email) OrElse Not IsValidEmail(email) Then
            MessageBox.Show("Please enter a valid email address.", "Validation Error")
            Exit Sub
        End If

        ' ✅ Get professor record
        Dim prof As ProfessorModel = GetProfessorByEmail(email)
        If prof Is Nothing Then
            MessageBox.Show("No professor found with that email.", "Not Found")
            Exit Sub
        End If

        Try
            Cursor.Current = Cursors.WaitCursor

            ' ✅ QR content in format [username][password]
            Dim qrContent As String = $"[{prof.Username}][{prof.Password}]"
            Dim qrGenerator As New QRCoder.QRCodeGenerator()
            Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCoder.QRCodeGenerator.ECCLevel.H)
            Dim qrCode = New QRCoder.QRCode(qrData)
            Dim qrImage As Bitmap = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

            ' ✅ Load and resize logo
            Dim originalLogo As Image = My.Resources.PLP  ' <- directly use the resource

            Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
            Dim resizedLogo As New Bitmap(logoSize, logoSize)
            Using g As Graphics = Graphics.FromImage(resizedLogo)
                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(originalLogo, 0, 0, logoSize, logoSize)
            End Using

            ' ✅ Draw QR with logo
            Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
            Using g As Graphics = Graphics.FromImage(qrWithLogo)
                g.Clear(Color.White)
                g.DrawImage(qrImage, 0, 0)

                Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                Dim padding As Integer = 6
                Dim whiteRect As New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2)
                g.FillRectangle(Brushes.White, whiteRect)
                g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
            End Using

            ' ✅ Convert image to stream
            Dim ms As New MemoryStream()
            qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
            ms.Position = 0

            ' ✅ Send email
            Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
            Dim toAddress As New MailAddress(prof.Email)
            Dim fromPassword As String = "swlqbwgztcqbneuw"

            Dim message As New MailMessage(fromAddress, toAddress)
            message.Subject = "Your Professor Login QR Code"
            message.Body = $"Hi {prof.FirstName}," & vbCrLf & vbCrLf &
                           "Here are your login credentials:" & vbCrLf &
                           $"Username: {prof.Username}" & vbCrLf &
                           $"Password: {prof.Password}" & vbCrLf & vbCrLf &
                           "We've also attached a QR code containing this information for easier login access." & vbCrLf &
                           "Please keep this secure and do not share with anyone." & vbCrLf & vbCrLf &
                           "Thank you," & vbCrLf

            message.Attachments.Add(New Attachment(ms, "ProfessorQRCode.png", "image/png"))

            Dim smtp As New SmtpClient("smtp.gmail.com", 587)
            smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
            smtp.EnableSsl = True
            smtp.Send(message)

            ClearInputs(professorForgotPanel)
            MessageBox.Show("Login credentials and QR sent to email.", "Email Sent")

        Catch ex As Exception
            MessageBox.Show("Failed to send email: " & ex.Message, "Email Error")
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub


    Private Sub adminUsernamePasswordHolder_TextChanged(sender As Object, e As EventArgs) Handles adminUsernamePasswordHolder.TextChanged
        ' Only process if the panel is active
        If adminQrCodePanel.Location <> New System.Drawing.Point(530, 0) Then
            Return
        End If

        Dim input As String = adminUsernamePasswordHolder.Text.Trim()

        ' ✅ Wait until the full QR is scanned — must end with ']'
        If Not input.EndsWith("]") Then Exit Sub

        ' ✅ Try match full format [username][password]
        Dim match As Match = Regex.Match(input, "^\[(.*?)\]\[(.*?)\]$")

        If match.Success Then
            Dim username As String = match.Groups(1).Value
            Dim password As String = match.Groups(2).Value

            ' ✅ Authenticate
            Dim admin As AdminModel = LoginAdmin(username, password)
            If admin IsNot Nothing Then
                CurrentAdmin = admin

                adminQrCodePanel.Location = New System.Drawing.Point(1000, 1000)
                sidePanel.Location = New System.Drawing.Point(1000, 1000)
                adminDashboard.Location = New System.Drawing.Point(0, 0)
                adminDashboardPanel.Location = New System.Drawing.Point(250, 0)

                LoadChart2()
                LoadCounts()
                LoadChart3()

                adminNameLabel.Text = CurrentAdmin.FirstName &
    If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".") &
    " " & CurrentAdmin.LastName &
    If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", ", " & CurrentAdmin.Suffix)

            Else
                MessageBox.Show("Invalid username or password.", "Login Failed")
            End If

            ' ✅ Clear the input for next scan
            adminUsernamePasswordHolder.Clear()
        End If
    End Sub

    Private Sub professorUsernamePasswordHolder_TextChanged(sender As Object, e As EventArgs) Handles professorUsernamePasswordHolder.TextChanged
        ' Only process if the panel is active
        If professorQrCodePanel.Location <> New System.Drawing.Point(530, 0) Then
            Return
        End If

        Dim input As String = professorUsernamePasswordHolder.Text.Trim()

        ' ✅ Wait until the full QR is scanned — must end with ']'
        If Not input.EndsWith("]") Then Exit Sub

        ' ✅ Try match full format [username][password]
        Dim match As Match = Regex.Match(input, "^\[(.*?)\]\[(.*?)\]$")

        If match.Success Then
            Dim username As String = match.Groups(1).Value
            Dim password As String = match.Groups(2).Value

            ' ✅ Authenticate
            Dim professor As ProfessorModel = LoginProfessor(username, password)
            If professor IsNot Nothing Then
                CurrentProfessor = professor
                professorQrCodePanel.Location = New System.Drawing.Point(1000, 1000)
                sidePanel.Location = New System.Drawing.Point(1000, 1000)
                professorHomePanel.Location = New System.Drawing.Point(0, 0)
                LoadSectionsIntoComboBox8()
                professorSurnameText.Text = "Welcome, Prof. " & CurrentProfessor.LastName & "!"
                professorDateText.Text = "Today's Date: " & DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")
                LoadProfessorCounts()
                LoadConsultations()

            Else
                MessageBox.Show("Invalid username or password.", "Login Failed")
            End If

            ' ✅ Clear the input for next scan
            professorUsernamePasswordHolder.Clear()
        End If
    End Sub


    Private Sub importBtn_Click(sender As Object, e As EventArgs) Handles importBtn.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Excel Files (*.xlsx)|*.xlsx"

        If ofd.ShowDialog() = DialogResult.OK Then
            ImportStudentsFromExcel(ofd.FileName)
        End If
    End Sub

    Private Sub PreviewStudentsFromValidSheets(filePath As String)
        Dim wb As New XLWorkbook(filePath)
        Dim sheetPattern As New Regex("^[A-Z]{2,}\s[0-9][A-Z]$", RegexOptions.IgnoreCase)
        Dim currentStatus As String = ""
        Dim statuses As String() = {"regular", "irregular"}

        For Each ws In wb.Worksheets
            If Not sheetPattern.IsMatch(ws.Name.Trim()) Then Continue For

            Debug.WriteLine("═════════════════════════════════════")
            Debug.WriteLine("📄 Reading sheet: " & ws.Name)

            Dim rowCount = If(ws.LastRowUsed()?.RowNumber(), 0)
            If rowCount = 0 Then
                Debug.WriteLine("⚠️ Sheet is empty: " & ws.Name)
                Continue For
            End If

            For row = 1 To rowCount
                Dim colA = ws.Cell(row, 1).GetValue(Of String)().Trim().ToLower()

                ' 🔍 Check for status only in column A
                If statuses.Any(Function(s) colA.Contains(s)) Then
                    currentStatus = Char.ToUpper(colA(0)) & colA.Substring(1).ToLower()
                    Debug.WriteLine($"🔎 Found {currentStatus} status at row {row}.")
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(currentStatus) Then Continue For

                ' 📥 Read student data only if currentStatus is set
                Dim studentNum = ws.Cell(row, 2).GetValue(Of String)().Trim()
                Dim studentName = ws.Cell(row, 3).GetValue(Of String)().Trim()
                Dim email = ws.Cell(row, 12).GetValue(Of String)().Trim()

                ' Skip blank rows or headers
                If String.IsNullOrWhiteSpace(studentNum) OrElse studentNum.ToLower().Contains("student no") Then Continue For

                Debug.WriteLine($"📥 [{currentStatus}] - {row} | {studentNum} | {studentName} | {email}")
            Next

            Debug.WriteLine("✅ Done with sheet: " & ws.Name)
        Next

        MessageBox.Show("Preview completed. See Output window.", "Done")
    End Sub

    Private Sub ImportStudentsFromExcel(filePath As String)
        Cursor.Current = Cursors.WaitCursor
        Dim wb = New XLWorkbook(filePath)
        Dim statuses As String() = {"regular", "irregular"}

        For Each ws In wb.Worksheets
            Dim sheetName = ws.Name.Trim()
            Dim rowCount = If(ws.LastRowUsed()?.RowNumber(), 0)
            If rowCount = 0 Then
                Debug.WriteLine($"⚠️ Skipping empty sheet: {sheetName}")
                Continue For
            End If

            Debug.WriteLine($"📄 Reading sheet: {sheetName}")
            Dim currentStatus As String = ""

            For row = 1 To rowCount
                Dim colA = ws.Cell(row, 1).GetValue(Of String)().Trim().ToLower()
                Debug.WriteLine(colA)
                If colA.Contains("irregular") Then
                    currentStatus = "irregular"
                    Debug.WriteLine($"🔎 Found irregular status at row {row}.")
                    Continue For
                ElseIf colA.Contains("regular") Then
                    currentStatus = "regular"
                    Debug.WriteLine($"🔎 Found regular status at row {row}.")
                    Continue For
                End If

                Dim studentnum = ws.Cell(row, 2).GetValue(Of String)().Trim()
                If String.IsNullOrWhiteSpace(studentnum) OrElse studentnum.ToLower().Contains("student") Then Continue For

                Dim fullName = ws.Cell(row, 3).GetValue(Of String)().Trim()
                Dim email = ws.Cell(row, 12).GetValue(Of String)().Trim()
                Dim sectionName = sheetName.Replace(" ", "") ' e.g., BSCS 2A => BSCS2A

                ' 🧠 Lookup section ID
                Dim sectionDt = ExecuteQuery("SELECT id FROM sections WHERE section = '" & MySqlHelper.EscapeString(sectionName) & "'")
                If sectionDt.Rows.Count = 0 Then
                    Debug.WriteLine("❌ Section not found: " & sectionName)
                    Continue For
                End If
                Dim sectionId = Convert.ToInt32(sectionDt.Rows(0)("id"))

                ' 🔤 Split name
                Dim lastName As String = ""
                Dim firstName As String = ""
                Dim middleInitial As String = ""
                Dim suffix As String = ""

                If fullName.Contains(",") Then
                    Dim nameParts = fullName.Split(","c)
                    Dim leftPart = nameParts(0).Trim() ' Lastname (with optional suffix)
                    Dim rightPart = nameParts(1).Trim() ' Firstname + Middle Initial

                    ' Split left part (LastName and optional Suffix)
                    Dim leftWords = leftPart.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)

                    If leftWords.Length >= 2 Then
                        lastName = leftWords(0).Trim()
                        suffix = String.Join(" ", leftWords.Skip(1)).Trim()
                    Else
                        lastName = leftPart
                    End If

                    ' Split right part (FirstName + Middle Initial)
                    Dim rightWords = rightPart.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)

                    If rightWords.Length >= 2 Then
                        Dim possibleMI = rightWords.Last().Trim()
                        If possibleMI.Length = 1 OrElse possibleMI.EndsWith(".") Then
                            middleInitial = possibleMI
                            firstName = String.Join(" ", rightWords.Take(rightWords.Length - 1)).Trim()
                        Else
                            firstName = rightPart
                        End If
                    Else
                        firstName = rightPart
                    End If
                End If

                Debug.WriteLine($"📝 Inserting: {studentnum} | Status: {currentStatus}")

                Dim query As String =
            "INSERT INTO students (student_number, first_name, last_name, middle_initial, suffix, email, section, status) " &
            "VALUES (@studentnum, @first_name, @last_name, @middle_initial, @suffix, @email, @sectionId, @status)"

                Try
                    Connect()
                    Using cmd As New MySqlCommand(query, conn)
                        cmd.Parameters.AddWithValue("@studentnum", studentnum)
                        cmd.Parameters.AddWithValue("@first_name", firstName)
                        cmd.Parameters.AddWithValue("@last_name", lastName)
                        cmd.Parameters.AddWithValue("@middle_initial", If(String.IsNullOrWhiteSpace(middleInitial), DBNull.Value, middleInitial))
                        cmd.Parameters.AddWithValue("@suffix", If(String.IsNullOrWhiteSpace(suffix), DBNull.Value, suffix))
                        cmd.Parameters.AddWithValue("@email", email)
                        cmd.Parameters.AddWithValue("@sectionId", sectionId)
                        cmd.Parameters.AddWithValue("@status", currentStatus)
                        cmd.ExecuteNonQuery()
                    End Using

                    ' ✅ CALL SEND QR CODE HERE - AFTER SUCCESSFUL INSERT
                    If Not String.IsNullOrWhiteSpace(email) Then

                        Debug.WriteLine($"📧 QR code sent to: {email}")
                    Else
                        Debug.WriteLine($"⚠️ No email for {studentnum}, skipping QR send")
                    End If

                    Debug.WriteLine($"📥 [{currentStatus}] - {row} | {studentnum} | {fullName} | {email}")
                Catch ex As MySqlException
                    MessageBox.Show("Error inserting student '" & studentnum & "': " & ex.Message)
                Finally
                    Disconnect() ' Ensure connection is closed
                End Try
            Next

            Debug.WriteLine($"✅ Done with sheet: {sheetName}")
            Debug.WriteLine("═══════════════════════════════════════")
        Next

        Cursor.Current = Cursors.Default
        MessageBox.Show("Import and QR code distribution completed.")
        LoadFilteredStudents()
    End Sub


    Private Sub SendStudentQRCode(studentnum As String, email As String, firstName As String)
        Try
            Dim qrContent As String = $"[{studentnum}]"

            ' ✅ Generate QR code
            Dim qrGenerator As New QRCodeGenerator()
            Dim qrData = qrGenerator.CreateQrCode(qrContent, QRCodeGenerator.ECCLevel.H)
            Dim qrCode = New QRCode(qrData)
            Dim qrImage = qrCode.GetGraphic(20, Color.Black, Color.White, drawQuietZones:=True)

            ' Use embedded resource (no need for path or File.Exists)
            Dim logo As Image = My.Resources.PLP

            ' Resize logo
            Dim logoSize As Integer = CInt(qrImage.Width * 0.15)
            Dim resizedLogo As New Bitmap(logoSize, logoSize)
            Using g As Graphics = Graphics.FromImage(resizedLogo)
                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(logo, 0, 0, logoSize, logoSize)
            End Using


            Dim qrWithLogo As New Bitmap(qrImage.Width, qrImage.Height)
            Using g As Graphics = Graphics.FromImage(qrWithLogo)
                g.Clear(Color.White)
                g.DrawImage(qrImage, 0, 0)

                ' White background for logo
                Dim centerX As Integer = (qrImage.Width - logoSize) \ 2
                Dim centerY As Integer = (qrImage.Height - logoSize) \ 2
                Dim padding As Integer = 6
                g.FillRectangle(Brushes.White, New Rectangle(centerX - padding, centerY - padding, logoSize + padding * 2, logoSize + padding * 2))

                g.DrawImage(resizedLogo, centerX, centerY, logoSize, logoSize)
            End Using

            ' ✅ Save to stream
            Dim ms As New MemoryStream()
            qrWithLogo.Save(ms, Imaging.ImageFormat.Png)
            ms.Position = 0

            ' ✅ Send email
            Dim fromAddress As New MailAddress("alvarez_juanito@plpasig.edu.ph", "Jhun Alvarez")
            Dim toAddress As New MailAddress(email)
            Dim fromPassword As String = "swlqbwgztcqbneuw"

            Dim message As New MailMessage(fromAddress, toAddress)
            message.Subject = "Your Student QR Code"
            message.Body = $"Hi {firstName}," & vbCrLf & vbCrLf &
                   "Attached is your student QR code." & vbCrLf &
                   "You may use it for login or authentication." & vbCrLf
            message.Attachments.Add(New Attachment(ms, "StudentQR.png", "image/png"))

            Dim smtp As New SmtpClient("smtp.gmail.com", 587)
            smtp.Credentials = New NetworkCredential(fromAddress.Address, fromPassword)
            smtp.EnableSsl = True
            smtp.Send(message)

            ' Optional: Log or show feedback
            Debug.WriteLine("QR sent to: " & email)

        Catch ex As Exception
            MessageBox.Show("QR email failed for " & email & ": " & ex.Message)
        End Try
    End Sub

    Private Sub changeLogoBtn_Click(sender As Object, e As EventArgs) Handles changeLogoBtn.Click
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif"
            ofd.Title = "Select a Logo Image"

            If ofd.ShowDialog() = DialogResult.OK Then
                Try
                    logoBox.Image = Image.FromFile(ofd.FileName)
                    logoBox.Tag = ofd.FileName

                    ' 🔐 Save the image path to app settings
                    My.Settings.LogoPath = ofd.FileName
                    My.Settings.Save()

                Catch ex As Exception
                    MessageBox.Show("Failed to load image: " & ex.Message)
                End Try
            End If
        End Using
    End Sub

    Private Sub changeDepartmentBtn_Click(sender As Object, e As EventArgs) Handles changeDepartmentBtn.Click
        Dim deptName As String = departmentNameInput.Text.Trim()

        ' Save to user-scoped setting
        My.Settings.DepartmentName = deptName
        My.Settings.Save()

        MessageBox.Show("Department name saved successfully.")
    End Sub
    Private Sub LoadDepartmentName()
        Dim savedDeptName As String = My.Settings.DepartmentName
        departmentNameInput.Text = savedDeptName
    End Sub

    Private Sub LoadReasonsToFirstThirdReasonBox()
        Dim dt As DataTable = GetAllReasons()

        ' Add a custom "Empty" row
        Dim emptyRow As DataRow = dt.NewRow()
        emptyRow("id") = DBNull.Value  ' Use 0 if your DB doesn't allow nulls
        emptyRow("reason") = "Empty"

        ' Insert it at the top
        dt.Rows.InsertAt(emptyRow, 0)

        ' Bind to firstThirdReasonBox
        firstThirdReasonBox.DataSource = dt
        firstThirdReasonBox.DisplayMember = "reason"
        firstThirdReasonBox.ValueMember = "id"
        firstThirdReasonBox.SelectedIndex = 0  ' Select "Empty" by default
    End Sub
    Private Sub LoadReasonsToFourthReasonBox()
        Dim dt As DataTable = GetAllReasons()

        ' Add a custom "Empty" row
        Dim emptyRow As DataRow = dt.NewRow()
        emptyRow("id") = DBNull.Value  ' Use 0 if your DB doesn't allow nulls
        emptyRow("reason") = "Empty"

        ' Insert it at the top
        dt.Rows.InsertAt(emptyRow, 0)

        ' Bind to firstThirdReasonBox
        fourthReasonBox.DataSource = dt
        fourthReasonBox.DisplayMember = "reason"
        fourthReasonBox.ValueMember = "id"
        fourthReasonBox.SelectedIndex = 0  ' Select "Empty" by default
    End Sub
    Private Sub firstThirdReasonBtn_Click(sender As Object, e As EventArgs) Handles firstThirdReasonBtn.Click
        If firstThirdReasonBox.SelectedItem IsNot Nothing Then
            Dim selectedReason As DataRowView = CType(firstThirdReasonBox.SelectedItem, DataRowView)
            Dim reasonText As String = selectedReason("reason").ToString()

            ' Save to setting
            My.Settings.firstThirdReason = reasonText
            My.Settings.Save()

            MessageBox.Show("Reason saved to settings.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Please select a reason first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub fourthReasonBtn_Click(sender As Object, e As EventArgs) Handles fourthReasonBtn.Click
        If fourthReasonBox.SelectedItem IsNot Nothing Then
            Dim selectedReason As DataRowView = CType(fourthReasonBox.SelectedItem, DataRowView)
            Dim reasonText As String = selectedReason("reason").ToString()

            ' Save to setting
            My.Settings.fourthReason = reasonText
            My.Settings.Save()

            MessageBox.Show("Reason saved to settings.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Please select a reason first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub changeTitleBtn_Click(sender As Object, e As EventArgs) Handles changeTitleBtn.Click
        Dim newTitle As String = titleInput.Text.Trim()
        My.Settings.systemTitle = newTitle
        My.Settings.Save()
        MessageBox.Show("System title saved successfully.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Private Sub LoadSystemTitle()
        Dim savedTitle As String = My.Settings.systemTitle
        titleInput.Text = savedTitle

        titleText.Text = savedTitle
    End Sub

    Private Sub consultView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles consultView.CellContentClick

    End Sub

    Private Sub adminGenerateBtn_Click(sender As Object, e As EventArgs) Handles generateReportForms.Click
        Dim selectedYear As Integer
        If Not Integer.TryParse(dashboardYear.Text, selectedYear) Then
            MessageBox.Show("Please select a valid year from the dashboard.", "Invalid Year")
            Return
        End If

        Dim fromDate As New Date(selectedYear, 1, 1)
        Dim toDate As New Date(selectedYear, 12, 31)

        Dim totalConsultations As Integer = 0
        Dim uniqueStudents As Integer = 0
        Dim topProfessors As New Dictionary(Of String, Integer)
        Dim professorSummaries As New List(Of ProfessorSummary)
        Dim monthlyCounts As New Dictionary(Of Integer, Integer) ' Month number (1-12) to count

        Try
            Connect()

            ' Total and unique students for the year
            Dim query1 As String = "
SELECT COUNT(*) AS total, COUNT(DISTINCT student_id) AS unique_students
FROM reports
WHERE YEAR(consultation_date) = @year"
            Using cmd As New MySqlCommand(query1, conn)
                cmd.Parameters.AddWithValue("@year", selectedYear)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        totalConsultations = Convert.ToInt32(reader("total"))
                        uniqueStudents = Convert.ToInt32(reader("unique_students"))
                    End If
                End Using
            End Using

            ' Get monthly counts for the line chart
            Dim queryMonthly As String = "
SELECT MONTH(consultation_date) AS month, COUNT(*) AS count
FROM reports
WHERE YEAR(consultation_date) = @year
GROUP BY MONTH(consultation_date)
ORDER BY month"
            Using cmd As New MySqlCommand(queryMonthly, conn)
                cmd.Parameters.AddWithValue("@year", selectedYear)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim month As Integer = Convert.ToInt32(reader("month"))
                        Dim count As Integer = Convert.ToInt32(reader("count"))
                        monthlyCounts(month) = count
                    End While
                End Using
            End Using

            ' Top professors (and store their IDs)
            Dim query2 As String = "
SELECT p.id, CONCAT(p.first_name, ' ', p.last_name) AS prof_name, COUNT(*) AS count
FROM reports r
JOIN professors p ON r.professor_id = p.id
WHERE YEAR(r.consultation_date) = @year
GROUP BY r.professor_id
ORDER BY count DESC"
            Using cmd As New MySqlCommand(query2, conn)
                cmd.Parameters.AddWithValue("@year", selectedYear)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim name As String = reader("prof_name").ToString()
                        Dim count As Integer = Convert.ToInt32(reader("count"))
                        topProfessors(name) = count

                        professorSummaries.Add(New ProfessorSummary With {
                .Id = Convert.ToInt32(reader("id")),
                .Name = name,
                .TotalConsultations = count
            })
                    End While
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Failed to retrieve data: " & ex.Message, "DB Error")
            Return
        Finally
            Disconnect()
        End Try

        Dim confirmResult As DialogResult = MessageBox.Show(
$"A report will be generated for the year {selectedYear}.{Environment.NewLine}Do you want to continue?",
"Confirm Report Generation",
MessageBoxButtons.YesNo,
MessageBoxIcon.Question
)

        If confirmResult <> DialogResult.Yes Then Exit Sub

        ' Save dialog
        Dim sfd As New SaveFileDialog()
        sfd.Filter = "PDF Files (*.pdf)|*.pdf"
        sfd.FileName = $"AdminSummary_{selectedYear}.pdf"
        If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

        ' Create PDF
        Dim doc As New PdfDocument()
        doc.Info.Title = "Consultation Report"

        Dim fontTitle As New XFont("Segoe UI", 16, XFontStyle.Bold)
        Dim fontSubTitle As New XFont("Segoe UI", 9, XFontStyle.Regular)
        Dim fontHeader As New XFont("Segoe UI", 9, XFontStyle.Bold)
        Dim fontValue As New XFont("Segoe UI", 14, XFontStyle.Bold)
        Dim fontSmall As New XFont("Segoe UI", 8, XFontStyle.Regular)
        Dim fontAxis As New XFont("Segoe UI", 7, XFontStyle.Regular)

        Dim leftLogoPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Resources\plplogo.png")
        Dim leftLogo As XImage = If(File.Exists(leftLogoPath), XImage.FromFile(leftLogoPath), Nothing)

        Dim rightLogoPath As String = My.Settings.LogoPath
        Dim rightLogo As XImage = If(File.Exists(rightLogoPath), XImage.FromFile(rightLogoPath), Nothing)

        Dim DrawHeader = Sub(hdrGfx As XGraphics, hdrPage As PdfPage)
                             ' Background rectangle
                             hdrGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(245, 245, 245)), 0, 0, hdrPage.Width.Point, 80)

                             ' Top-left text: small, single line
                             Dim xLeft As Double = 20
                             Dim yTop As Double = 5  ' small margin from top
                             Dim smallFont As New XFont("Segoe UI", 7, XFontStyle.Regular)

                             Dim adminName As String = "Administrator"
                             If CurrentAdmin IsNot Nothing AndAlso Not String.IsNullOrEmpty(CurrentAdmin.LastName) Then
                                 adminName = CurrentAdmin.LastName.ToUpper() ' optional uppercase
                             End If

                             Dim mi As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".")
                             Dim suffix As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", " " & CurrentAdmin.Suffix)
                             Dim fullName As String = $"{CurrentAdmin.FirstName}{mi} {CurrentAdmin.LastName}{suffix}".Trim()

                             Dim generatedText As String = $"Report generated by Admin Prof. {fullName} as of {DateTime.Now:MMMM dd, yyyy, 'at' hh:mm tt}"

                             hdrGfx.DrawString(generatedText, smallFont, XBrushes.Black, New XRect(xLeft, yTop, 400, 10), XStringFormats.TopLeft)

                             Dim targetHeight As Double = 50

                             If leftLogo IsNot Nothing Then
                                 Dim aspectLeft As Double = leftLogo.PixelWidth / leftLogo.PixelHeight
                                 Dim newWidthLeft As Double = targetHeight * aspectLeft
                                 hdrGfx.DrawImage(leftLogo, xLeft, yTop + 15, newWidthLeft, targetHeight)
                             End If

                             If rightLogo IsNot Nothing Then
                                 Dim aspectRight As Double = rightLogo.PixelWidth / rightLogo.PixelHeight
                                 Dim newWidthRight As Double = targetHeight * aspectRight
                                 hdrGfx.DrawImage(rightLogo, hdrPage.Width.Point - newWidthRight - 20, yTop + 15, newWidthRight, targetHeight)
                             End If


                             ' Centered text below logos
                             Dim centerX As Double = hdrPage.Width.Point / 2
                             Dim yCenter As Double = yTop + 20  ' adjust vertical position below small top text

                             hdrGfx.DrawString("Annual Report", fontTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter, 300, 20), XStringFormats.TopCenter)
                             hdrGfx.DrawString("Pamantasan ng Lungsod ng Pasig", fontSubTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter + 20, 300, 15), XStringFormats.TopCenter)
                             hdrGfx.DrawString(My.Settings.DepartmentName, fontSubTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter + 35, 300, 15), XStringFormats.TopCenter)
                         End Sub


        ' === Page 1 ===
        Dim page As PdfPage = doc.AddPage()
        page.Size = PageSize.A4
        Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
        DrawHeader(gfx, page)

        ' === Summary Boxes ===
        Dim cardY As Double = 90
        Dim cardWidth As Double = 220
        Dim cardHeight As Double = 50
        Dim cardSpacing As Double = 30

        ' First card (Total Consultations)
        gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(230, 248, 255)), 30, cardY, cardWidth, cardHeight)
        gfx.DrawRectangle(XPens.DeepSkyBlue, 30, cardY, cardWidth, cardHeight)
        gfx.DrawString("Total Consultations", fontHeader, XBrushes.Black, New XPoint(40, cardY + 15))
        gfx.DrawString(totalConsultations.ToString(), fontValue, XBrushes.Black, New XPoint(40, cardY + 38))

        ' Second card (Unique Students)
        gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(230, 255, 240)), 30 + cardWidth + cardSpacing, cardY, cardWidth, cardHeight)
        gfx.DrawRectangle(XPens.MediumSeaGreen, 30 + cardWidth + cardSpacing, cardY, cardWidth, cardHeight)
        gfx.DrawString("Unique Students", fontHeader, XBrushes.Black, New XPoint(40 + cardWidth + cardSpacing, cardY + 15))
        gfx.DrawString(uniqueStudents.ToString(), fontValue, XBrushes.Black, New XPoint(40 + cardWidth + cardSpacing, cardY + 38))

        ' === Pie Chart Section ===
        Dim pieY As Double = cardY + cardHeight + 30
        Dim pieLeftMargin As Double = 30
        Dim pieRadius As Double = 40
        Dim pieCenterY As Double = pieY + 60
        Dim pieCenterX As Double = pieLeftMargin + pieRadius + 20

        ' Get report type distribution data
        Dim specialCount As Integer = 0
        Dim regularCount As Integer = 0

        Try
            Connect()
            Dim querySpecial As String = "
SELECT 
    SUM(CASE WHEN rs.is_special = 1 THEN 1 ELSE 0 END) AS special,
    SUM(CASE WHEN rs.is_special = 0 THEN 1 ELSE 0 END) AS regular
FROM reports r
JOIN reasons rs ON r.reason_id = rs.id
WHERE YEAR(r.consultation_date) = @year"

            Using cmd As New MySqlCommand(querySpecial, conn)
                cmd.Parameters.AddWithValue("@year", selectedYear)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        specialCount = If(IsDBNull(reader("special")), 0, Convert.ToInt32(reader("special")))
                        regularCount = If(IsDBNull(reader("regular")), 0, Convert.ToInt32(reader("regular")))
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error getting report distribution: " & ex.Message)
        Finally
            Disconnect()
        End Try

        ' Draw pie chart title
        gfx.DrawString("Report Type Distribution", fontHeader, XBrushes.Black,
           New XPoint(pieLeftMargin, pieY))

        If (specialCount + regularCount) > 0 Then
            ' Calculate angles and draw pie chart (same as before)
            Dim specialAngle As Double = (specialCount / (specialCount + regularCount)) * 360
            Dim regularAngle As Double = 360 - specialAngle

            gfx.DrawPie(New XSolidBrush(XColor.FromArgb(255, 99, 132)), pieCenterX - pieRadius, pieCenterY - pieRadius,
            pieRadius * 2, pieRadius * 2, 0, specialAngle)
            gfx.DrawPie(New XSolidBrush(XColor.FromArgb(54, 162, 235)), pieCenterX - pieRadius, pieCenterY - pieRadius,
            pieRadius * 2, pieRadius * 2, specialAngle, regularAngle)

            ' Draw legend
            Dim legendX As Double = pieCenterX + pieRadius + 20
            Dim legendY As Double = pieCenterY - 20

            gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(255, 99, 132)), legendX, legendY, 12, 12)
            gfx.DrawString($"Special Events: {specialCount} ({Math.Round((specialCount / (specialCount + regularCount)) * 100)}%)",
               fontSmall, XBrushes.Black, New XPoint(legendX + 15, legendY + 10))

            gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(54, 162, 235)), legendX, legendY + 25, 12, 12)
            gfx.DrawString($"Regular Consultations: {regularCount} ({Math.Round((regularCount / (specialCount + regularCount)) * 100)}%)",
               fontSmall, XBrushes.Black, New XPoint(legendX + 15, legendY + 35))

            ' Add percentage labels
            Dim labelFont As New XFont("Segoe UI", 7, XFontStyle.Bold)
            Dim specialPercent As Double = Math.Round((specialCount / (specialCount + regularCount)) * 100)
            Dim regularPercent As Double = 100 - specialPercent

            Dim specialMidAngle As Double = specialAngle / 2
            Dim specialLabelX As Double = pieCenterX + (pieRadius * 0.6) * Math.Cos(specialMidAngle * Math.PI / 180)
            Dim specialLabelY As Double = pieCenterY + (pieRadius * 0.6) * Math.Sin(specialMidAngle * Math.PI / 180)
            gfx.DrawString($"{specialPercent}%", labelFont, XBrushes.White,
               New XPoint(specialLabelX, specialLabelY), XStringFormats.Center)

            Dim regularMidAngle As Double = specialAngle + (regularAngle / 2)
            Dim regularLabelX As Double = pieCenterX + (pieRadius * 0.6) * Math.Cos(regularMidAngle * Math.PI / 180)
            Dim regularLabelY As Double = pieCenterY + (pieRadius * 0.6) * Math.Sin(regularMidAngle * Math.PI / 180)
            gfx.DrawString($"{regularPercent}%", labelFont, XBrushes.White,
               New XPoint(regularLabelX, regularLabelY), XStringFormats.Center)
        Else
            gfx.DrawString("No consultation data available", fontSmall, XBrushes.Gray,
               New XPoint(pieLeftMargin + 100, pieCenterY), XStringFormats.Center)
        End If

        ' === Monthly Line Chart Implementation ===
        Dim lineChartY As Double = pieCenterY + pieRadius + 40
        Dim lineChartHeight As Double = 120
        Dim lineChartWidth As Double = 500
        Dim lineChartLeft As Double = 40
        Dim lineChartRight As Double = lineChartLeft + lineChartWidth

        ' Prepare monthly data for all 12 months
        Dim completeMonthlyCounts As New Dictionary(Of Integer, Integer)()
        For month As Integer = 1 To 12
            completeMonthlyCounts(month) = If(monthlyCounts.ContainsKey(month), monthlyCounts(month), 0)
        Next

        ' Calculate chart parameters
        Dim maxMonthlyCount As Integer = If(completeMonthlyCounts.Values.Any(), completeMonthlyCounts.Values.Max(), 1)
        If maxMonthlyCount = 0 Then maxMonthlyCount = 1

        ' Draw main title
        gfx.DrawString("Monthly Consultation Counts", fontHeader, XBrushes.Black,
           New XPoint(lineChartLeft, lineChartY - 20))
        lineChartY += 10

        ' Draw chart area
        gfx.DrawRectangle(XPens.LightGray, lineChartLeft, lineChartY, lineChartWidth, lineChartHeight)
        gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(250, 250, 250)), lineChartLeft, lineChartY, lineChartWidth, lineChartHeight)

        ' Calculate X spacing
        Dim pointGap As Double = lineChartWidth / 11 ' 12 months = 11 gaps

        ' Draw Y-axis grid and labels
        Dim yStep As Integer = Math.Max(1, CInt(Math.Ceiling(maxMonthlyCount / 5)))
        For i As Integer = 0 To maxMonthlyCount Step yStep
            Dim yPos As Double = lineChartY + lineChartHeight - (i / maxMonthlyCount * lineChartHeight)
            gfx.DrawLine(XPens.LightGray, lineChartLeft, yPos, lineChartRight, yPos)
            gfx.DrawString(i.ToString(), fontAxis, XBrushes.Black,
              New XPoint(lineChartLeft - 5, yPos - 5), XStringFormats.TopRight)
        Next

        ' Draw axes
        gfx.DrawLine(XPens.Black, lineChartLeft, lineChartY, lineChartLeft, lineChartY + lineChartHeight)
        gfx.DrawLine(XPens.Black, lineChartLeft, lineChartY + lineChartHeight, lineChartRight, lineChartY + lineChartHeight)

        ' Draw data points and connecting lines
        Dim prevPoint As XPoint? = Nothing
        For month As Integer = 1 To 12
            Dim count = completeMonthlyCounts(month)
            Dim xPos = lineChartLeft + ((month - 1) * pointGap)
            Dim yPos = lineChartY + lineChartHeight - (count / maxMonthlyCount * lineChartHeight)
            Dim currentPoint = New XPoint(xPos, yPos)

            ' Draw data point
            gfx.DrawEllipse(XBrushes.DodgerBlue, xPos - 2, yPos - 2, 4, 4)

            ' Draw connecting line
            If prevPoint.HasValue Then
                gfx.DrawLine(New XPen(XColor.FromArgb(30, 144, 255), 2), prevPoint.Value, currentPoint)
            End If
            prevPoint = currentPoint

            ' Draw month label
            Dim monthName As String = New Date(selectedYear, month, 1).ToString("MMM")
            gfx.DrawString(monthName, fontAxis, XBrushes.Black,
             New XRect(xPos - 15, lineChartY + lineChartHeight + 5, 30, 10),
             XStringFormats.TopCenter)
        Next

        ' === Bar Chart: Professors by Count ===
        Dim barYStart As Double = lineChartY + lineChartHeight + 40
        gfx.DrawString("Professors by Consultation Count", fontHeader, XBrushes.Black, New XPoint(30, barYStart - 20))

        Dim maxVal As Integer = If(topProfessors.Count > 0, topProfessors.Values.Max(), 1)
        Dim currentY As Double = barYStart
        Dim colorIndex As Integer = 0
        For Each kvp In topProfessors
            If currentY + 30 > page.Height.Point Then
                page = doc.AddPage()
                page.Size = PageSize.A4
                gfx = XGraphics.FromPdfPage(page)
                DrawHeader(gfx, page)
                currentY = 90
            End If

            Dim barWidth As Double = (kvp.Value / maxVal) * 300
            Dim brush As XBrush = GetColorByIndex(colorIndex)

            gfx.DrawString(kvp.Key, fontSmall, XBrushes.Black, New XPoint(30, currentY + 10))
            gfx.DrawRectangle(brush, 150, currentY, barWidth, 14)
            gfx.DrawString(kvp.Value.ToString(), fontSmall, XBrushes.Black, New XPoint(160 + barWidth, currentY + 10))

            currentY += 24
            colorIndex += 1
        Next

        ' === For each professor: Summary + Pie Chart ===
        Dim yPosition As Double = currentY + 20
        For Each prof In professorSummaries
            If yPosition + 160 > page.Height.Point Then
                page = doc.AddPage()
                page.Size = PageSize.A4
                gfx = XGraphics.FromPdfPage(page)
                DrawHeader(gfx, page)
                yPosition = 90
            End If

            gfx.DrawString("Professor: " & prof.Name, fontHeader, XBrushes.Black, New XPoint(30, yPosition))
            gfx.DrawString("Total Consultations: " & prof.TotalConsultations, fontSmall, XBrushes.Black, New XPoint(30, yPosition + 20))

            ' Query reasons per professor for the year
            Dim reasonCounts As New Dictionary(Of String, Integer)
            Try
                Connect()
                Using cmd As New MySqlCommand("
        SELECT rs.reason, COUNT(*) AS count
        FROM reports r
        JOIN reasons rs ON r.reason_id = rs.id
        WHERE r.professor_id = @pid AND YEAR(r.consultation_date) = @year
        GROUP BY rs.reason", conn)
                    cmd.Parameters.AddWithValue("@pid", prof.Id)
                    cmd.Parameters.AddWithValue("@year", selectedYear)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            reasonCounts(reader("reason").ToString()) = Convert.ToInt32(reader("count"))
                        End While
                    End Using
                End Using
            Catch ex As Exception
                ' Handle error
            Finally
                Disconnect()
            End Try

            ' Draw Pie Chart
            If reasonCounts.Count > 0 Then
                Dim centerX As Double = 300
                Dim centerY As Double = yPosition + 70
                Dim radius As Double = 40
                Dim totalReasons As Integer = reasonCounts.Values.Sum()
                Dim startAngle As Double = 0
                Dim pieColorIndex As Integer = 0

                For Each kvp In reasonCounts
                    Dim sweepAngle As Double = (kvp.Value / totalReasons) * 360
                    Dim brush As XBrush = GetColorByIndex(pieColorIndex)
                    gfx.DrawPie(brush, centerX - radius, centerY - radius, radius * 2, radius * 2, startAngle, sweepAngle)
                    startAngle += sweepAngle
                    pieColorIndex += 1
                Next

                ' Legend
                Dim legendY As Double = yPosition + 30
                pieColorIndex = 0
                For Each kvp In reasonCounts
                    Dim brush As XBrush = GetColorByIndex(pieColorIndex)
                    gfx.DrawRectangle(brush, centerX + radius + 20, legendY, 10, 10)
                    gfx.DrawString($"{kvp.Key} ({kvp.Value})", fontSmall, XBrushes.Black, New XPoint(centerX + radius + 35, legendY + 8))
                    legendY += 14
                    pieColorIndex += 1
                Next
            Else
                gfx.DrawString("No reasons available.", fontSmall, XBrushes.Gray, New XPoint(300, yPosition + 20))
            End If

            yPosition += 180
        Next

        doc.Save(sfd.FileName)
        MessageBox.Show($"Admin summary PDF for {selectedYear} generated successfully!", "Success")
    End Sub

    Private Class ProfessorSummary
        Public Property Id As Integer
        Public Property Name As String
        Public Property TotalConsultations As Integer
    End Class

    Private Sub IconButton1_Click(sender As Object, e As EventArgs) Handles IconButton1.Click
        Dim currentYear As Integer
        If Integer.TryParse(dashboardYear.Text.Trim(), currentYear) Then
            currentYear -= 1
            dashboardYear.Text = currentYear.ToString()
            LoadDashboardCharts()
        Else
            MsgBox("Invalid year: " & dashboardYear.Text)
        End If
    End Sub

    Private Sub IconButton2_Click(sender As Object, e As EventArgs) Handles IconButton2.Click
        Dim currentYear As Integer
        If Integer.TryParse(dashboardYear.Text.Trim(), currentYear) Then
            currentYear += 1
            dashboardYear.Text = currentYear.ToString()
            LoadDashboardCharts()
        Else
            MsgBox("Invalid year: " & dashboardYear.Text)
        End If
    End Sub

    Private Sub LoadDashboardCharts()
        LoadChart3() ' Area chart (Consultations per day in selected year)
        LoadChart2() ' Optional: any other chart like top sections or reasons
    End Sub
    Private Sub IconPictureBox4_Click(sender As Object, e As EventArgs) Handles IconPictureBox4.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit the application?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = DialogResult.Yes Then
            Me.Close()

        End If
    End Sub

    Private Sub IconPictureBox5_Click(sender As Object, e As EventArgs) Handles IconPictureBox5.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit the application?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = DialogResult.Yes Then
            Me.Close()

        End If
    End Sub

    Private Sub IconPictureBox6_Click(sender As Object, e As EventArgs) Handles IconPictureBox6.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit the application?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = DialogResult.Yes Then
            Me.Close()

        End If
    End Sub

    Private Sub IconPictureBox20_Click(sender As Object, e As EventArgs) Handles IconPictureBox20.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit the application?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = DialogResult.Yes Then
            Me.Close()

        End If
    End Sub
    Private Sub generateReportSection_Click(sender As Object, e As EventArgs) Handles generateReportSection.Click
        Try
            ' Get section data from the DataGridView (respects sorting)
            Dim sectionData As DataTable = GetSortedSectionDataFromGridView()
            If sectionData Is Nothing OrElse sectionData.Rows.Count = 0 Then
                MessageBox.Show("No section data available to generate report.", "Information")
                Return
            End If

            ' Setup save file dialog
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
            saveFileDialog.Title = "Save Section Summary Report"
            saveFileDialog.FileName = "Section_Summary_Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"

            If saveFileDialog.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            ' Create PDF document
            Dim document As New PdfDocument()
            document.Info.Title = "Section Summary Report"

            ' Setup fonts
            Dim fontTitle As New XFont("Segoe UI", 16, XFontStyle.Bold)
            Dim fontSubTitle As New XFont("Segoe UI", 9, XFontStyle.Regular)
            Dim fontHeader As New XFont("Segoe UI", 9, XFontStyle.Bold)
            Dim fontRow As New XFont("Segoe UI", 10, XFontStyle.Regular)

            ' === Load left logo from relative path ===
            Dim leftLogoPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Resources\plplogo.png")
            Dim leftLogo As XImage = Nothing
            If File.Exists(leftLogoPath) Then leftLogo = XImage.FromFile(leftLogoPath)

            ' === Load right logo from settings ===
            Dim rightLogoPath As String = My.Settings.LogoPath
            Dim rightLogo As XImage = Nothing
            If File.Exists(rightLogoPath) Then rightLogo = XImage.FromFile(rightLogoPath)

            ' Get department name with fallback
            Dim departmentName As String = My.Settings.DepartmentName
            If String.IsNullOrEmpty(departmentName) Then
                departmentName = "Not Specified"
            End If

            Dim DrawHeader = Sub(hdrGfx As XGraphics, hdrPage As PdfPage)
                                 ' Background rectangle
                                 hdrGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(245, 245, 245)), 0, 0, hdrPage.Width.Point, 80)

                                 ' Top-left text: small, single line
                                 Dim xLeft As Double = 20
                                 Dim yTop As Double = 5  ' small margin from top
                                 Dim smallFont As New XFont("Segoe UI", 7, XFontStyle.Regular)

                                 Dim adminName As String = "Administrator"
                                 If CurrentAdmin IsNot Nothing AndAlso Not String.IsNullOrEmpty(CurrentAdmin.LastName) Then
                                     adminName = CurrentAdmin.LastName.ToUpper() ' optional uppercase
                                 End If

                                 Dim mi As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".")
                                 Dim suffix As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", " " & CurrentAdmin.Suffix)
                                 Dim fullName As String = $"{CurrentAdmin.FirstName}{mi} {CurrentAdmin.LastName}{suffix}".Trim()

                                 Dim generatedText As String = $"Report generated by Admin Prof. {fullName} as of {DateTime.Now:MMMM dd, yyyy, 'at' hh:mm tt}"

                                 hdrGfx.DrawString(generatedText, smallFont, XBrushes.Black, New XRect(xLeft, yTop, 400, 10), XStringFormats.TopLeft)



                                 Dim targetHeight As Double = 50

                                 If leftLogo IsNot Nothing Then
                                     Dim aspectLeft As Double = leftLogo.PixelWidth / leftLogo.PixelHeight
                                     Dim newWidthLeft As Double = targetHeight * aspectLeft
                                     hdrGfx.DrawImage(leftLogo, xLeft, yTop + 15, newWidthLeft, targetHeight)
                                 End If

                                 If rightLogo IsNot Nothing Then
                                     Dim aspectRight As Double = rightLogo.PixelWidth / rightLogo.PixelHeight
                                     Dim newWidthRight As Double = targetHeight * aspectRight
                                     hdrGfx.DrawImage(rightLogo, hdrPage.Width.Point - newWidthRight - 20, yTop + 15, newWidthRight, targetHeight)
                                 End If


                                 ' Centered text below logos
                                 Dim centerX As Double = hdrPage.Width.Point / 2
                                 Dim yCenter As Double = yTop + 20  ' adjust vertical position below small top text

                                 hdrGfx.DrawString("Section Report", fontTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter, 300, 20), XStringFormats.TopCenter)
                                 hdrGfx.DrawString("Pamantasan ng Lungsod ng Pasig", fontSubTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter + 20, 300, 15), XStringFormats.TopCenter)
                                 hdrGfx.DrawString(departmentName, fontSubTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter + 35, 300, 15), XStringFormats.TopCenter)
                             End Sub



            ' Create first page
            Dim page As PdfPage = document.AddPage()
            page.Size = PageSize.A4
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            DrawHeader(gfx, page)

            ' Section table setup - Start table with proper margin below header
            Dim yPoint As Double = 85 ' Reduced from 90 to 85 for closer spacing
            Dim marginLeft As Integer = 20
            Dim marginRight As Integer = 20
            Dim lineHeight As Double = 15
            Dim tableWidth As Double = page.Width.Point - marginLeft - marginRight
            Dim headerText As String = "Section"

            ' Draw table header - Full width
            gfx.DrawRectangle(XPens.LightGray, marginLeft, yPoint, tableWidth, lineHeight + 4)
            gfx.DrawString(headerText, fontHeader, XBrushes.Black,
                          New XRect(marginLeft + 5, yPoint + 2, tableWidth - 10, lineHeight),
                          XStringFormats.TopLeft)
            yPoint += lineHeight + 6

            ' Draw section data rows
            Dim rowColorToggle As Boolean = False

            For Each row As DataRow In sectionData.Rows
                Dim maxRowHeight As Double = lineHeight * 1.5 ' Default row height

                ' Get only the section value
                Dim sectionValue As String = row("sectionName").ToString()

                ' Apply alternating row colors
                If rowColorToggle Then
                    gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(240, 240, 240)), marginLeft, yPoint, tableWidth, maxRowHeight)
                End If

                ' Draw section data - Full width with left alignment
                gfx.DrawRectangle(XPens.LightGray, marginLeft, yPoint, tableWidth, maxRowHeight)
                gfx.DrawString(sectionValue, fontRow, XBrushes.Black,
                              New XRect(marginLeft + 5, yPoint + 2, tableWidth - 10, maxRowHeight),
                              XStringFormats.TopLeft)

                yPoint += maxRowHeight
                rowColorToggle = Not rowColorToggle

                ' Check if we need a new page
                If yPoint > page.Height.Point - 40 Then
                    page = document.AddPage()
                    page.Size = PageSize.A4
                    gfx = XGraphics.FromPdfPage(page)
                    DrawHeader(gfx, page)

                    ' Reset to position below header on new page
                    yPoint = 85

                    ' Draw table header on new page - Full width
                    gfx.DrawRectangle(XPens.LightGray, marginLeft, yPoint, tableWidth, lineHeight + 4)
                    gfx.DrawString(headerText, fontHeader, XBrushes.Black,
                                  New XRect(marginLeft + 5, yPoint + 2, tableWidth - 10, lineHeight),
                                  XStringFormats.TopLeft)
                    yPoint += lineHeight + 6
                End If
            Next

            ' Save the document
            document.Save(saveFileDialog.FileName)
            MessageBox.Show("Section Summary Report generated successfully!", "Success")

        Catch ex As Exception
            MessageBox.Show("Failed to generate PDF: " & ex.Message, "PDF Error")
        End Try
    End Sub

    ' Function to get sorted section data from DataGridView
    Private Function GetSortedSectionDataFromGridView() As DataTable
        Try
            ' Create a new DataTable with the same structure
            Dim sortedData As New DataTable()
            sortedData.Columns.Add("sectionName", GetType(String))

            ' Get the sorted data from the DataGridView
            For Each row As DataGridViewRow In sectionView.Rows
                If Not row.IsNewRow Then
                    Dim newRow As DataRow = sortedData.NewRow()
                    newRow("sectionName") = row.Cells("sectionName").Value.ToString()
                    sortedData.Rows.Add(newRow)
                End If
            Next

            Return sortedData
        Catch ex As Exception
            MessageBox.Show("Failed to get sorted section data: " & ex.Message, "Error")
            Return GetAllSections() ' Fallback to database query
        End Try
    End Function

    Private Sub generateReportReason_Click(sender As Object, e As EventArgs) Handles generateReportReason.Click
        Try
            ' Get reason data from the DataGridView (respects sorting)
            Dim reasonData As DataTable = GetSortedReasonDataFromGridView()
            If reasonData Is Nothing OrElse reasonData.Rows.Count = 0 Then
                MessageBox.Show("No reason data available to generate report.", "Information")
                Return
            End If

            ' Setup save file dialog
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
            saveFileDialog.Title = "Save Reason Summary Report"
            saveFileDialog.FileName = "Reason_Summary_Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"

            If saveFileDialog.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            ' Create PDF document
            Dim document As New PdfDocument()
            document.Info.Title = "Reason Summary Report"

            ' Setup fonts
            Dim fontTitle As New XFont("Segoe UI", 16, XFontStyle.Bold)
            Dim fontSubTitle As New XFont("Segoe UI", 9, XFontStyle.Regular)
            Dim fontHeader As New XFont("Segoe UI", 9, XFontStyle.Bold)
            Dim fontRow As New XFont("Segoe UI", 10, XFontStyle.Regular)

            ' === Load left logo from relative path ===
            Dim leftLogoPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Resources\plplogo.png")
            Dim leftLogo As XImage = Nothing
            If File.Exists(leftLogoPath) Then leftLogo = XImage.FromFile(leftLogoPath)

            ' === Load right logo from settings ===
            Dim rightLogoPath As String = My.Settings.LogoPath
            Dim rightLogo As XImage = Nothing
            If File.Exists(rightLogoPath) Then rightLogo = XImage.FromFile(rightLogoPath)

            ' Get department name with fallback
            Dim departmentName As String = My.Settings.DepartmentName
            If String.IsNullOrEmpty(departmentName) Then
                departmentName = "Not Specified"
            End If

            ' === HEADER DRAWER with two logos ===
            Dim DrawHeader = Sub(hdrGfx As XGraphics, hdrPage As PdfPage)
                                 ' Background rectangle
                                 hdrGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(245, 245, 245)), 0, 0, hdrPage.Width.Point, 80)

                                 ' Top-left text: small, single line
                                 Dim xLeft As Double = 20
                                 Dim yTop As Double = 5  ' small margin from top
                                 Dim smallFont As New XFont("Segoe UI", 7, XFontStyle.Regular)

                                 Dim adminName As String = "Administrator"
                                 If CurrentAdmin IsNot Nothing AndAlso Not String.IsNullOrEmpty(CurrentAdmin.LastName) Then
                                     adminName = CurrentAdmin.LastName.ToUpper() ' optional uppercase
                                 End If

                                 Dim mi As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".")
                                 Dim suffix As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", " " & CurrentAdmin.Suffix)
                                 Dim fullName As String = $"{CurrentAdmin.FirstName}{mi} {CurrentAdmin.LastName}{suffix}".Trim()

                                 Dim generatedText As String = $"Report generated by Admin Prof. {fullName} as of {DateTime.Now:MMMM dd, yyyy, 'at' hh:mm tt}"

                                 hdrGfx.DrawString(generatedText, smallFont, XBrushes.Black, New XRect(xLeft, yTop, 400, 10), XStringFormats.TopLeft)




                                 ' Left logo (PLP logo from Resources) below the small text
                                 Dim targetHeight As Double = 50

                                 If leftLogo IsNot Nothing Then
                                     Dim aspectLeft As Double = leftLogo.PixelWidth / leftLogo.PixelHeight
                                     Dim newWidthLeft As Double = targetHeight * aspectLeft
                                     hdrGfx.DrawImage(leftLogo, xLeft, yTop + 15, newWidthLeft, targetHeight)
                                 End If

                                 If rightLogo IsNot Nothing Then
                                     Dim aspectRight As Double = rightLogo.PixelWidth / rightLogo.PixelHeight
                                     Dim newWidthRight As Double = targetHeight * aspectRight
                                     hdrGfx.DrawImage(rightLogo, hdrPage.Width.Point - newWidthRight - 20, yTop + 15, newWidthRight, targetHeight)
                                 End If


                                 ' Centered text below logos
                                 Dim centerX As Double = hdrPage.Width.Point / 2
                                 Dim yCenter As Double = yTop + 20  ' adjust vertical position below small top text

                                 hdrGfx.DrawString("Reason Report", fontTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter, 300, 20), XStringFormats.TopCenter)
                                 hdrGfx.DrawString("Pamantasan ng Lungsod ng Pasig", fontSubTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter + 20, 300, 15), XStringFormats.TopCenter)
                                 hdrGfx.DrawString(departmentName, fontSubTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter + 35, 300, 15), XStringFormats.TopCenter)
                             End Sub


            ' Create first page
            Dim page As PdfPage = document.AddPage()
            page.Size = PageSize.A4
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            DrawHeader(gfx, page)

            ' Reason table setup - Start table with proper margin below header
            Dim yPoint As Double = 85 ' Reduced from 90 to 85 for closer spacing
            Dim marginLeft As Integer = 20
            Dim marginRight As Integer = 20
            Dim lineHeight As Double = 15
            Dim tableWidth As Double = page.Width.Point - marginLeft - marginRight

            ' Column widths (70% for reason, 30% for status)
            Dim colWidthReason As Double = tableWidth * 0.7
            Dim colWidthStatus As Double = tableWidth * 0.3

            Dim headers() As String = {"Reason", "Special Event"}

            ' Draw table headers
            Dim xPos As Double = marginLeft
            gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidthReason, lineHeight + 4)
            gfx.DrawString(headers(0), fontHeader, XBrushes.Black,
                      New XRect(xPos + 5, yPoint + 2, colWidthReason - 10, lineHeight),
                      XStringFormats.TopLeft)

            xPos += colWidthReason
            gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidthStatus, lineHeight + 4)
            gfx.DrawString(headers(1), fontHeader, XBrushes.Black,
                      New XRect(xPos + 5, yPoint + 2, colWidthStatus - 10, lineHeight),
                      XStringFormats.TopLeft)

            yPoint += lineHeight + 6

            ' Draw reason data rows
            Dim rowColorToggle As Boolean = False

            For Each row As DataRow In reasonData.Rows
                Dim maxRowHeight As Double = lineHeight * 1.5 ' Default row height

                ' Get reason and special values - FIXED: Using "special" instead of "is_special"
                Dim reasonValue As String = row("reason").ToString()
                Dim statusValue As String = row("special").ToString()

                ' Apply alternating row colors
                If rowColorToggle Then
                    gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(240, 240, 240)), marginLeft, yPoint, tableWidth, maxRowHeight)
                End If

                ' Draw reason data
                xPos = marginLeft
                gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidthReason, maxRowHeight)
                gfx.DrawString(reasonValue, fontRow, XBrushes.Black,
                          New XRect(xPos + 5, yPoint + 2, colWidthReason - 10, maxRowHeight),
                          XStringFormats.TopLeft)

                xPos += colWidthReason
                gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidthStatus, maxRowHeight)
                gfx.DrawString(statusValue, fontRow, XBrushes.Black,
                          New XRect(xPos + 5, yPoint + 2, colWidthStatus - 10, maxRowHeight),
                          XStringFormats.TopLeft)

                yPoint += maxRowHeight
                rowColorToggle = Not rowColorToggle

                ' Check if we need a new page
                If yPoint > page.Height.Point - 40 Then
                    page = document.AddPage()
                    page.Size = PageSize.A4
                    gfx = XGraphics.FromPdfPage(page)
                    DrawHeader(gfx, page)

                    ' Reset to position below header on new page
                    yPoint = 85
                    xPos = marginLeft

                    ' Draw table headers on new page
                    gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidthReason, lineHeight + 4)
                    gfx.DrawString(headers(0), fontHeader, XBrushes.Black,
                              New XRect(xPos + 5, yPoint + 2, colWidthReason - 10, lineHeight),
                              XStringFormats.TopLeft)

                    xPos += colWidthReason
                    gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidthStatus, lineHeight + 4)
                    gfx.DrawString(headers(1), fontHeader, XBrushes.Black,
                              New XRect(xPos + 5, yPoint + 2, colWidthStatus - 10, lineHeight),
                              XStringFormats.TopLeft)

                    yPoint += lineHeight + 6
                End If
            Next

            ' Save the document
            document.Save(saveFileDialog.FileName)
            MessageBox.Show("Reason Summary Report generated successfully!", "Success")

        Catch ex As Exception
            MessageBox.Show("Failed to generate PDF: " & ex.Message, "PDF Error")
        End Try
    End Sub

    ' Function to get sorted reason data from DataGridView
    Private Function GetSortedReasonDataFromGridView() As DataTable
        Try
            ' Create a new DataTable with the same structure
            Dim sortedData As New DataTable()
            sortedData.Columns.Add("reason", GetType(String))
            sortedData.Columns.Add("special", GetType(String)) ' FIXED: Changed from "is_special" to "special"

            ' Get the sorted data from the DataGridView
            For Each row As DataGridViewRow In reasonView.Rows
                If Not row.IsNewRow Then
                    Dim newRow As DataRow = sortedData.NewRow()
                    newRow("reason") = row.Cells("reason").Value.ToString()
                    newRow("special") = (row.Cells("special").Value.ToString()) ' FIXED: Changed from "is_special" to "special"
                    sortedData.Rows.Add(newRow)
                End If
            Next

            Return sortedData
        Catch ex As Exception
            MessageBox.Show("Failed to get sorted reason data: " & ex.Message, "Error")

        End Try
    End Function

    Private Sub generateReportAdmin_Click(sender As Object, e As EventArgs) Handles generateReportAdmin.Click
        Try
            ' Get admin data from the DataGridView (respects sorting)
            Dim adminData As DataTable = GetSortedAdminDataFromGridView()
            If adminData Is Nothing OrElse adminData.Rows.Count = 0 Then
                MessageBox.Show("No administrator data available to generate report.", "Information")
                Return
            End If

            ' Setup save file dialog
            Using saveFileDialog As New SaveFileDialog()
                saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
                saveFileDialog.Title = "Save Administrator Report"
                saveFileDialog.FileName = "Administrator_Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"

                If saveFileDialog.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                ' Create PDF document
                Using document As New PdfDocument()
                    document.Info.Title = "Administrator Report"

                    ' Setup fonts
                    Dim fontTitle As New XFont("Segoe UI", 16, XFontStyle.Bold)
                    Dim fontSubTitle As New XFont("Segoe UI", 9, XFontStyle.Regular)
                    Dim fontHeader As New XFont("Segoe UI", 9, XFontStyle.Bold)
                    Dim fontRow As New XFont("Segoe UI", 9, XFontStyle.Regular)

                    ' === Load left logo from relative path ===
                    Dim leftLogoPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Resources\plplogo.png")
                    Dim leftLogo As XImage = If(File.Exists(leftLogoPath), XImage.FromFile(leftLogoPath), Nothing)

                    ' === Load right logo from settings ===
                    Dim rightLogoPath As String = My.Settings.LogoPath
                    Dim rightLogo As XImage = If(File.Exists(rightLogoPath), XImage.FromFile(rightLogoPath), Nothing)

                    ' Get department name with fallback
                    Dim departmentName As String = If(String.IsNullOrEmpty(My.Settings.DepartmentName), "Not Specified", My.Settings.DepartmentName)

                    ' === HEADER DRAWER with two logos ===
                    Dim DrawHeader = Sub(hdrGfx As XGraphics, hdrPage As PdfPage)
                                         ' Background rectangle
                                         hdrGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(245, 245, 245)), 0, 0, hdrPage.Width.Point, 80)

                                         ' Top-left text: small, single line
                                         Dim xLeft As Double = 20
                                         Dim yTop As Double = 5  ' small margin from top
                                         Dim smallFont As New XFont("Segoe UI", 7, XFontStyle.Regular)

                                         Dim adminName As String = "Administrator"
                                         If CurrentAdmin IsNot Nothing AndAlso Not String.IsNullOrEmpty(CurrentAdmin.LastName) Then
                                             adminName = CurrentAdmin.LastName.ToUpper() ' optional uppercase
                                         End If

                                         Dim mi As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".")
                                         Dim suffix As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", " " & CurrentAdmin.Suffix)
                                         Dim fullName As String = $"{CurrentAdmin.FirstName}{mi} {CurrentAdmin.LastName}{suffix}".Trim()

                                         Dim generatedText As String = $"Report generated by Admin Prof. {fullName} as of {DateTime.Now:MMMM dd, yyyy, 'at' hh:mm tt}"

                                         hdrGfx.DrawString(generatedText, smallFont, XBrushes.Black, New XRect(xLeft, yTop, 400, 10), XStringFormats.TopLeft)



                                         ' Left logo (PLP logo from Resources) below the small text
                                         Dim targetHeight As Double = 50

                                         If leftLogo IsNot Nothing Then
                                             Dim aspectLeft As Double = leftLogo.PixelWidth / leftLogo.PixelHeight
                                             Dim newWidthLeft As Double = targetHeight * aspectLeft
                                             hdrGfx.DrawImage(leftLogo, xLeft, yTop + 15, newWidthLeft, targetHeight)
                                         End If

                                         If rightLogo IsNot Nothing Then
                                             Dim aspectRight As Double = rightLogo.PixelWidth / rightLogo.PixelHeight
                                             Dim newWidthRight As Double = targetHeight * aspectRight
                                             hdrGfx.DrawImage(rightLogo, hdrPage.Width.Point - newWidthRight - 20, yTop + 15, newWidthRight, targetHeight)
                                         End If


                                         ' Centered text below logos
                                         Dim centerX As Double = hdrPage.Width.Point / 2
                                         Dim yCenter As Double = yTop + 20  ' adjust vertical position below small top text

                                         hdrGfx.DrawString("Admin Report", fontTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter, 300, 20), XStringFormats.TopCenter)
                                         hdrGfx.DrawString("Pamantasan ng Lungsod ng Pasig", fontSubTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter + 20, 300, 15), XStringFormats.TopCenter)
                                         hdrGfx.DrawString(departmentName, fontSubTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter + 35, 300, 15), XStringFormats.TopCenter)
                                     End Sub



                    ' Create first page
                    Dim page As PdfPage = document.AddPage()
                    page.Size = PageSize.A4
                    Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                    Dim tf As New XTextFormatter(gfx)
                    DrawHeader(gfx, page)

                    Dim yPoint As Double = 90 ' <- Slight margin before the table
                    Dim marginLeft As Integer = 10
                    Dim lineHeight As Double = 12
                    Dim colWidths() As Integer = {90, 90, 30, 40, 180, 80, 70} ' Adjusted widths
                    Dim headers() As String = {"First Name", "Last Name", "M.I.", "Suffix", "Email", "Username", "Password"}

                    Dim xPos As Double = marginLeft
                    For i = 0 To headers.Length - 1
                        gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                        gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopLeft)
                        xPos += colWidths(i)
                    Next
                    yPoint += lineHeight + 6

                    Dim rowColorToggle As Boolean = False

                    For Each row As DataRow In adminData.Rows
                        xPos = marginLeft
                        Dim maxRowHeight As Double = 0
                        Dim rowData() As String = {
                        row("adminFirstName").ToString(),
                        row("adminLastName").ToString(),
                        row("adminMiddleInitial").ToString(),
                        row("adminSuffix").ToString(),
                        row("adminEmail").ToString(),
                        row("adminUsername").ToString(),
                        "••••••••" ' Masked password
                    }

                        ' PROPER MEASUREMENT - This is what makes it work!
                        For i = 0 To rowData.Length - 1
                            Dim layoutRect As New XRect(0, 0, colWidths(i) - 4, Double.MaxValue)
                            Dim dummyGfx As XGraphics = XGraphics.CreateMeasureContext(New XSize(colWidths(i) - 4, Double.MaxValue), XGraphicsUnit.Point, XPageDirection.Downwards)
                            Dim size As XSize = dummyGfx.MeasureString(rowData(i), fontRow)
                            Dim linesNeeded As Integer = Math.Ceiling(size.Width / layoutRect.Width)
                            Dim heightNeeded As Double = linesNeeded * lineHeight
                            If heightNeeded > maxRowHeight Then maxRowHeight = heightNeeded
                        Next

                        If maxRowHeight < lineHeight * 2 Then maxRowHeight = lineHeight * 2

                        If rowColorToggle Then
                            gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(240, 240, 240)), marginLeft, yPoint, colWidths.Sum(), maxRowHeight)
                        End If

                        xPos = marginLeft
                        For i = 0 To rowData.Length - 1
                            tf.DrawString(rowData(i), fontRow, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, maxRowHeight), XStringFormats.TopLeft)
                            gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), maxRowHeight)
                            xPos += colWidths(i)
                        Next

                        yPoint += maxRowHeight
                        rowColorToggle = Not rowColorToggle

                        If yPoint > page.Height.Point - 40 Then
                            page = document.AddPage()
                            page.Size = PageSize.A4
                            gfx = XGraphics.FromPdfPage(page)
                            tf = New XTextFormatter(gfx)
                            DrawHeader(gfx, page)

                            yPoint = 90 ' Maintain top margin
                            xPos = marginLeft
                            For i = 0 To headers.Length - 1
                                gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                                gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopLeft)
                                xPos += colWidths(i)
                            Next
                            yPoint += lineHeight + 6
                        End If
                    Next

                    document.Save(saveFileDialog.FileName)
                    MessageBox.Show("Administrator Report generated successfully!", "Success")
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Failed to generate PDF: " & ex.Message, "PDF Error")
        End Try
    End Sub

    ' Helper function to get cell value based on header
    Private Function GetCellValue(row As DataRow, header As String) As String
        Select Case header
            Case "First Name"
                Return row("adminFirstName").ToString()
            Case "Last Name"
                Return row("adminLastName").ToString()
            Case "M.I."
                Return row("adminMiddleInitial").ToString()
            Case "Suffix"
                Return row("adminSuffix").ToString()
            Case "Email"
                Return row("adminEmail").ToString()
            Case "Username"
                Return row("adminUsername").ToString()
            Case "Password"
                Return "••••••••" ' Mask password for security
            Case Else
                Return ""
        End Select
    End Function

    ' Helper function to calculate text height needed for wrapping
    Private Function CalculateTextHeight(text As String, maxWidth As Double, font As XFont) As Double
        If String.IsNullOrEmpty(text) Then Return 15

        ' Simple estimation: assume average character width and calculate lines needed
        Dim avgCharWidth As Double = font.Size * 0.6 ' Approximate average character width
        Dim charsPerLine As Integer = CInt(Math.Floor(maxWidth / avgCharWidth))

        If charsPerLine <= 0 Then Return 15

        Dim linesNeeded As Integer = CInt(Math.Ceiling(text.Length / charsPerLine))
        Return linesNeeded * font.Height
    End Function

    ' Function to get sorted admin data from DataGridView
    Private Function GetSortedAdminDataFromGridView() As DataTable
        Try
            ' Create a new DataTable with the same structure
            Dim sortedData As New DataTable()
            sortedData.Columns.Add("adminFirstName", GetType(String))
            sortedData.Columns.Add("adminLastName", GetType(String))
            sortedData.Columns.Add("adminMiddleInitial", GetType(String))
            sortedData.Columns.Add("adminSuffix", GetType(String))
            sortedData.Columns.Add("adminEmail", GetType(String))
            sortedData.Columns.Add("adminUsername", GetType(String))
            sortedData.Columns.Add("adminPassword", GetType(String))

            ' Get the sorted data from the DataGridView
            For Each row As DataGridViewRow In adminView.Rows
                If Not row.IsNewRow Then
                    Dim newRow As DataRow = sortedData.NewRow()
                    newRow("adminFirstName") = If(row.Cells("adminFirstName").Value IsNot Nothing, row.Cells("adminFirstName").Value.ToString(), "")
                    newRow("adminLastName") = If(row.Cells("adminLastName").Value IsNot Nothing, row.Cells("adminLastName").Value.ToString(), "")
                    newRow("adminMiddleInitial") = If(row.Cells("adminMiddleInitial").Value IsNot Nothing, row.Cells("adminMiddleInitial").Value.ToString(), "")
                    newRow("adminSuffix") = If(row.Cells("adminSuffix").Value IsNot Nothing, row.Cells("adminSuffix").Value.ToString(), "")
                    newRow("adminEmail") = If(row.Cells("adminEmail").Value IsNot Nothing, row.Cells("adminEmail").Value.ToString(), "")
                    newRow("adminUsername") = If(row.Cells("adminUsername").Value IsNot Nothing, row.Cells("adminUsername").Value.ToString(), "")
                    newRow("adminPassword") = If(row.Cells("adminPassword").Value IsNot Nothing, row.Cells("adminPassword").Value.ToString(), "")
                    sortedData.Rows.Add(newRow)
                End If
            Next

            Return sortedData
        Catch ex As Exception
            MessageBox.Show("Failed to get sorted admin data: " & ex.Message, "Error")
            Return Nothing
        End Try
    End Function

    Private Sub generateReportProfessor_Click(sender As Object, e As EventArgs) Handles generateReportProfessor.Click
        Try
            ' Get professor data from the DataGridView (respects sorting)
            Dim professorData As DataTable = GetSortedProfessorDataFromGridView()
            If professorData Is Nothing OrElse professorData.Rows.Count = 0 Then
                MessageBox.Show("No professor data available to generate report.", "Information")
                Return
            End If

            ' Setup save file dialog
            Using saveFileDialog As New SaveFileDialog()
                saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
                saveFileDialog.Title = "Save Professor Report"
                saveFileDialog.FileName = "Professor_Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"

                If saveFileDialog.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                ' Create PDF document
                Using document As New PdfDocument()
                    document.Info.Title = "Professor Report"

                    ' Setup fonts
                    Dim fontTitle As New XFont("Segoe UI", 16, XFontStyle.Bold)
                    Dim fontSubTitle As New XFont("Segoe UI", 9, XFontStyle.Regular)
                    Dim fontHeader As New XFont("Segoe UI", 9, XFontStyle.Bold)
                    Dim fontRow As New XFont("Segoe UI", 9, XFontStyle.Regular)

                    ' === Load left logo from relative path ===
                    Dim leftLogoPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Resources\plplogo.png")
                    Dim leftLogo As XImage = If(File.Exists(leftLogoPath), XImage.FromFile(leftLogoPath), Nothing)

                    ' === Load right logo from settings ===
                    Dim rightLogoPath As String = My.Settings.LogoPath
                    Dim rightLogo As XImage = If(File.Exists(rightLogoPath), XImage.FromFile(rightLogoPath), Nothing)

                    ' Get department name with fallback
                    Dim departmentName As String = If(String.IsNullOrEmpty(My.Settings.DepartmentName), "Not Specified", My.Settings.DepartmentName)

                    ' === HEADER DRAWER with two logos ===
                    Dim DrawHeader = Sub(hdrGfx As XGraphics, hdrPage As PdfPage)
                                         ' Background rectangle
                                         hdrGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(245, 245, 245)), 0, 0, hdrPage.Width.Point, 80)

                                         ' Top-left text: small, single line
                                         Dim xLeft As Double = 20
                                         Dim yTop As Double = 5  ' small margin from top
                                         Dim smallFont As New XFont("Segoe UI", 7, XFontStyle.Regular)

                                         Dim adminName As String = "Administrator"
                                         If CurrentAdmin IsNot Nothing AndAlso Not String.IsNullOrEmpty(CurrentAdmin.LastName) Then
                                             adminName = CurrentAdmin.LastName.ToUpper() ' optional uppercase
                                         End If
                                         Dim mi As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".")
                                         Dim suffix As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", " " & CurrentAdmin.Suffix)
                                         Dim fullName As String = $"{CurrentAdmin.FirstName}{mi} {CurrentAdmin.LastName}{suffix}".Trim()

                                         Dim generatedText As String = $"Report generated by Admin Prof. {fullName} as of {DateTime.Now:MMMM dd, yyyy, 'at' hh:mm tt}"

                                         hdrGfx.DrawString(generatedText, smallFont, XBrushes.Black, New XRect(xLeft, yTop, 400, 10), XStringFormats.TopLeft)


                                         Dim targetHeight As Double = 50

                                         If leftLogo IsNot Nothing Then
                                             Dim aspectLeft As Double = leftLogo.PixelWidth / leftLogo.PixelHeight
                                             Dim newWidthLeft As Double = targetHeight * aspectLeft
                                             hdrGfx.DrawImage(leftLogo, xLeft, yTop + 15, newWidthLeft, targetHeight)
                                         End If

                                         If rightLogo IsNot Nothing Then
                                             Dim aspectRight As Double = rightLogo.PixelWidth / rightLogo.PixelHeight
                                             Dim newWidthRight As Double = targetHeight * aspectRight
                                             hdrGfx.DrawImage(rightLogo, hdrPage.Width.Point - newWidthRight - 20, yTop + 15, newWidthRight, targetHeight)
                                         End If

                                         ' Centered text below logos
                                         Dim centerX As Double = hdrPage.Width.Point / 2
                                         Dim yCenter As Double = yTop + 20  ' adjust vertical position below small top text

                                         hdrGfx.DrawString("Professor Report", fontTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter, 300, 20), XStringFormats.TopCenter)
                                         hdrGfx.DrawString("Pamantasan ng Lungsod ng Pasig", fontSubTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter + 20, 300, 15), XStringFormats.TopCenter)
                                         hdrGfx.DrawString(departmentName, fontSubTitle, XBrushes.Black,
                                                   New XRect(centerX - 150, yCenter + 35, 300, 15), XStringFormats.TopCenter)
                                     End Sub


                    ' Create first page
                    Dim page As PdfPage = document.AddPage()
                    page.Size = PageSize.A4
                    Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                    Dim tf As New XTextFormatter(gfx)
                    DrawHeader(gfx, page)

                    Dim yPoint As Double = 90 ' <- Slight margin before the table
                    Dim marginLeft As Integer = 10
                    Dim lineHeight As Double = 12
                    Dim colWidths() As Integer = {90, 90, 30, 40, 180, 80, 70} ' Adjusted widths
                    Dim headers() As String = {"First Name", "Last Name", "M.I.", "Suffix", "Email", "Username", "Password"}

                    Dim xPos As Double = marginLeft
                    For i = 0 To headers.Length - 1
                        gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                        gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopLeft)
                        xPos += colWidths(i)
                    Next
                    yPoint += lineHeight + 6

                    Dim rowColorToggle As Boolean = False

                    For Each row As DataRow In professorData.Rows
                        xPos = marginLeft
                        Dim maxRowHeight As Double = 0
                        Dim rowData() As String = {
                        row("firstNameCol").ToString(),
                        row("lastNameCol").ToString(),
                        row("middleInitialCol").ToString(),
                        row("suffixCol").ToString(),
                        row("emailCol").ToString(),
                        row("usernameCol").ToString(),
                        "••••••••" ' Masked password
                    }

                        ' PROPER MEASUREMENT - This is what makes it work!
                        For i = 0 To rowData.Length - 1
                            Dim layoutRect As New XRect(0, 0, colWidths(i) - 4, Double.MaxValue)
                            Dim dummyGfx As XGraphics = XGraphics.CreateMeasureContext(New XSize(colWidths(i) - 4, Double.MaxValue), XGraphicsUnit.Point, XPageDirection.Downwards)
                            Dim size As XSize = dummyGfx.MeasureString(rowData(i), fontRow)
                            Dim linesNeeded As Integer = Math.Ceiling(size.Width / layoutRect.Width)
                            Dim heightNeeded As Double = linesNeeded * lineHeight
                            If heightNeeded > maxRowHeight Then maxRowHeight = heightNeeded
                        Next

                        If maxRowHeight < lineHeight * 2 Then maxRowHeight = lineHeight * 2

                        If rowColorToggle Then
                            gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(240, 240, 240)), marginLeft, yPoint, colWidths.Sum(), maxRowHeight)
                        End If

                        xPos = marginLeft
                        For i = 0 To rowData.Length - 1
                            tf.DrawString(rowData(i), fontRow, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, maxRowHeight), XStringFormats.TopLeft)
                            gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), maxRowHeight)
                            xPos += colWidths(i)
                        Next

                        yPoint += maxRowHeight
                        rowColorToggle = Not rowColorToggle

                        If yPoint > page.Height.Point - 40 Then
                            page = document.AddPage()
                            page.Size = PageSize.A4
                            gfx = XGraphics.FromPdfPage(page)
                            tf = New XTextFormatter(gfx)
                            DrawHeader(gfx, page)

                            yPoint = 90 ' Maintain top margin
                            xPos = marginLeft
                            For i = 0 To headers.Length - 1
                                gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                                gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopLeft)
                                xPos += colWidths(i)
                            Next
                            yPoint += lineHeight + 6
                        End If
                    Next

                    document.Save(saveFileDialog.FileName)
                    MessageBox.Show("Professor Report generated successfully!", "Success")
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Failed to generate PDF: " & ex.Message, "PDF Error")
        End Try
    End Sub

    ' Function to get sorted professor data from DataGridView
    Private Function GetSortedProfessorDataFromGridView() As DataTable
        Try
            ' Create a new DataTable with the same structure
            Dim sortedData As New DataTable()
            sortedData.Columns.Add("firstNameCol", GetType(String))
            sortedData.Columns.Add("lastNameCol", GetType(String))
            sortedData.Columns.Add("middleInitialCol", GetType(String))
            sortedData.Columns.Add("suffixCol", GetType(String))
            sortedData.Columns.Add("emailCol", GetType(String))
            sortedData.Columns.Add("usernameCol", GetType(String))
            sortedData.Columns.Add("passwordCol", GetType(String))

            ' Get the sorted data from the DataGridView
            For Each row As DataGridViewRow In professorView.Rows
                If Not row.IsNewRow Then
                    Dim newRow As DataRow = sortedData.NewRow()
                    newRow("firstNameCol") = If(row.Cells("firstNameCol").Value IsNot Nothing, row.Cells("firstNameCol").Value.ToString(), "")
                    newRow("lastNameCol") = If(row.Cells("lastNameCol").Value IsNot Nothing, row.Cells("lastNameCol").Value.ToString(), "")
                    newRow("middleInitialCol") = If(row.Cells("middleInitialCol").Value IsNot Nothing, row.Cells("middleInitialCol").Value.ToString(), "")
                    newRow("suffixCol") = If(row.Cells("suffixCol").Value IsNot Nothing, row.Cells("suffixCol").Value.ToString(), "")
                    newRow("emailCol") = If(row.Cells("emailCol").Value IsNot Nothing, row.Cells("emailCol").Value.ToString(), "")
                    newRow("usernameCol") = If(row.Cells("usernameCol").Value IsNot Nothing, row.Cells("usernameCol").Value.ToString(), "")
                    newRow("passwordCol") = If(row.Cells("passwordCol").Value IsNot Nothing, row.Cells("passwordCol").Value.ToString(), "")
                    sortedData.Rows.Add(newRow)
                End If
            Next

            Return sortedData
        Catch ex As Exception
            MessageBox.Show("Failed to get sorted professor data: " & ex.Message, "Error")
            Return Nothing
        End Try
    End Function
    Private Sub generateReportStudent_Click(sender As Object, e As EventArgs) Handles generateReportStudent.Click
        Try
            ' Get student data from the DataGridView (respects sorting)
            Dim studentData As DataTable = GetSortedStudentDataFromGridView()
            If studentData Is Nothing OrElse studentData.Rows.Count = 0 Then
                MessageBox.Show("No student data available to generate report.", "Information")
                Return
            End If

            ' Setup save file dialog
            Using saveFileDialog As New SaveFileDialog()
                saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
                saveFileDialog.Title = "Save Student Report"
                saveFileDialog.FileName = "Student_Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"

                If saveFileDialog.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                ' Create PDF document
                Using document As New PdfDocument()
                    document.Info.Title = "Student Report"

                    ' Setup fonts
                    Dim fontTitle As New XFont("Segoe UI", 16, XFontStyle.Bold)
                    Dim fontSubTitle As New XFont("Segoe UI", 9, XFontStyle.Regular)
                    Dim fontHeader As New XFont("Segoe UI", 8, XFontStyle.Bold) ' Smaller font for headers
                    Dim fontRow As New XFont("Segoe UI", 8, XFontStyle.Regular) ' Smaller font for rows

                    ' === Load left logo from relative path ===
                    Dim leftLogoPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Resources\plplogo.png")
                    Dim leftLogo As XImage = If(File.Exists(leftLogoPath), XImage.FromFile(leftLogoPath), Nothing)

                    ' === Load right logo from settings ===
                    Dim rightLogoPath As String = My.Settings.LogoPath
                    Dim rightLogo As XImage = If(File.Exists(rightLogoPath), XImage.FromFile(rightLogoPath), Nothing)

                    ' Get department name with fallback
                    Dim departmentName As String = If(String.IsNullOrEmpty(My.Settings.DepartmentName), "Not Specified", My.Settings.DepartmentName)

                    Dim DrawHeader = Sub(hdrGfx As XGraphics, hdrPage As PdfPage)
                                         ' Background rectangle
                                         hdrGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(245, 245, 245)), 0, 0, hdrPage.Width.Point, 80)

                                         ' Top-left text: small, single line
                                         Dim xLeft As Double = 20
                                         Dim yTop As Double = 5  ' small margin from top
                                         Dim smallFont As New XFont("Segoe UI", 7, XFontStyle.Regular)

                                         Dim adminName As String = "Administrator"
                                         If CurrentAdmin IsNot Nothing AndAlso Not String.IsNullOrEmpty(CurrentAdmin.LastName) Then
                                             adminName = CurrentAdmin.LastName.ToUpper() ' optional uppercase
                                         End If
                                         Dim mi As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".")
                                         Dim suffix As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", " " & CurrentAdmin.Suffix)
                                         Dim fullName As String = $"{CurrentAdmin.FirstName}{mi} {CurrentAdmin.LastName}{suffix}".Trim()

                                         Dim generatedText As String = $"Report generated by Admin Prof. {fullName} as of {DateTime.Now:MMMM dd, yyyy, 'at' hh:mm tt}"

                                         hdrGfx.DrawString(generatedText, smallFont, XBrushes.Black, New XRect(xLeft, yTop, 400, 10), XStringFormats.TopLeft)

                                         Dim targetHeight As Double = 50

                                         If leftLogo IsNot Nothing Then
                                             Dim aspectLeft As Double = leftLogo.PixelWidth / leftLogo.PixelHeight
                                             Dim newWidthLeft As Double = targetHeight * aspectLeft
                                             hdrGfx.DrawImage(leftLogo, xLeft, yTop + 15, newWidthLeft, targetHeight)
                                         End If

                                         If rightLogo IsNot Nothing Then
                                             Dim aspectRight As Double = rightLogo.PixelWidth / rightLogo.PixelHeight
                                             Dim newWidthRight As Double = targetHeight * aspectRight
                                             hdrGfx.DrawImage(rightLogo, hdrPage.Width.Point - newWidthRight - 20, yTop + 15, newWidthRight, targetHeight)
                                         End If

                                         ' Centered text below logos
                                         Dim centerX As Double = hdrPage.Width.Point / 2
                                         Dim yCenter As Double = yTop + 20  ' adjust vertical position below small top text

                                         hdrGfx.DrawString("Student Report", fontTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter, 300, 20), XStringFormats.TopCenter)
                                         hdrGfx.DrawString("Pamantasan ng Lungsod ng Pasig", fontSubTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter + 20, 300, 15), XStringFormats.TopCenter)
                                         hdrGfx.DrawString(departmentName, fontSubTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter + 35, 300, 15), XStringFormats.TopCenter)
                                     End Sub



                    ' Create first page
                    Dim page As PdfPage = document.AddPage()
                    page.Size = PageSize.A4
                    Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                    Dim tf As New XTextFormatter(gfx)
                    DrawHeader(gfx, page)

                    Dim yPoint As Double = 90 ' <- Slight margin before the table
                    Dim marginLeft As Integer = 10
                    Dim lineHeight As Double = 12
                    ' Column widths for 8 columns (adjusted for portrait with text wrapping)
                    Dim colWidths() As Integer = {55, 100, 70, 30, 40, 50, 180, 50} ' Narrower widths with text wrapping
                    Dim headers() As String = {"Stud No", "First", "Last", "MI", "Suffix", "Section", "Email", "Status"}

                    Dim xPos As Double = marginLeft
                    For i = 0 To headers.Length - 1
                        gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                        gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopCenter)
                        xPos += colWidths(i)
                    Next
                    yPoint += lineHeight + 6

                    Dim rowColorToggle As Boolean = False

                    For Each row As DataRow In studentData.Rows
                        xPos = marginLeft
                        Dim maxRowHeight As Double = 0
                        Dim rowData() As String = {
                        row("studentNumber").ToString(),
                        row("firstName").ToString(),
                        row("lastName").ToString(),
                        row("middleInitial").ToString(),
                        row("suffix").ToString(),
                        row("section").ToString(),
                        row("email").ToString(),
                        row("status").ToString()
                    }

                        ' PROPER MEASUREMENT - This is what makes it work!
                        For i = 0 To rowData.Length - 1
                            Dim layoutRect As New XRect(0, 0, colWidths(i) - 4, Double.MaxValue)
                            Dim dummyGfx As XGraphics = XGraphics.CreateMeasureContext(New XSize(colWidths(i) - 4, Double.MaxValue), XGraphicsUnit.Point, XPageDirection.Downwards)
                            Dim size As XSize = dummyGfx.MeasureString(rowData(i), fontRow)
                            Dim linesNeeded As Integer = Math.Ceiling(size.Width / layoutRect.Width)
                            Dim heightNeeded As Double = linesNeeded * lineHeight
                            If heightNeeded > maxRowHeight Then maxRowHeight = heightNeeded
                        Next

                        If maxRowHeight < lineHeight * 2 Then maxRowHeight = lineHeight * 2

                        If rowColorToggle Then
                            gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(240, 240, 240)), marginLeft, yPoint, colWidths.Sum(), maxRowHeight)
                        End If

                        xPos = marginLeft
                        For i = 0 To rowData.Length - 1
                            tf.DrawString(rowData(i), fontRow, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, maxRowHeight), XStringFormats.TopLeft)
                            gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), maxRowHeight)
                            xPos += colWidths(i)
                        Next

                        yPoint += maxRowHeight
                        rowColorToggle = Not rowColorToggle

                        If yPoint > page.Height.Point - 40 Then
                            page = document.AddPage()
                            page.Size = PageSize.A4
                            gfx = XGraphics.FromPdfPage(page)
                            tf = New XTextFormatter(gfx)
                            DrawHeader(gfx, page)

                            yPoint = 90 ' Maintain top margin
                            xPos = marginLeft
                            For i = 0 To headers.Length - 1
                                gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                                gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopCenter)
                                xPos += colWidths(i)
                            Next
                            yPoint += lineHeight + 6
                        End If
                    Next

                    document.Save(saveFileDialog.FileName)
                    MessageBox.Show("Student Report generated successfully!", "Success")
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Failed to generate PDF: " & ex.Message, "PDF Error")
        End Try
    End Sub

    ' Function to get sorted student data from DataGridView
    Private Function GetSortedStudentDataFromGridView() As DataTable
        Try
            ' Create a new DataTable with the same structure
            Dim sortedData As New DataTable()
            sortedData.Columns.Add("studentNumber", GetType(String))
            sortedData.Columns.Add("firstName", GetType(String))
            sortedData.Columns.Add("lastName", GetType(String))
            sortedData.Columns.Add("middleInitial", GetType(String))
            sortedData.Columns.Add("suffix", GetType(String))
            sortedData.Columns.Add("section", GetType(String))
            sortedData.Columns.Add("email", GetType(String))
            sortedData.Columns.Add("status", GetType(String))

            ' Get the sorted data from the DataGridView
            For Each row As DataGridViewRow In studentView.Rows
                If Not row.IsNewRow Then
                    Dim newRow As DataRow = sortedData.NewRow()
                    newRow("studentNumber") = If(row.Cells("studentNumber").Value IsNot Nothing, row.Cells("studentNumber").Value.ToString(), "")
                    newRow("firstName") = If(row.Cells("firstName").Value IsNot Nothing, row.Cells("firstName").Value.ToString(), "")
                    newRow("lastName") = If(row.Cells("lastName").Value IsNot Nothing, row.Cells("lastName").Value.ToString(), "")
                    newRow("middleInitial") = If(row.Cells("middleInitial").Value IsNot Nothing, row.Cells("middleInitial").Value.ToString(), "")
                    newRow("suffix") = If(row.Cells("suffix").Value IsNot Nothing, row.Cells("suffix").Value.ToString(), "")
                    newRow("section") = If(row.Cells("section").Value IsNot Nothing, row.Cells("section").Value.ToString(), "")
                    newRow("email") = If(row.Cells("email").Value IsNot Nothing, row.Cells("email").Value.ToString(), "")
                    newRow("status") = If(row.Cells("status").Value IsNot Nothing, row.Cells("status").Value.ToString(), "")
                    sortedData.Rows.Add(newRow)
                End If
            Next

            Return sortedData
        Catch ex As Exception
            MessageBox.Show("Failed to get sorted student data: " & ex.Message, "Error")
            Return Nothing
        End Try
    End Function

    Private Sub generateReportForm_Click(sender As Object, e As EventArgs) Handles generateReportForm.Click
        Try
            ' Get report data from the DataGridView (respects sorting)
            Dim reportData As DataTable = GetSortedReportDataFromGridView()
            If reportData Is Nothing OrElse reportData.Rows.Count = 0 Then
                MessageBox.Show("No report data available to generate report.", "Information")
                Return
            End If

            ' Setup save file dialog
            Using saveFileDialog As New SaveFileDialog()
                saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
                saveFileDialog.Title = "Save Consultation Report"
                saveFileDialog.FileName = "Consultation_Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"

                If saveFileDialog.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                ' Create PDF document
                Using document As New PdfDocument()
                    document.Info.Title = "Consultation Report"

                    ' Setup fonts
                    Dim fontTitle As New XFont("Segoe UI", 16, XFontStyle.Bold)
                    Dim fontSubTitle As New XFont("Segoe UI", 9, XFontStyle.Regular)
                    Dim fontHeader As New XFont("Segoe UI", 9, XFontStyle.Bold)
                    Dim fontRow As New XFont("Segoe UI", 8, XFontStyle.Regular)

                    ' === Load left logo from relative path ===
                    Dim leftLogoPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Resources\plplogo.png")
                    Dim leftLogo As XImage = If(File.Exists(leftLogoPath), XImage.FromFile(leftLogoPath), Nothing)

                    ' === Load right logo from settings ===
                    Dim rightLogoPath As String = My.Settings.LogoPath
                    Dim rightLogo As XImage = If(File.Exists(rightLogoPath), XImage.FromFile(rightLogoPath), Nothing)

                    ' Get department name with fallback
                    Dim departmentName As String = If(String.IsNullOrEmpty(My.Settings.DepartmentName), "Not Specified", My.Settings.DepartmentName)

                    Dim DrawHeader = Sub(hdrGfx As XGraphics, hdrPage As PdfPage)
                                         ' Background rectangle
                                         hdrGfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(245, 245, 245)), 0, 0, hdrPage.Width.Point, 80)

                                         ' Top-left text: small, single line
                                         Dim xLeft As Double = 20
                                         Dim yTop As Double = 5  ' small margin from top
                                         Dim smallFont As New XFont("Segoe UI", 7, XFontStyle.Regular)

                                         Dim adminName As String = "Administrator"
                                         If CurrentAdmin IsNot Nothing AndAlso Not String.IsNullOrEmpty(CurrentAdmin.LastName) Then
                                             adminName = CurrentAdmin.LastName.ToUpper() ' optional uppercase
                                         End If

                                         Dim mi As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.MiddleInitial), "", " " & CurrentAdmin.MiddleInitial & ".")
                                         Dim suffix As String = If(String.IsNullOrWhiteSpace(CurrentAdmin.Suffix), "", " " & CurrentAdmin.Suffix)
                                         Dim fullName As String = $"{CurrentAdmin.FirstName}{mi} {CurrentAdmin.LastName}{suffix}".Trim()

                                         Dim generatedText As String = $"Report generated by Admin Prof. {fullName} as of {DateTime.Now:MMMM dd, yyyy, 'at' hh:mm tt}"

                                         hdrGfx.DrawString(generatedText, smallFont, XBrushes.Black, New XRect(xLeft, yTop, 400, 10), XStringFormats.TopLeft)


                                         Dim targetHeight As Double = 50

                                         If leftLogo IsNot Nothing Then
                                             Dim aspectLeft As Double = leftLogo.PixelWidth / leftLogo.PixelHeight
                                             Dim newWidthLeft As Double = targetHeight * aspectLeft
                                             hdrGfx.DrawImage(leftLogo, xLeft, yTop + 15, newWidthLeft, targetHeight)
                                         End If

                                         If rightLogo IsNot Nothing Then
                                             Dim aspectRight As Double = rightLogo.PixelWidth / rightLogo.PixelHeight
                                             Dim newWidthRight As Double = targetHeight * aspectRight
                                             hdrGfx.DrawImage(rightLogo, hdrPage.Width.Point - newWidthRight - 20, yTop + 15, newWidthRight, targetHeight)
                                         End If

                                         ' Centered text below logos
                                         Dim centerX As Double = hdrPage.Width.Point / 2
                                         Dim yCenter As Double = yTop + 20  ' adjust vertical position below small top text

                                         hdrGfx.DrawString("Consultation Report", fontTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter, 300, 20), XStringFormats.TopCenter)
                                         hdrGfx.DrawString("Pamantasan ng Lungsod ng Pasig", fontSubTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter + 20, 300, 15), XStringFormats.TopCenter)
                                         hdrGfx.DrawString(departmentName, fontSubTitle, XBrushes.Black,
                                                           New XRect(centerX - 150, yCenter + 35, 300, 15), XStringFormats.TopCenter)
                                     End Sub



                    ' Create first page
                    Dim page As PdfPage = document.AddPage()
                    page.Size = PageSize.A4
                    Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                    Dim tf As New XTextFormatter(gfx)
                    DrawHeader(gfx, page)

                    Dim yPoint As Double = 90 ' <- Slight margin before the table
                    Dim marginLeft As Integer = 10
                    Dim lineHeight As Double = 12
                    ' Column widths for 9 columns (adjusted for portrait with text wrapping)
                    Dim colWidths() As Integer = {60, 70, 50, 70, 60, 130, 50, 40, 40} ' Adjusted widths
                    Dim headers() As String = {"Student No.", "Name", "Section", "Professor", "Reason", "Message", "Date", "Time In", "Time Out"}

                    Dim xPos As Double = marginLeft
                    For i = 0 To headers.Length - 1
                        gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                        gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopLeft)
                        xPos += colWidths(i)
                    Next
                    yPoint += lineHeight + 6

                    Dim rowColorToggle As Boolean = False

                    For Each row As DataRow In reportData.Rows
                        xPos = marginLeft
                        Dim maxRowHeight As Double = 0
                        Dim rowData() As String = {
                        row("studentNumberCol").ToString(),
                        row("studentNameCol").ToString(),
                        row("studentSectionCol").ToString(),
                        row("professorNameCol").ToString(),
                        row("reasonCol").ToString(),
                        ForceWrap(row("messageCol").ToString()), ' Use ForceWrap for message column
                        row("dateCol").ToString(),
                        row("timeInCol").ToString(),
                        row("timeOutCol").ToString()
                    }

                        ' PROPER MEASUREMENT - This is what makes text wrapping work!
                        For i = 0 To rowData.Length - 1
                            Dim layoutRect As New XRect(0, 0, colWidths(i) - 4, Double.MaxValue)
                            Dim dummyGfx As XGraphics = XGraphics.CreateMeasureContext(New XSize(colWidths(i) - 4, Double.MaxValue), XGraphicsUnit.Point, XPageDirection.Downwards)
                            Dim size As XSize = dummyGfx.MeasureString(rowData(i), fontRow)
                            Dim linesNeeded As Integer = Math.Ceiling(size.Width / layoutRect.Width)
                            Dim heightNeeded As Double = linesNeeded * lineHeight
                            If heightNeeded > maxRowHeight Then maxRowHeight = heightNeeded
                        Next

                        If maxRowHeight < lineHeight * 2 Then maxRowHeight = lineHeight * 2

                        If rowColorToggle Then
                            gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(240, 240, 240)), marginLeft, yPoint, colWidths.Sum(), maxRowHeight)
                        End If

                        xPos = marginLeft
                        For i = 0 To rowData.Length - 1
                            tf.DrawString(rowData(i), fontRow, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, maxRowHeight), XStringFormats.TopLeft)
                            gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), maxRowHeight)
                            xPos += colWidths(i)
                        Next

                        yPoint += maxRowHeight
                        rowColorToggle = Not rowColorToggle

                        If yPoint > page.Height.Point - 40 Then
                            page = document.AddPage()
                            page.Size = PageSize.A4
                            gfx = XGraphics.FromPdfPage(page)
                            tf = New XTextFormatter(gfx)
                            DrawHeader(gfx, page)

                            yPoint = 90 ' Maintain top margin
                            xPos = marginLeft
                            For i = 0 To headers.Length - 1
                                gfx.DrawRectangle(XPens.LightGray, xPos, yPoint, colWidths(i), lineHeight + 4)
                                gfx.DrawString(headers(i), fontHeader, XBrushes.Black, New XRect(xPos + 2, yPoint + 2, colWidths(i) - 4, lineHeight), XStringFormats.TopLeft)
                                xPos += colWidths(i)
                            Next
                            yPoint += lineHeight + 6
                        End If
                    Next

                    document.Save(saveFileDialog.FileName)
                    MessageBox.Show("Consultation Report generated successfully!", "Success")
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Failed to generate PDF: " & ex.Message, "PDF Error")
        End Try
    End Sub


    ' Function to get sorted report data from DataGridView
    Private Function GetSortedReportDataFromGridView() As DataTable
        Try
            ' Create a new DataTable with the same structure
            Dim sortedData As New DataTable()
            sortedData.Columns.Add("studentNumberCol", GetType(String))
            sortedData.Columns.Add("studentNameCol", GetType(String))
            sortedData.Columns.Add("studentSectionCol", GetType(String))
            sortedData.Columns.Add("professorNameCol", GetType(String))
            sortedData.Columns.Add("reasonCol", GetType(String))
            sortedData.Columns.Add("messageCol", GetType(String))
            sortedData.Columns.Add("dateCol", GetType(String))
            sortedData.Columns.Add("timeInCol", GetType(String))
            sortedData.Columns.Add("timeOutCol", GetType(String))

            ' Get the sorted data from the DataGridView
            For Each row As DataGridViewRow In reportView.Rows
                If Not row.IsNewRow Then
                    Dim newRow As DataRow = sortedData.NewRow()
                    newRow("studentNumberCol") = If(row.Cells("studentNumberCol").Value IsNot Nothing, row.Cells("studentNumberCol").Value.ToString(), "")
                    newRow("studentNameCol") = If(row.Cells("studentNameCol").Value IsNot Nothing, row.Cells("studentNameCol").Value.ToString(), "")
                    newRow("studentSectionCol") = If(row.Cells("studentSectionCol").Value IsNot Nothing, row.Cells("studentSectionCol").Value.ToString(), "")
                    newRow("professorNameCol") = If(row.Cells("professorNameCol").Value IsNot Nothing, row.Cells("professorNameCol").Value.ToString(), "")
                    newRow("reasonCol") = If(row.Cells("reasonCol").Value IsNot Nothing, row.Cells("reasonCol").Value.ToString(), "")
                    newRow("messageCol") = If(row.Cells("messageCol").Value IsNot Nothing, row.Cells("messageCol").Value.ToString(), "")
                    newRow("dateCol") = If(row.Cells("dateCol").Value IsNot Nothing, row.Cells("dateCol").Value.ToString(), "")
                    newRow("timeInCol") = If(row.Cells("timeInCol").Value IsNot Nothing, row.Cells("timeInCol").Value.ToString(), "")
                    newRow("timeOutCol") = If(row.Cells("timeOutCol").Value IsNot Nothing, row.Cells("timeOutCol").Value.ToString(), "")
                    sortedData.Rows.Add(newRow)
                End If
            Next

            Return sortedData
        Catch ex As Exception
            MessageBox.Show("Failed to get sorted report data: " & ex.Message, "Error")
            Return Nothing
        End Try
    End Function
    Private Sub LoadHiddenProfessorsToGrid()
        Dim dt As DataTable = GetHiddenProfessors()

        hiddenView.Rows.Clear()

        For Each row As DataRow In dt.Rows
            hiddenView.Rows.Add(
            row("first_name").ToString(),
            row("last_name").ToString(),
            row("middle_initial").ToString(),
            row("suffix").ToString(),
            row("email").ToString(),
            row("username").ToString(),
            row("password").ToString(),
            row("id").ToString() ' hiddenIdCol (if you need it)
        )
        Next
    End Sub

    ' Function to get all hidden professors
    Public Function GetHiddenProfessors() As DataTable
        Try
            Connect()

            Dim query As String = "SELECT * FROM professors WHERE isHidden = 1"
            Dim cmd As New MySqlCommand(query, conn)

            Dim dt As New DataTable()
            Using adapter As New MySqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using

            Return dt

        Catch ex As Exception
            MessageBox.Show("Failed to retrieve hidden professors: " & ex.Message, "Query Error")
            Return New DataTable()
        End Try
    End Function
    Private Sub professorArchiveBtn_Click(sender As Object, e As EventArgs) Handles professorArchiveBtn.Click
        ' Check if a professor is selected
        If selectedProfessorId <= 0 Then
            MessageBox.Show("Please select a professor to archive.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Confirm the archive action
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to archive this professor?", "Confirm Archive",
                                                MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.No Then
            Return
        End If

        Try
            Dim query As String = $"UPDATE professors SET isHidden = 1 WHERE id = {selectedProfessorId}"
            ExecuteNonQuery(query)

            MessageBox.Show("Professor archived successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Refresh the data grid view to remove the archived professor
            LoadProfessorsToGrid()

            ' Clear the form fields
            ClearInputs(adminProfessorPanel)

            ' Reset the selected professor ID
            selectedProfessorId = 0

        Catch ex As Exception
            MessageBox.Show("Error archiving professor: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private selectedHiddenProfessorId As Integer

    Private Sub hiddenView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles hiddenView.CellContentClick
        If e.RowIndex >= 0 AndAlso e.RowIndex < hiddenView.Rows.Count Then
            Dim row As DataGridViewRow = hiddenView.Rows(e.RowIndex)

            ' Fill the text fields from the selected row
            hiddenFirstInput.Text = If(row.Cells(0).Value Is Nothing, "", row.Cells(0).Value.ToString())
            hiddenLastInput.Text = If(row.Cells(1).Value Is Nothing, "", row.Cells(1).Value.ToString())
            hiddenMiddleInput.Text = If(row.Cells(2).Value Is Nothing, "", row.Cells(2).Value.ToString())

            ' Handle suffix in combobox
            Dim suffixValue As String = If(row.Cells(3).Value Is Nothing, "", row.Cells(3).Value.ToString().Trim())
            If String.IsNullOrEmpty(suffixValue) OrElse Not hiddenSuffixBox.Items.Contains(suffixValue) Then
                hiddenSuffixBox.SelectedIndex = -1
            Else
                hiddenSuffixBox.SelectedItem = suffixValue
            End If

            hiddenEmailInput.Text = If(row.Cells(4).Value Is Nothing, "", row.Cells(4).Value.ToString())
            hiddenUsernameInput.Text = If(row.Cells(5).Value Is Nothing, "", row.Cells(5).Value.ToString())
            hiddenPasswordInput.Text = If(row.Cells(6).Value Is Nothing, "", row.Cells(6).Value.ToString())

            ' Store the hidden professor ID (from column 7)
            selectedHiddenProfessorId = Convert.ToInt32(row.Cells(7).Value)
        End If
    End Sub

    Private Sub unarchiveProfBtn_Click(sender As Object, e As EventArgs) Handles unarchiveProfBtn.Click
        ' Check if a hidden professor is selected
        If selectedHiddenProfessorId <= 0 Then
            MessageBox.Show("Please select a professor to unarchive.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Confirm the unarchive action
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to unarchive this professor?", "Confirm Unarchive",
                                                MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.No Then
            Return
        End If

        Try
            Dim query As String = $"UPDATE professors SET isHidden = 0 WHERE id = {selectedHiddenProfessorId}"
            ExecuteNonQuery(query)

            MessageBox.Show("Professor unarchived successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Refresh both grid views
            LoadHiddenProfessorsToGrid()  ' Refresh hidden professors list


            ' Clear the hidden professor form fields
            ClearInputs(adminHidePanel)
            ' Reset the selected hidden professor ID
            selectedHiddenProfessorId = 0

        Catch ex As Exception
            MessageBox.Show("Error unarchiving professor: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub professorSearch2_TextChanged(sender As Object, e As EventArgs)
        ' Reload the report every time the text changes
        LoadReports()
    End Sub

    Private Sub professorDateText_Click(sender As Object, e As EventArgs) Handles professorDateText.Click

    End Sub

    Private Sub professorBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles professorBox2.SelectedIndexChanged
        ' Prevent firing when DataSource is being set (during LoadProfessorsToComboBox3)
        If professorBox2.Focused OrElse professorBox2.SelectedIndex >= 0 Then
            LoadReports()
        End If
    End Sub


    ' Helper method to clear the hidden professor form fields

End Class
