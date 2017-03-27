Imports System.IO

Public Class Customers

    Private Structure CustomersData

        Public FirstName As String
        Public LastName As String
        Public Age As String
        Public EmailAddress As String
        Public PhoneNumber As String

    End Structure

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim CustomersData() As String = File.ReadAllLines(Dir$("customerdetails.txt"))

        Dim CountNeeded As Integer
        Dim CountGot As Integer
        Dim DataValidated As Integer
        Dim Same As Integer

        CountNeeded = 0
        CountGot = 5
        Same = 0

        LengthValidate()

        For I = 0 To UBound(CustomersData)

            If Trim(Mid(CustomersData(I), 1, 30)) = txtFirstName.Text Then Same = Same + 1
            If Trim(Mid(CustomersData(I), 31, 30)) = txtLastName.Text Then Same = Same + 1
            If Trim(Mid(CustomersData(I), 61, 4)) = txtAge.Text Then Same = Same + 1
            If Trim(Mid(CustomersData(I), 65, 45)) = txtEmailAddress.Text Then Same = Same + 1
            If Trim(Mid(CustomersData(I), 110, 11)) = txtPhoneNumber.Text Then Same = Same + 1

        Next I

        If Same = 5 Then

            MsgBox("This data has already been saved.") : Exit Sub

        End If

        If Not txtFirstName.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtLastName.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtAge.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtEmailAddress.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtPhoneNumber.Text = "" Then CountNeeded = CountNeeded + 1

        If CountNeeded < 5 Then MsgBox("Please enter something to count!") : Exit Sub

        If CountGot = CountNeeded Then DataValidated = DataValidated + 1

        If DataValidated = 1 Then

            Dim CustomersPersonalData As New CustomersData
            Dim sw As New System.IO.StreamWriter(Dir$("customerdetails.txt"), True)
            CustomersPersonalData.FirstName = LSet(txtFirstName.Text, 30)
            CustomersPersonalData.LastName = LSet(txtLastName.Text, 30)
            CustomersPersonalData.Age = LSet(txtAge.Text, 4)
            CustomersPersonalData.EmailAddress = LSet(txtEmailAddress.Text, 45)
            CustomersPersonalData.PhoneNumber = LSet(txtPhoneNumber.Text, 11)

            sw.WriteLine(CustomersPersonalData.FirstName & CustomersPersonalData.LastName & CustomersPersonalData.Age & CustomersPersonalData.EmailAddress & CustomersPersonalData.PhoneNumber)
            sw.Close()

            Clear()

        End If

    End Sub

    Private Sub Customers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Dir$("customerdetails.txt") = "" Then

            Dim sw As New StreamWriter(Application.StartupPath & "\customerdetails.txt", True)

            sw.WriteLine("                                                                                                                                                                                                                                                                                                                                    ")
            sw.Close()

            MsgBox("A new database has been created", vbExclamation, "Warning!")

        End If

    End Sub

    Private Sub cmdRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRead.Click

        Dim CountNeeded As Integer
        Dim CountGot As Integer
        Dim ReadSuccess As Integer
        CountNeeded = 5
        ReadSuccess = 0

        Dim CustomersData() As String = File.ReadAllLines(Dir$("customerdetails.txt"))

        For I = 0 To UBound(CustomersData)

            CountGot = 0

            If Trim(Mid(CustomersData(I), 1, 30)) = txtFirstName.Text And Not txtFirstName.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(CustomersData(I), 31, 30)) = txtLastName.Text And Not txtLastName.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(CustomersData(I), 61, 4)) = txtAge.Text And Not txtAge.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(CustomersData(I), 65, 45)) = txtEmailAddress.Text And Not txtEmailAddress.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(CustomersData(I), 110, 11)) = txtPhoneNumber.Text And Not txtPhoneNumber.Text = "" Then CountGot = CountGot + 1

            If CountGot = CountNeeded Then ReadSuccess = ReadSuccess + 1

        Next I

        MsgBox(ReadSuccess & " Set(s) of customer data read.")

        Clear()

    End Sub

    Private Sub LengthValidate()

        If txtFirstName.Text.Length > 15 Then

            MsgBox("Too many characters in First Name.")

            Exit Sub

        End If

        If txtLastName.Text.Length > 15 Then

            MsgBox("Too many characters in Last Name.")

            Exit Sub

        End If

        If txtAge.Text.Length > 3 Then

            MsgBox("Too many characters in Age.")

            Exit Sub

        End If

        If txtEmailAddress.Text.Length > 30 Then

            MsgBox("Too many characters in Email Address.")

            Exit Sub

        End If

        If txtPhoneNumber.Text.Length > 11 Then

            MsgBox("Too many characters in Phone Number.")

            Exit Sub

        End If

    End Sub

    Private Sub Clear()

        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtAge.Text = ""
        txtEmailAddress.Text = ""
        txtPhoneNumber.Text = ""

    End Sub

End Class