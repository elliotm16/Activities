Imports System.IO

Public Class Details

    Private Structure LeisureData

        Public Type As String
        Public Session As String
        Public Day As String
        Public Location As String
        Public Level As String

    End Structure

    Private Sub cmdCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCount.Click

        Dim CountNeeded As Integer
        Dim CountGot As Integer
        Dim I As Integer
        Dim SessionCount As Integer

        SessionCount = 0
        CountNeeded = 0

        LengthValidate()

        If Not txtType.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtSession.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtDay.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtLocation.Text = "" Then CountNeeded = CountNeeded + 1
        If Not TxtLevel.Text = "" Then CountNeeded = CountNeeded + 1

        If CountNeeded = 0 Then

            MsgBox("Please enter something to count!")

            Exit Sub

        End If

        Dim Leisuredata() As String = File.ReadAllLines(Dir$("Leisure.txt"))

        For I = 0 To UBound(Leisuredata)

            CountGot = 0

            If Trim(Mid(Leisuredata(I), 1, 30)) = txtType.Text And Not txtType.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(Leisuredata(I), 31, 5)) = txtSession.Text And Not txtSession.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(Leisuredata(I), 36, 10)) = txtDay.Text And Not txtDay.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(Leisuredata(I), 46, 30)) = txtLocation.Text And Not txtLocation.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(Leisuredata(I), 76, 15)) = TxtLevel.Text And Not TxtLevel.Text = "" Then CountGot = CountGot + 1

            If CountGot = CountNeeded Then

                SessionCount = SessionCount + 1

            End If

        Next I

        MsgBox(SessionCount & " Sessions have been found! Contact the centre for more information.")

        txtType.Text = ""
        txtSession.Text = ""
        txtDay.Text = ""
        txtLocation.Text = ""
        TxtLevel.Text = ""

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim LeisureTimeData As New LeisureData
        Dim sw As New System.IO.StreamWriter(Dir$("Leisure.txt"), True)

        LengthValidate()

        LeisureTimeData.Type = LSet(txtType.Text, 30)
        LeisureTimeData.Session = LSet(txtSession.Text, 5)
        LeisureTimeData.Day = LSet(txtDay.Text, 10)
        LeisureTimeData.Location = LSet(txtLocation.Text, 30)
        LeisureTimeData.Level = LSet(TxtLevel.Text, 15)

        sw.WriteLine(LeisureTimeData.Type & LeisureTimeData.Session & LeisureTimeData.Day & LeisureTimeData.Location & LeisureTimeData.Level)
        sw.Close()

    End Sub

    Private Sub LeisureSessionsComplete_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Hide()
        Customers.Show()

        If Dir$("Leisure.txt") = "" Then

            Dim sw As New StreamWriter(Application.StartupPath & "\Leisure.txt", True)

            sw.WriteLine("                                                                                                                                                                                                                                                                                                                                    ")
            sw.Close()

            MsgBox("A new database has been created", vbExclamation, "Warning!")

        End If

    End Sub

    Private Sub LengthValidate()

        If txtType.Text.Length > 15 Then

            MsgBox("Too many characters in Type.")

            Exit Sub

        End If

        If txtSession.Text.Length > 15 Then

            MsgBox("Too many characters in Session.")

            Exit Sub

        End If

        If txtDay.Text.Length > 15 Then

            MsgBox("Too many characters in Day.")

            Exit Sub

        End If

        If txtLocation.Text.Length > 15 Then

            MsgBox("Too many characters in Location.")

            Exit Sub

        End If

        If TxtLevel.Text.Length > 15 Then

            MsgBox("Too many characters in Level.")

            Exit Sub

        End If

    End Sub

End Class