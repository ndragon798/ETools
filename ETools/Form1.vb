Imports System.IO
Public Class Form1
    Dim NF_MinPassWord_Leg As String
    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim CmdStr As String
        CmdStr = "/export /cfg c:\cfg.txt"
        Dim myProcess As System.Diagnostics.Process
        myProcess = New System.Diagnostics.Process()
        myProcess.StartInfo.FileName = "C:\Windows\System32\SecEdit.exe"
        myProcess.StartInfo.Arguments = CmdStr
        myProcess.Start()
        Threading.Thread.Sleep(2000)
        Dim EntireFile As String
        Dim oFile As System.IO.File
        Dim oRead As System.IO.StreamReader
        oRead = File.OpenText("C:\cfg.txt")

        EntireFile = oRead.ReadToEnd()
        TextBox1.Text = (EntireFile)

        'PASSWORD COMPLEXTIY SET
        If TextBox1.Text.Contains("PasswordComplexity = 0") Then
            RadioButton2.Checked = True
        ElseIf TextBox1.Text.Contains("PasswordComplexity = 1") Then
            RadioButton1.Checked = True
        End If
        '/END PASSWORD COMPLETIXTY SET

        'GUESTACCOUNT SET
        If TextBox1.Text.Contains("EnableGuestAccount = 0") Then
            RadioButton4.Checked = True
        ElseIf TextBox1.Text.Contains("EnableGuestAccount = 1") Then
            RadioButton3.Checked = True
        End If
        '/END GUEST ACCOUNT SET

        'MIN PASSWORD LENGTH SET 
        Dim MinLeg As String
        Dim MinLegpp As String
        For MinLegp As Integer = -1000 To 1000

            MinLeg = "MinimumPasswordLength = "
            MinLegpp = CStr(MinLegp)
            MinLeg = MinLeg & MinLegpp
            If TextBox1.Text.Contains(MinLeg) Then
                NumericUpDown3.Value = MinLegp
            End If
        Next
        Dim NumUpDown3 As Integer = NumericUpDown3.Value
        NF_MinPassWord_Leg = "MinimumPasswordLength = " + CStr(NumUpDown3)
        'Label6.Text = NF_MinPassWord_Leg

        '/END MIN PASSWORD LENGTHSET

        'MIN PASSWORD AGE SET
        Dim MinAge As String
        For MinAgep As Integer = -1000 To 1000
            MinAge = "MinimumPasswordAge = "
            MinAge = MinAge & CStr(MinAgep)
            If TextBox1.Text.Contains(MinAge) Then
                If MinAgep < 9 & MinAgep > 0 Then
                    For a As Integer = 0 To 9
                        If TextBox1.Text.Contains(MinAge & CStr(a)) Then
                            NumericUpDown1.Value = Convert.ToInt32(CStr(MinAgep) & CStr(a))
                        Else
                            NumericUpDown1.Value = MinAgep
                        End If
                    Next
                    NumericUpDown1.Value = MinAgep
                End If


            End If

        Next
        '/END MIN PASSWORD AGE SET
        'MAX PASSWORD AGE SET
        Dim MaxAge As String
        For MaxAgep As Integer = -1000 To 1000
            MaxAge = "MaximumPasswordAge = "
            MaxAge = MaxAge & CStr(MaxAgep)
            If TextBox1.Text.Contains(MaxAge) Then
                NumericUpDown2.Value = MaxAgep
            End If

        Next
        '/END MAX PASSWORD AGE SET

        'PASSWORDS REMEMBERED SET
        Dim PasRem As String
        For PasRemp As Integer = -1000 To 1000
            PasRem = "PasswordHistorySize = "
            PasRem = PasRem & CStr(PasRemp)
            If TextBox1.Text.Contains(PasRem) Then
                NumericUpDown4.Value = PasRemp
            End If

        Next
        '/END PASSWORDS REMEMBERED SET
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        'ON PASSWORD COMPLEXITY
        TextBox1.Text = TextBox1.Text.Replace("PasswordComplexity = 0", "PasswordComplexity = 1")
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        'OFF PASSWORD COMPLEXITY
        TextBox1.Text = TextBox1.Text.Replace("PasswordComplexity = 1", "PasswordComplexity = 0")
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        'OFF GUEST ACCOUNT
        TextBox1.Text = TextBox1.Text.Replace("EnableGuestAccount = 1", "EnableGuestAccount = 0")
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        'ON GUEST ACCOUNT
        TextBox1.Text = TextBox1.Text.Replace("EnableGuestAccount = 0", "EnableGuestAccount = 1")
    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        'NUMERIC UPDOWN FOR MIN PASSWORD LENGTH
        Dim l As String

        For x As Integer = -1000 To 1000
            l = "MinimumPasswordLength = " + CStr(x)

            If TextBox1.Text.Contains(l) Then
                TextBox1.Text = TextBox1.Text.Replace(l, "MinimumPasswordLength = " + CStr(NumericUpDown3.Value))
            End If

        Next

    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        'NUMERIC UPDOWN FOR MIN PASSWORD AGE
        Dim l As String

        For x As Integer = -1000 To 1000
            l = "MinimumPasswordAge = " + CStr(x)

            If TextBox1.Text.Contains(l) Then
                TextBox1.Text = TextBox1.Text.Replace(l, "MinimumPasswordAge = " + CStr(NumericUpDown1.Value))
            End If

        Next
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        'NUMERIC UPDWON FOR MAX PASSWORD AGE
        Dim l As String

        For x As Integer = -1000 To 1000
            l = "MaximumPasswordAge = " + CStr(x)

            If TextBox1.Text.Contains(l) Then
                TextBox1.Text = TextBox1.Text.Replace(l, "MaximumPasswordAge = " + CStr(NumericUpDown2.Value))
            End If

        Next
    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        Dim l As String

        For x As Integer = -1000 To 1000
            l = "PasswordHistorySize = " + CStr(x)

            If TextBox1.Text.Contains(l) Then
                TextBox1.Text = TextBox1.Text.Replace(l, "PasswordHistorySize = " + CStr(NumericUpDown4.Value))
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        File.Create("C:\ncfg.inf").Dispose()
        Dim FILE_NAME As String = "c:\ncfg.inf"
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
        objWriter.Write(TextBox1.Text)
        objWriter.Close()
        Button2.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim CmdStr As String
        CmdStr = "/configure /db secedit.sdb /cfg ""c:\ncfg.inf"" /quiet"
        Dim myProcess As System.Diagnostics.Process
        myProcess = New System.Diagnostics.Process()
        myProcess.StartInfo.FileName = "C:\Windows\System32\SecEdit.exe"
        myProcess.StartInfo.Arguments = CmdStr
        myProcess.Start()

        CmdStr = "/refreshpolicy machine_policy /enforce /quiet"

        myProcess = New System.Diagnostics.Process()
        myProcess.StartInfo.FileName = "C:\Windows\System32\SecEdit.exe"
        myProcess.StartInfo.Arguments = CmdStr
        myProcess.Start()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("http://www.nathaneaston.com/ETools/")
    End Sub
End Class
