Imports System.IO
Imports System.Environment
Imports System.Text.RegularExpressions

Public Class Form1
    Dim NF_MinPassWord_Leg As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim CmdStr As String
        CmdStr = "/export /cfg " + GetFolderPath(SpecialFolder.ApplicationData) + "\cfg.inf"
        Dim myProcess As Process
        myProcess = New Process()
        myProcess.StartInfo.FileName = "C:\Windows\System32\SecEdit.exe"
        myProcess.StartInfo.Arguments = CmdStr
        myProcess.Start()
        myProcess.WaitForExit()
        'Threading.Thread.Sleep(10000)
        Dim EntireFile As String
        'Dim oFile As System.IO.File
        Dim oRead As StreamReader
        Try
            oRead = File.OpenText(GetFolderPath(SpecialFolder.ApplicationData) + "\cfg.inf")

            EntireFile = oRead.ReadToEnd()
            TextBox1.Text = (EntireFile)
            oRead.Dispose()
            oRead.Close()

            PassWordComplexity()

        GuestAccountSet()

        setVal(New Regex("MinimumPasswordAge = -?\d+"), 1)

        setVal(New Regex("MaximumPasswordAge = -?\d+"), 2)

        setVal(New Regex("MinimumPasswordLength = -?\d+"), 3)

            setVal(New Regex("PasswordHistorySize = -?\d+"), 4)

        Catch ex As Exception
            MsgBox("Something went wrong with secedit making the cfg file try and rerun the program.")
        End Try
    End Sub
    Public Sub PassWordComplexity()
        If TextBox1.Text.Contains("PasswordComplexity = 0") Then
            RadioButton2.Checked = True
        ElseIf TextBox1.Text.Contains("PasswordComplexity = 1") Then
            RadioButton1.Checked = True
        Else
            MsgBox("Password Complexity Setting Cannot be found.")
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
        End If
    End Sub

    Public Sub GuestAccountSet()
        If TextBox1.Text.Contains("EnableGuestAccount = 0") Then
            RadioButton4.Checked = True
        ElseIf TextBox1.Text.Contains("EnableGuestAccount = 1") Then
            RadioButton3.Checked = True
        Else
            MsgBox("Guest account settings cannot be found.")
            RadioButton4.Enabled = False
            RadioButton3.Enabled = False
        End If
    End Sub

    Public Sub setVal(ByVal rega As Regex, ByVal numupdw As Integer)
        Dim pwdMemb = New Regex("-?\d+")
        Try
            Dim matchPwdMema As MatchCollection = rega.Matches(TextBox1.Text)
            Dim matchPwdMemb As MatchCollection = pwdMemb.Matches(matchPwdMema(0).ToString)
            Dim PwdMem As Integer = Convert.ToInt32(matchPwdMemb(0).ToString)
            If numupdw = 1 Then
                NumericUpDown1.Value = PwdMem
            ElseIf numupdw = 2 Then
                NumericUpDown2.Value = PwdMem
            ElseIf numupdw = 3 Then
                NumericUpDown3.Value = PwdMem
            ElseIf numupdw = 4 Then
                NumericUpDown4.Value = PwdMem
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        'ON PASSWORD COMPLEXITY
        TextBox1.Text = TextBox1.Text.Replace("PasswordComplexity = 0", "PasswordComplexity = 1")
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        'OFF PASSWORD COMPLEXITY
        TextBox1.Text = TextBox1.Text.Replace("PasswordComplexity = 1", "PasswordComplexity = 0")
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        'ON GUEST ACCOUNT
        TextBox1.Text = TextBox1.Text.Replace("EnableGuestAccount = 0", "EnableGuestAccount = 1")
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        'OFF GUEST ACCOUNT
        TextBox1.Text = TextBox1.Text.Replace("EnableGuestAccount = 1", "EnableGuestAccount = 0")
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        'NUMERIC UPDOWN FOR MIN PASSWORD AGE
        Dim MinAgea = New Regex("MinimumPasswordAge = -?\d+")
        Dim matchMinAgea As MatchCollection = MinAgea.Matches(TextBox1.Text)
        Try
            TextBox1.Text = TextBox1.Text.Replace(matchMinAgea(0).ToString, "MinimumPasswordAge = " + CStr(NumericUpDown1.Value))
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        'NUMERIC UPDOWN FOR MAX PASSWORD AGE
        Dim MaxAgea = New Regex("MaximumPasswordAge = -?\d+")
        Dim matchMaxAgea As MatchCollection = MaxAgea.Matches(TextBox1.Text)
        Try
            TextBox1.Text = TextBox1.Text.Replace(matchMaxAgea(0).ToString, "MaximumPasswordAge = " + CStr(NumericUpDown2.Value))
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        'NUMERIC UPDOWN FOR MIN PASSWORD LENGTH
        Dim MinLega = New Regex("MinimumPasswordLength = -?\d+")
        Dim matchMinLega As MatchCollection = MinLega.Matches(TextBox1.Text)
        Try
            TextBox1.Text = TextBox1.Text.Replace(matchMinLega(0).ToString, "MinimumPasswordLength = " + CStr(NumericUpDown3.Value))
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        'Numeric updown for pwd memory
        Dim PwdMema = New Regex("PasswordHistorySize = -?\d+")
        Dim matchPwdMema As MatchCollection = PwdMema.Matches(TextBox1.Text)
        Try
            TextBox1.Text = TextBox1.Text.Replace(matchPwdMema(0).ToString, "PasswordHistorySize = " + CStr(NumericUpDown4.Value))
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    'Dim FILE_NAME As String = GetFolderPath(SpecialFolder.ApplicationData) + "\ncfg.inf"
    Dim FILE_NAME As String = """C:\ncfg.inf"""
    Dim FILE_NAMEA As String = "C:\ncfg.inf"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If File.Exists(FILE_NAMEA) Then
            File.Delete(FILE_NAMEA)
        End If
        File.Create(FILE_NAMEA).Dispose()
        Dim objWriter As New StreamWriter(FILE_NAMEA)
        objWriter.Write(TextBox1.Text)
        objWriter.Close()
        Button2.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim CmdStr As String
            CmdStr = "/configure /db secedit.sdb /cfg " + FILE_NAME
            Dim myProcess As Process
            myProcess = New Process()
            myProcess.StartInfo.FileName = "C:\Windows\System32\SecEdit.exe"
            myProcess.StartInfo.Arguments = CmdStr
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.Start()
            myProcess.WaitForExit()
            CmdStr = "/refreshpolicy machine_policy /enforce /quiet"

            myProcess = New Process()
            myProcess.StartInfo.FileName = "C:\Windows\System32\SecEdit.exe"
            myProcess.StartInfo.Arguments = CmdStr
            myProcess.Start()
            myProcess.WaitForExit()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Threading.Thread.Sleep(2000)
        'Application.Restart()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("https://github.com/ndragon798")
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            File.Delete(FILE_NAMEA)
            File.Delete(GetFolderPath(SpecialFolder.ApplicationData) + "\cfg.inf")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
End Class
