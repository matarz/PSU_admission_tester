Public Class Main
    Dim Lines(,) As Integer = {{19, 0}, {21, 0}, {27, 0}, {29, 0}, {35, 0}, {37, 0}, {39, 0}, {44, 0}, {46, 0}, {48, 0}, {54, 0}, {60, 0}, {62, 0}, {67, 0}, {72, 0}, {73, 0}, {78, 0}, {79, 0}, {84, 0}}

    Dim CurrentYear As Integer = 2011       'to calculate time since graduation

    'array to hold GPA values under 3
    Dim Under3GPA() As String = {2.99, 2.98, 2.97, 2.96, 2.95, 2.94, 2.93, 2.92, 2.91, 2.9, 2.89, 2.88, 2.87, 2.86, 2.85, 2.84, 2.83, 2.82, 2.81, 2.8, 2.79, 2.78, 2.77, 2.76, 2.75, 2.74, 2.73, 2.72, 2.71, 2.7}

    'array to hold SAT values for GPA values under 3
    Dim Under3GPA_SAT() As String = {1060, 1070, 1080, 1080, 1090, 1100, 1110, 1120, 1120, 1130, 1140, 1150, 1160, 1160, 1170, 1180, 1190, 1200, 1200, 1210, 1220, 1230, 1230, 1240, 1250, 1260, 1270, 1280, 1280, 1290}

    'array to hold ACT values for GPA values under 3
    Dim Under3GPA_ACT() As String = {23, 23, 23, 23, 24, 24, 24, 24, 24, 25, 25, 25, 25, 25, 26, 26, 26, 26, 26, 27, 27, 27, 27, 27, 28, 28, 28, 28, 28, 29}

    Private Sub CheckButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckButton.Click

        'Check if year, gpa, and units are valid inputs
        Lines(0, 1) = 1 'LC
        If Not PosInt(YearOfGraduationBox.Text) Or Not PosInt(GPABox.Text) Or Not PosInt(EnglishBox.Text) Or Not PosInt(MathBox.Text) Or Not PosInt(ScienceBox.Text) Or Not PosInt(SocialStudiesBox.Text) Or Not PosInt(LanguageBox.Text) Then
            Lines(1, 1) = 1 'LC
            MsgBox("Field value not valid")
            Exit Sub
        End If

        'check gpa and units for accepted rates
        Lines(2, 1) = 1 'LC
        If CInt(GPABox.Text) < 2.7 Or CInt(EnglishBox.Text) < 4 Or CInt(MathBox.Text) < 3 Or CInt(ScienceBox.Text) < 3 Or CInt(SocialStudiesBox.Text) < 3 Or CInt(LanguageBox.Text) < 2 Then
            Lines(3, 1) = 1 'LC
            DecisionBox.Text = "Denied"
            Exit Sub
        End If

        'if gpa is less than 3 but more than 2.69 then apply gpa & SAT & ACT calculations
        Lines(4, 1) = 1 'LC
        If GPABox.Text < 3 Then
            Lines(5, 1) = 1 'LC
            If Not PosInt(SATBox.Text) Or Not PosInt(ACTBox.Text) Then
                Lines(6, 1) = 1 'LC
                MsgBox("You need a SAT and ACT score with this GPA")
                Exit Sub
            End If

            Lines(7, 1) = 1 'LC
            For i = 0 To Under3GPA.Length - 1 Step 1
                Lines(8, 1) = 1 'LC
                If CDbl(GPABox.Text) = Under3GPA(i) And CInt(SATBox.Text) >= Under3GPA_SAT(i) And CInt(ACTBox.Text) >= Under3GPA_ACT(i) Then
                    Lines(9, 1) = 1 'LC
                    DecisionBox.Text = "Accepted"
                    Exit Sub
                End If
            Next

            Lines(10, 1) = 1 'LC
            DecisionBox.Text = "Denied"
            Exit Sub
        End If

        'if gpa is more than 3
        Lines(11, 1) = 1 'LC
        If CurrentYear - CInt(YearOfGraduationBox.Text) < 3 And (Not PosInt(SATBox.Text) Or Not PosInt(ACTBox.Text)) Then
            Lines(12, 1) = 1 'LC
            MsgBox("You are required to provide SAT and ACT Scores")
            Exit Sub
        End If

        Lines(13, 1) = 1 'LC
        DecisionBox.Text = "Accepted"
    End Sub

    Private Function PosInt(ByVal Number As String) As Boolean
        Lines(14, 1) = 1 'LC
        If Number.Length < 1 Then
            Lines(15, 1) = 1 'LC
            Return False
        End If

        Lines(16, 1) = 1 'LC
        If Not IsNumeric(Number) Or Number < 0 Then
            Lines(17, 1) = 1 'LC
            Return False
        End If

        Lines(18, 1) = 1 'LC
        Return True
    End Function

    Private Sub DoneButton_Click(sender As System.Object, e As System.EventArgs) Handles DoneButton.Click
        Dim LinesExecuted = 0
        Dim LinesNotExecuted = ""

        For i = 0 To Lines.GetLength(0) - 1 Step 1
            If Lines(i, 1) = 1 Then
                LinesExecuted += 1
            Else
                LinesNotExecuted += Lines(i, 0).ToString() + ", "
            End If
        Next

        Dim Coverage As Double = Math.Round(100 * (LinesExecuted / Lines.GetLength(0)), 2)
        If LinesNotExecuted.Length > 2 Then
            LinesNotExecuted = LinesNotExecuted.Remove(LinesNotExecuted.Length - 2)
        End If
        MsgBox("Coverage: " + Coverage.ToString() + "%" + Environment.NewLine + "Lines not exercised: " + LinesNotExecuted)
    End Sub
End Class
