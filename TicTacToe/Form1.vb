Public Class Form1

    Dim cal1 As New Calculator

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(0, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button1.Text = "O"
            Else
                Button1.Text = "X"
            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(1, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button2.Text = "O"
            Else
                Button2.Text = "X"
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(2, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button3.Text = "O"
            Else
                Button3.Text = "X"
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(3, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button4.Text = "O"
            Else
                Button4.Text = "X"
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(4, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button5.Text = "O"
            Else
                Button5.Text = "X"
            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(5, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button6.Text = "O"
            Else
                Button6.Text = "X"
            End If
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(6, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button7.Text = "O"
            Else
                Button7.Text = "X"
            End If
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(7, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button8.Text = "O"
            Else
                Button8.Text = "X"
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim tmp As Integer
        tmp = cal1.getCounter
        Dim chk As Integer = 0
        chk = cal1.addbox(8, tmp)
        If (chk = 0) Then
            If (tmp = 0) Then
                Button9.Text = "O"
            Else
                Button9.Text = "X"
            End If
        End If
    End Sub

    Private Sub Result_Click(sender As Object, e As EventArgs) Handles Result.Click

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Button1.Text = ""
        Button2.Text = ""
        Button3.Text = ""
        Button4.Text = ""
        Button5.Text = ""
        Button6.Text = ""
        Button7.Text = ""
        Button8.Text = ""
        Button9.Text = ""
        Result.Visible = False
        For i As Integer = 0 To 8
            cal1.setBox(i, 0)
        Next
        For i As Integer = 0 To 7
            cal1.setChecker(i, 0)
        Next
        cal1.setCounter(0)
    End Sub
End Class

Public Class Calculator
    Dim board(9) As Integer
    Dim check(8) As Integer
    Dim counter As Integer = 0

    Public Sub Game()

    End Sub

    Function getCounter()
        Return counter Mod 2
    End Function

    Function getBox(ByVal box As Integer)
        Return board(box)
    End Function

    Function setBox(ByVal box As Integer, ByVal numba As Integer)
        board(box) = numba
    End Function

    Function setCounter(ByVal numba As Integer)
        counter = 0
    End Function

    Function setChecker(ByVal box As Integer, ByVal numba As Integer)
        check(box) = numba
    End Function

    Function addbox(ByVal box As Integer, ByVal num1 As Integer)
        Dim num As Integer
        If (counter < 9 And board(box) = 0) Then
            counter = counter + 1
            If (num1 = 1) Then
                board(box) = -1
                num = -1
            Else
                board(box) = 1
                num = 1
            End If

            Select Case box
                Case 0
                    check(4) = check(4) + num
                    check(1) = check(1) + num
                    check(7) = check(7) + num
                Case 1
                    check(1) = check(1) + num
                    check(6) = check(6) + num
                Case 2
                    check(1) = check(1) + num
                    check(5) = check(5) + num
                    check(0) = check(0) + num
                Case 3
                    check(2) = check(2) + num
                    check(7) = check(7) + num
                Case 4
                    check(0) = check(0) + num
                    check(2) = check(2) + num
                    check(4) = check(4) + num
                    check(6) = check(6) + num
                Case 5
                    check(2) = check(2) + num
                    check(5) = check(5) + num
                Case 6
                    check(0) = check(0) + num
                    check(3) = check(3) + num
                    check(7) = check(7) + num
                Case 7
                    check(3) = check(3) + num
                    check(6) = check(6) + num
                Case 8
                    check(3) = check(3) + num
                    check(4) = check(4) + num
                    check(5) = check(5) + num
            End Select
            For i As Integer = 0 To 7
                If (check(i) >= 3 Or check(i) <= -3) Then
                    If (counter Mod 2 = 0) Then
                        Form1.Result.Text = "X Win!"
                        Form1.Result.Visible = True
                        counter = 10
                    Else
                        Form1.Result.Text = "O Win!"
                        Form1.Result.Visible = True
                        counter = 10
                    End If
                End If
            Next
            If (counter = 9) Then
                Form1.Result.Text = "Tie!"
                Form1.Result.Visible = True
                counter = 10
            End If
        Else
            Return 1
        End If
    End Function

End Class
