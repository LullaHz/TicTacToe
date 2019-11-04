Public Class MainGame
    Dim board(9) As Boolean
    Dim check(8) As Integer
    Dim counter As Integer = 0


    Public Sub Game()

    End Sub

    Function getCounter()
        counter = (counter + 1) Mod 2
        Return counter
    End Function

    Function getBox(ByVal box As Integer)
        Return board(box)
    End Function

    Function addbox(ByVal box As Integer, ByVal num As Integer)
        If (board(box) = 0) Then
            If (num = 1) Then
                board(box) = 1
            Else
                board(box) = -1
            End If

            Select Case box
                Case 0
                    check(0) = check(0) + num
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
                    check(7) = check(7) + num
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
        End If
    End Function

End Class
