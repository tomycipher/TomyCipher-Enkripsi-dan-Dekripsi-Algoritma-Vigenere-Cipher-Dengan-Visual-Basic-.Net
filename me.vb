Public Class Form1
    Function Vigenere_Cipher(ByVal Text As String, ByVal key As String, ByVal Encrypt As Boolean)
        Dim Result As String = ""
        Dim temp As String = ""
        Dim j As Integer = 0

        For i As Integer = 0 To Text.Length - 1
            If j = key.Length Then
                j = 0
            End If

            If Char.IsLetter(key(j)) Then
                If Text(i) <> " " And Char.IsLetter(Text(i)) Then
                    temp += key(j)
                    j += 1
                Else
                    temp += Text(i)
                End If
            Else
                j += 1
                If j >= key.Length Then
                    j = 0
                End If
                i -= 1
            End If
        Next

        For i As Integer = 0 To Text.Length - 1

            Dim N As Integer
            Dim NewAscii As Integer

            If Char.IsLetter(Text(i)) Then
                If Char.IsLower(temp(i)) Then
                    N = Asc(temp(i)) - Asc("a")
                ElseIf Char.IsUpper(temp(i)) Then
                    N = Asc(temp(i)) - Asc("A")
                End If

                If Encrypt Then
                    NewAscii = N + Asc(Text(i))
                Else
                    NewAscii = 26 - N + Asc(Text(i))
                End If

                If (NewAscii > Asc("z") And Char.IsLower(Text(i))) Or (NewAscii > Asc("Z") And Char.IsUpper(Text(i))) Then
                    NewAscii -= 26
                End If
            Else
                NewAscii = Asc(Text(i))
            End If
            Result += Chr(NewAscii)
        Next
        Return Result
    End Function
    Dim x As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        x = txtKunci.Text
        txtEnkrip.Text = Vigenere_Cipher(txtPlain.Text, x, True)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        x = txtKunci.Text
        txtPlain.Text = Vigenere_Cipher(txtEnkrip.Text, x, False)
    End Sub
    'http://tomycipher.blogspot.com/2020/10/Algoritma%20Vigenere%20Cipher%20Dengan%20Visual%20Basic%20.Net.html
End Class
