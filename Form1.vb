Public Class Form1
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim orbon As New OrbonSeoul
        orbon.getShippingList(TextBox1.Text)



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ofd As New OpenFileDialog()
        Dim res As DialogResult
        With ofd
            .Filter = "폴더|\n"
            .InitialDirectory = TextBox1.Text
            .Title = "저장될 폴더 선택"
            .AddExtension = False
            .CheckFileExists = False
            .DereferenceLinks = False
            .Multiselect = False
            .FileName = "폴더 선택"
            .ValidateNames = False
            .CheckPathExists = True

            res = .ShowDialog
            If res = DialogResult.OK Then
                TextBox1.Text = .FileName.Substring(0, .FileName.LastIndexOf("\"))
            ElseIf res = DialogResult.Cancel Then
            End If
        End With

    End Sub
End Class
