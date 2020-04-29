Option Strict On
Option Explicit On
Public Class frmOSMSmart
    Private pobjPNR As PNR
    Private Sub cmdRead_Click(sender As Object, e As EventArgs) Handles cmdRead.Click

        Try
            pobjPNR = New PNR(txtInput.Text)
            Dim i As Integer = 1
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
