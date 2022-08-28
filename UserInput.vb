
Imports System.Windows.Forms

Public Class UserInput
    Dim inFileNameToRead As String
    Dim inPartsNumber As Integer


    Private Sub Button1OK_Click(sender As Object, e As EventArgs) Handles Button1OK.Click
        'While (txtPartName.Text Is Nothing Or txtHolesNum.Text Is Nothing Or txtMoveSide.Text Is Nothing Or txtMoveForFreeSafe.Text Is Nothing Or txtLengthForJumpToNextPort.Text)
        'MessageBox.Show("Please fill everything! ")
        'End While

        inFileNameToRead = txtFileNameToRead.Text
        inPartsNumber = txtPartsNumber.Text
        Me.Close()



    End Sub
    Private Sub Button2Clear_Click(sender As Object, e As EventArgs) Handles Button2Clear.Click
        txtFileNameToRead.Text = String.Empty
        txtPartsNumber.Text = String.Empty

    End Sub


    Public ReadOnly Property GSFileNameToRead As String
        Get
            Return inFileNameToRead
        End Get

    End Property


    Public ReadOnly Property GSPartsNumber As Integer
        Get
            Return inPartsNumber
        End Get

    End Property

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class