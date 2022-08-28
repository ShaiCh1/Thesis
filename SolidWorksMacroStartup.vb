
Imports System
Imports System.Text
Imports System.Runtime.InteropServices

<Guid("fcaaccd5-4cd6-4682-a2d6-5f5c3d34c3d8")>
Partial Public Class SolidWorksMacro
    ' [TODO] Non-user code interferes with exception handling.
    '[System.Diagnostics.DebuggerNonUserCodeAttribute()]

    <System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> _
    Public Function Execute(strDebug As [String]) As [String]
        Dim result As [String] = ""

        If strDebug <> "" Then
            '
            ' In debug mode, do not enclose in try/catch.
            '

            If Me.swApp Is Nothing Then
                Throw New System.NullReferenceException("SolidWorksMacro.swApp == null")
            End If

            Me.Main()

            Me.swApp = Nothing
        Else
            '
            ' In non-debug mode, catch any exceptions, and feed them back to the caller.
            '

            Try
                If Me.swApp Is Nothing Then
                    Throw New System.NullReferenceException("SolidWorksMacro.swApp == null")
                End If

                Me.Main()
            Catch ex As Exception
                Dim topex As Exception = ex

                Dim sb As New StringBuilder()

                While ex IsNot Nothing
                    sb.AppendLine([String].Format("{0}: {1} (0x{2:X8})", ex.[GetType]().ToString(), ex.Message, ex.Message))
                    sb.AppendLine([String].Format("Source = {0}", If(ex.Source, "<empty>")))
                    sb.AppendLine([String].Format("Stack trace = {0}", If(ex.StackTrace, "<empty>")))

                    If ex.InnerException IsNot Nothing Then
                        sb.AppendLine()
                        sb.AppendLine("Inner exception:")
                    End If

                    ' Walk the exception chain.
                    ex = ex.InnerException
                End While

                result = sb.ToString()

                System.Diagnostics.Debug.WriteLine(result)
            Finally
                Me.swApp = Nothing
            End Try
        End If

        Return (result)
    End Function
End Class

