Module MatrixHelper
    Public Function MakeDisplayable(ByVal sourceMatrix(,) As Double) As String
        ' ----- Prepare a multi-line string that shows the contents
        '       of a matrix, a 2D array.
        Dim rows As Integer
        Dim cols As Integer
        Dim eachRow As Integer
        Dim eachCol As Integer
        Dim result As New System.Text.StringBuilder

        ' ----- Process all rows of the matrix, generating one
        '       output line per row.
        rows = UBound(sourceMatrix, 1) + 1
        cols = UBound(sourceMatrix, 2) + 1
        For eachRow = 0 To rows - 1
            ' ----- Process each column of the matrix on a single
            '       row, separating values by commas.
            If (eachRow > 0) Then result.AppendLine()
            For eachCol = 0 To cols - 1
                ' ----- Add a single matrix element to the output.
                If (eachCol > 0) Then result.Append(",")
                result.Append(sourceMatrix(eachRow, eachCol).ToString)
            Next eachCol
        Next eachRow

        ' ----- Finished.
        Return result.ToString
    End Function

    Public Function MakeDisplayable(ByVal sourceArray() As Double) As String
        ' ----- Present an array as multiple lines of output.
        Dim result As New System.Text.StringBuilder
        Dim scanValue As Double

        For Each scanValue In sourceArray
            result.AppendLine(scanValue.ToString)
        Next scanValue

        Return result.ToString
    End Function

    Public Function Inverse(ByVal sourceMatrix(,) As Double) As Double(,)
        ' ----- Build a new matrix that is the mathematical inverse
        '       of the supplied matrix. Multiplying a matrix and its
        '       inverse together will give the identity matrix.
        Dim eachCol As Integer
        Dim eachRow As Integer
        Dim rowsAndCols As Integer

        ' ----- Determine the size of each dimension of the matrix.
        '       Only square matrices can be inverted.
        If (UBound(sourceMatrix, 1) <> UBound(sourceMatrix, 2)) Then
            Throw New Exception("Matrix must be square.")
        End If
        Dim rank As Integer = UBound(sourceMatrix, 1)

        ' ----- Clone a copy of the matrix (not just a new reference).
        Dim workMatrix(,) As Double =
        CType(sourceMatrix.Clone, Double(,))

        ' ----- Variables used for backsolving.
        Dim destMatrix(rank, rank) As Double
        Dim rightHandSide(rank) As Double
        Dim solutions(rank) As Double
        Dim rowPivots(rank) As Integer
        Dim colPivots(rank) As Integer

        ' ----- Use LU decomposition to form a triangular matrix.
        workMatrix = FormLU(workMatrix, rowPivots, colPivots, rowsAndCols)

        ' ----- Backsolve the triangular matrix to get the inverted
        '       value for each position in the final matrix.
        For eachCol = 0 To rank
            rightHandSide(eachCol) = 1
            BackSolve(workMatrix, rightHandSide, solutions, rowPivots, colPivots)
            For eachRow = 0 To rank
                destMatrix(eachRow, eachCol) = solutions(eachRow)
                rightHandSide(eachRow) = 0
            Next eachRow
        Next eachCol

        ' ----- Return the inverted matrix result.
        Return destMatrix
    End Function

    Public Function Determinant(ByVal sourceMatrix(,) As Double) As Double
        ' ----- Calculate the determinant of a matrix.
        Dim result As Double
        Dim pivots As Integer
        Dim count As Integer

        ' ----- Only calculate the determinants of square matrices.
        If (UBound(sourceMatrix, 1) <> UBound(sourceMatrix, 2)) Then
            Throw New Exception(
            "Determinant only calculated for square matrices.")
        End If
        Dim rank As Integer = UBound(sourceMatrix, 1)

        ' ----- Make a copy of the matrix so we can work inside of it.
        Dim workMatrix(rank, rank) As Double
        Array.Copy(sourceMatrix, workMatrix, sourceMatrix.Length)

        ' ----- Use LU decomposition to form a triangular matrix.
        Dim rowPivots(rank) As Integer
        Dim colPivots(rank) As Integer
        workMatrix = FormLU(workMatrix, rowPivots, colPivots, count)

        ' ----- Get the product at each of the pivot points.
        result = 1
        For pivots = 0 To rank
            result *= workMatrix(rowPivots(pivots), colPivots(pivots))
        Next pivots

        ' ----- Determine the sign of the result using LaPlace's formula.
        result = (-1) ^ count * result
        Return result
    End Function

    Public Function SimultEq(ByVal sourceEquations(,) As Double,
        ByVal sourceRHS() As Double) As Double()
        ' ----- Use matrices to solve simultaneous equations.
        Dim rowsAndCols As Integer

        ' ----- The matrix must be square and the array size must match.
        Dim rank As Integer = UBound(sourceEquations, 1)
        If (UBound(sourceEquations, 2) <> rank) Or
            (UBound(sourceRHS, 1) <> rank) Then
            Throw New Exception("Size problem for simultaneous equations.")
        End If

        ' ----- Create some arrays for doing all of the work.
        Dim coefficientMatrix(rank, rank) As Double
        Dim rightHandSide(rank) As Double
        Dim solutions(rank) As Double
        Dim rowPivots(rank) As Integer
        Dim colPivots(rank) As Integer

        ' ----- Make copies of the original matrices so we don't
        '       mess them up.
        Array.Copy(sourceEquations, coefficientMatrix, sourceEquations.Length)
        Array.Copy(sourceRHS, rightHandSide, sourceRHS.Length)

        ' ----- Use LU decomposition to form a triangular matrix.
        coefficientMatrix = FormLU(coefficientMatrix, rowPivots,
        colPivots, rowsAndCols)

        ' ----- Find the unique solution for the upper-triangle.
        BackSolve(coefficientMatrix, rightHandSide, solutions,
        rowPivots, colPivots)

        ' ----- Return the simultaneous equations result in an array.
        Return solutions
    End Function

    Private Function FormLU(ByVal sourceMatrix(,) As Double,
        ByRef rowPivots() As Integer, ByRef colPivots() As Integer,
        ByRef rowsAndCols As Integer) As Double(,)
        ' ----- Perform an LU (lower and upper) decomposition of a matrix,
        '       a modified form of Gaussian elimination.
        Dim eachRow As Integer
        Dim eachCol As Integer
        Dim pivot As Integer
        Dim rowIndex As Integer
        Dim colIndex As Integer
        Dim bestRow As Integer
        Dim bestCol As Integer
        Dim rowToPivot As Integer
        Dim colToPivot As Integer
        Dim maxValue As Double
        Dim testValue As Double
        Dim oldMax As Double
        Const Deps As Double = 0.0000000000000001

        ' ----- Determine the size of the array.
        Dim rank As Integer = UBound(sourceMatrix, 1)
        Dim destMatrix(rank, rank) As Double
        Dim rowNorm(rank) As Double
        ReDim rowPivots(rank)
        ReDim colPivots(rank)

        ' ----- Make a copy of the array so we don't mess it up.
        Array.Copy(sourceMatrix, destMatrix, sourceMatrix.Length)

        ' ----- Initialize row and column pivot arrays.
        For eachRow = 0 To rank
            rowPivots(eachRow) = eachRow
            colPivots(eachRow) = eachRow
            For eachCol = 0 To rank
                rowNorm(eachRow) += Math.Abs(destMatrix(eachRow, eachCol))
            Next eachCol
            If (rowNorm(eachRow) = 0) Then
                Throw New Exception("Cannot invert a singular matrix.")
            End If
        Next eachRow

        ' ----- Use Gauss-Jordan elimination on the matrix rows.
        For pivot = 0 To rank - 1
            maxValue = 0
            For eachRow = pivot To rank
                rowIndex = rowPivots(eachRow)
                For eachCol = pivot To rank
                    colIndex = colPivots(eachCol)
                    testValue = Math.Abs(destMatrix(rowIndex, colIndex)) _
                    / rowNorm(rowIndex)
                    If (testValue > maxValue) Then
                        maxValue = testValue
                        bestRow = eachRow
                        bestCol = eachCol
                    End If
                Next eachCol
            Next eachRow

            ' ----- Detect a singular, or very nearly singular, matrix.
            If (maxValue = 0) Then
                Throw New Exception("Singular matrix used for LU.")
            ElseIf (pivot > 1) Then
                If (maxValue < (Deps * oldMax)) Then
                    Throw New Exception("Non-invertible matrix used for LU.")
                End If
            End If
            oldMax = maxValue

            ' ----- Swap row pivot values for the best row.
            If (rowPivots(pivot) <> rowPivots(bestRow)) Then
                rowsAndCols += 1
                Swap(rowPivots(pivot), rowPivots(bestRow))
            End If

            ' ----- Swap column pivot values for the best column.
            If (colPivots(pivot) <> colPivots(bestCol)) Then
                rowsAndCols += 1
                Swap(colPivots(pivot), colPivots(bestCol))
            End If

            ' ----- Work with the current pivot points.
            rowToPivot = rowPivots(pivot)
            colToPivot = colPivots(pivot)

            ' ----- Modify the remaining rows from the pivot points.
            For eachRow = (pivot + 1) To rank
                rowIndex = rowPivots(eachRow)
                destMatrix(rowIndex, colToPivot) =
                -destMatrix(rowIndex, colToPivot) /
                destMatrix(rowToPivot, colToPivot)
                For eachCol = (pivot + 1) To rank
                    colIndex = colPivots(eachCol)
                    destMatrix(rowIndex, colIndex) +=
                    destMatrix(rowIndex, colToPivot) *
                    destMatrix(rowToPivot, colIndex)
                Next eachCol
            Next eachRow
        Next pivot

        ' ----- Detect a non-invertible matrix.
        If (destMatrix(rowPivots(rank), colPivots(rank)) = 0) Then
            Throw New Exception("Non-invertible matrix used for LU.")
        ElseIf (Math.Abs(destMatrix(rowPivots(rank), colPivots(rank))) /
            rowNorm(rowPivots(rank))) < (Deps * oldMax) Then
            Throw New Exception("Non-invertible matrix used for LU.")
        End If

        ' ----- Success. Return the LU triangular matrix.
        Return destMatrix
    End Function

    Private Sub Swap(ByRef firstValue As Integer, ByRef secondValue As Integer)
        ' ----- Reverse the values of two reference integers.
        Dim holdValue As Integer
        holdValue = firstValue
        firstValue = secondValue
        secondValue = holdValue
    End Sub

    Private Sub BackSolve(ByVal sourceMatrix(,) As Double,
        ByVal rightHandSide() As Double, ByVal solutions() As Double,
        ByRef rowPivots() As Integer, ByRef colPivots() As Integer)
        ' ----- Solve an upper-right-triangle matrix.
        Dim pivot As Integer
        Dim rowToPivot As Integer
        Dim colToPivot As Integer
        Dim eachRow As Integer
        Dim eachCol As Integer
        Dim rank As Integer = UBound(sourceMatrix, 1)

        ' ----- Work through all pivot points. This section builds
        '       the "B" in the AX=B formula.
        For pivot = 0 To (rank - 1)
            colToPivot = colPivots(pivot)
            For eachRow = (pivot + 1) To rank
                rowToPivot = rowPivots(eachRow)
                rightHandSide(rowToPivot) +=
                sourceMatrix(rowToPivot, colToPivot) _
                * rightHandSide(rowPivots(pivot))
            Next eachRow
        Next pivot

        ' ----- Now solve for each X using the general formula
        '       x(i) = (b(i) - summation(a(i,j)x(j)))/a(i,i)
        For eachRow = rank To 0 Step -1
            colToPivot = colPivots(eachRow)
            rowToPivot = rowPivots(eachRow)
            solutions(colToPivot) = rightHandSide(rowToPivot)
            For eachCol = (eachRow + 1) To rank
                solutions(colToPivot) -=
                sourceMatrix(rowToPivot, colPivots(eachCol)) _
                * solutions(colPivots(eachCol))
            Next eachCol
            solutions(colToPivot) /= sourceMatrix(rowToPivot, colToPivot)
        Next eachRow
    End Sub
End Module