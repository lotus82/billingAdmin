
Imports System.Exception
Public Class ClassVvod

    Public Shared Function Format_Validation(Str As String, Format As Integer) As Boolean
        Dim Valid As Boolean = False
        If Format = 1 Then 'проверка на соответствие формату GUID
            Dim g As Guid
            If Str = "" Then
                Valid = True
            Else
                Try
                    g = Guid.Parse(Str)
                    Valid = True
                Catch ex As Exception
                    Valid = False

                End Try
            End If

        End If
        If Format = 2 Then 'проверка на соответствие формату Целочисленный
            Dim c As Int64
            If Str = "" Then
                Valid = True
            Else
                Try
                    c = Int64.Parse(Str)
                    Valid = True
                Catch ex As Exception
                    Valid = False
                End Try
            End If
        End If
        If Format = 3 Then 'проверка на соответствие формату Числовой
            Dim d As Decimal
            If Str = "" Then
                Valid = True
            Else
                Try
                    d = Decimal.Parse(Str)
                    Valid = True
                Catch ex As Exception
                    Valid = False
                End Try
            End If
        End If
        Return Valid
    End Function



End Class
