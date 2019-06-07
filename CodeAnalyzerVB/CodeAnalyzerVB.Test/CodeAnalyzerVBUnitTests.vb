Imports CodeAnalyzerVB
Imports CodeAnalyzerVB.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace CodeAnalyzerVB.Test
    <TestClass>
    Public Class UnitTest
        Inherits CodeFixVerifier

        'No diagnostics expected to show up
        <TestMethod>
        Public Sub TestMethod1()
            Dim test = ""
            VerifyBasicDiagnostic(test)
        End Sub

        'Diagnostic And CodeFix both triggered And checked for
        <TestMethod>
        Public Sub TestMethod2()

            Dim test = "
Public Class Class1


    Private Shared Sub Main(ByVal args As String())
        Dim a As String = ""a""
        Dim b As String = ""a""
        Dim c As String = ""a""
        Dim d As String = ""a""
        Dim e As String = ""a""
        Dim f As String = ""a""
        Dim g As String = ""a""
        Dim h As String = ""a""
        Dim i As String = ""a""
        Dim j As String = ""a""
        Dim k As String = ""a""
        Dim l As String = ""a""
        Dim c1 As String = ""a""
        Dim d2 As String = ""a""
        Dim e3 As String = ""a""
        Dim f4 As String = ""a""
        Dim g5 As String = ""a""
        Dim concant As String = a + b + c + d + e + f + g + h + i + j + k + l

        Console.WriteLine(Teste(43, 69, 26, 21))
    End Sub

    Public Shared Function Teste(ByVal cc As Integer, ByVal loc As Integer, ByVal pa As Integer, ByVal nv As Integer)

        If cc <= 11 Then
            Return ""Não""
        ElseIf cc > 11 Then

            If loc > 45 Then
                Return ""Sim""

            ElseIf loc <= 45 Then

                If pa <= 3 Then
                    Return ""Não""

                ElseIf pa > 3 Then

                    If loc > 25 Then
                        Return ""Sim""

                    ElseIf loc <= 25 Then

                        If nv <= 11 Then
                            Return ""Não""

                        ElseIf nv > 11 Then

                            If loc <= 22 Then
                                Return ""Não""

                            ElseIf loc > 22 Then
                                If loc > 25 Then
                                    Return ""Sim""

                                ElseIf loc <= 25 Then

                                    If nv <= 11 Then
                                        Return ""Não""

                                    ElseIf nv > 11 Then

                                        If loc <= 22 Then
                                            Return ""Não""

                                        ElseIf loc > 22 Then

                                        End If

                                    End If

                                End If
                            End If

                        End If

                    End If

                End If

            End If
        End If
        Return ""Não""
    End Function

End Class"
            Dim expected = New DiagnosticResult With {.Id = "QA0001",
                .Message = String.Format("O método '{0}' possui problemas de design que podem afetar negativamente a manutenibilidade do método.", "Teste"),
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {
                        New DiagnosticResultLocation("Class1.vb", 27, 5)
                    }
            }


            VerifyBasicDiagnostic(test, expected)

            Dim fixtest = "
Public Class Class1


    Private Shared Sub Main(ByVal args As String())
        Dim a As String = ""a""
        Dim b As String = ""a""
        Dim c As String = ""a""
        Dim d As String = ""a""
        Dim e As String = ""a""
        Dim f As String = ""a""
        Dim g As String = ""a""
        Dim h As String = ""a""
        Dim i As String = ""a""
        Dim j As String = ""a""
        Dim k As String = ""a""
        Dim l As String = ""a""
        Dim c1 As String = ""a""
        Dim d2 As String = ""a""
        Dim e3 As String = ""a""
        Dim f4 As String = ""a""
        Dim g5 As String = ""a""
        Dim concant As String = a + b + c + d + e + f + g + h + i + j + k + l

        Console.WriteLine(Teste(43, 69, 26, 21))
    End Sub
    <System.Diagnostics.CodeAnalysis.SuppressMessage(""Desenbahia"", ""QA0001"", Justification:=""O método não possui problemas de design."")>
    Public Shared Function Teste(ByVal cc As Integer, ByVal loc As Integer, ByVal pa As Integer, ByVal nv As Integer)

        If cc <= 11 Then
            Return ""Não""
        ElseIf cc > 11 Then

            If loc > 45 Then
                Return ""Sim""

            ElseIf loc <= 45 Then

                If pa <= 3 Then
                    Return ""Não""

                ElseIf pa > 3 Then

                    If loc > 25 Then
                        Return ""Sim""

                    ElseIf loc <= 25 Then

                        If nv <= 11 Then
                            Return ""Não""

                        ElseIf nv > 11 Then

                            If loc <= 22 Then
                                Return ""Não""

                            ElseIf loc > 22 Then
                                If loc > 25 Then
                                    Return ""Sim""

                                ElseIf loc <= 25 Then

                                    If nv <= 11 Then
                                        Return ""Não""

                                    ElseIf nv > 11 Then

                                        If loc <= 22 Then
                                            Return ""Não""

                                        ElseIf loc > 22 Then

                                        End If

                                    End If

                                End If
                            End If

                        End If

                    End If

                End If

            End If
        End If
        Return ""Não""
    End Function

End Class"
            VerifyBasicFix(test, fixtest)
        End Sub

        Protected Overrides Function GetBasicCodeFixProvider() As CodeFixProvider
            Return New CodeAnalyzerVBCodeFixProvider()
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New CodeAnalyzerVBAnalyzer()
        End Function

    End Class
End Namespace
