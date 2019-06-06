Imports System
Imports System.Collections.Generic
Imports System.Collections.Immutable
Imports System.Linq
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.Diagnostics
Imports CodeMetricsVB
Imports CodeMetricsVB.MetricsAnalyzers

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
Public Class CodeAnalyzerVBAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "QA0001"

    ' You can change these strings in the Resources.resx file. If you do not want your analyzer to be localize-able, you can use regular strings for Title and MessageFormat.
    ' See https://github.com/dotnet/roslyn/blob/master/docs/analyzers/Localizing%20Analyzers.md for more on localization
    Private Shared ReadOnly Title As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.AnalyzerTitle), My.Resources.ResourceManager, GetType(My.Resources.Resources))
    Private Shared ReadOnly MessageFormat As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.AnalyzerMessageFormat), My.Resources.ResourceManager, GetType(My.Resources.Resources))
    Private Shared ReadOnly Description As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.AnalyzerDescription), My.Resources.ResourceManager, GetType(My.Resources.Resources))
    Private Const Category = "Desenbahia"

    Private Shared Rule As New DiagnosticDescriptor(DiagnosticId, Title, MessageFormat, Category, DiagnosticSeverity.Warning, isEnabledByDefault:=True, description:=Description)

    Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
        Get
            Return ImmutableArray.Create(Rule)
        End Get
    End Property

    Public Overrides Sub Initialize(context As AnalysisContext)
        context.RegisterCodeBlockAction(AddressOf AnalyzeMethod)
    End Sub

    Dim metricsContext As MetricsContext = New MetricsContext
    Dim analyzers As List(Of MetricAnalyzer) = New List(Of MetricAnalyzer)

    Private Sub AnalyzeMethod(context As CodeBlockAnalysisContext)

        If TypeOf context.CodeBlock Is Syntax.MethodBlockSyntax Then

            Dim methodBlockSyntax = CType(context.CodeBlock, MethodBlockSyntax)
            Dim node As MemberNode = New MemberNode(metricsContext)
            node.Calculate(methodBlockSyntax)

            Dim c As MLTool.Classifier.WekaFacade = New MLTool.Classifier.WekaFacade("D:\\dataset.arff")
            Dim j48 As MLTool.Classifier.J48Classifier = c.BuildJ48Classifier()

            Dim metricsAttributes As String = MetricsToString(node)
            Dim result As Double = j48.ClassifyInstance(metricsAttributes)

            If result = "1" Then
                Dim diag = Diagnostic.Create(Rule, methodBlockSyntax.SubOrFunctionStatement.GetLocation, methodBlockSyntax.SubOrFunctionStatement.Identifier)
                context.ReportDiagnostic(diag)
            End If

            'For Each lMetrics In node.Metrics
            '    If lMetrics.Value > 6 Then
            '        Dim diag = Diagnostic.Create(Rule, methodBlockSyntax.SubOrFunctionStatement.GetLocation, methodBlockSyntax.SubOrFunctionStatement.Identifier)
            '        context.ReportDiagnostic(diag)
            '    End If
            'Next
        End If

    End Sub

    Private Function MetricsToString(ByVal node As MemberNode) As String

        Dim LOC As String = "0"
        Dim NP As String = "0"
        Dim NV As String = "0"
        Dim CC As String = "0"
        Dim PA As String = "0"
        Dim AE As String = "0"

        For Each lMetrics In node.Metrics

            If TypeOf lMetrics Is MetricsAnalyzers.Metrics.LinesOfCodeMetric Then
                LOC = lMetrics.Value.ToString()
            End If

            If TypeOf lMetrics Is MetricsAnalyzers.Metrics.NumberOfParametersMetric Then
                NP = lMetrics.Value.ToString()
            End If

            If TypeOf lMetrics Is MetricsAnalyzers.Metrics.NumberOfLocalVariablesMetric Then
                NV = lMetrics.Value.ToString()
            End If

            If TypeOf lMetrics Is MetricsAnalyzers.Metrics.CyclomaticComplexityMetric Then
                CC = lMetrics.Value.ToString()
            End If

            If TypeOf lMetrics Is MetricsAnalyzers.Metrics.NestingDepthMetric Then
                PA = lMetrics.Value.ToString()
            End If

            If TypeOf lMetrics Is MetricsAnalyzers.Metrics.NumberOfMethodsInvokedMetric Then
                AE = lMetrics.Value.ToString()
            End If

        Next

        Dim attrs As String = LOC & "," & NP & "," & NV & "," & CC & "," & PA & "," & AE
        Return attrs
    End Function
End Class
