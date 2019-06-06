Imports CodeMetricsVB
Imports CodeMetricsVB.MetricsAnalyzers

Public Class Util

    Public Shared Function MetricsToString(ByVal metrics As IEnumerable(Of Metrics.Metric)) As String

        Dim LOC As String = "0"
        Dim NP As String = "0"
        Dim NV As String = "0"
        Dim CC As String = "0"
        Dim PA As String = "0"
        Dim AE As String = "0"

        For Each lMetrics In metrics

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
