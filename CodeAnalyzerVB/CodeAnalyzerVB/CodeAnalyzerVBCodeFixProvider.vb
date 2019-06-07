Imports System
Imports System.Collections.Generic
Imports System.Collections.Immutable
Imports System.Composition
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.Rename
Imports Microsoft.CodeAnalysis.Text

<ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(CodeAnalyzerVBCodeFixProvider)), [Shared]>
Public Class CodeAnalyzerVBCodeFixProvider
    Inherits CodeFixProvider

    Private Const title As String = "Ignorar Mensagem"

    Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
        Get
            Return ImmutableArray.Create(CodeAnalyzerVBAnalyzer.DiagnosticId)
        End Get
    End Property

    Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
        Return WellKnownFixAllProviders.BatchFixer
    End Function

    Public NotOverridable Overrides Async Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
        Dim root = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)

        Dim diagnostic = context.Diagnostics.First()
        Dim diagnosticSpan = diagnostic.Location.SourceSpan

        Dim declaration = root.FindToken(diagnosticSpan.Start).Parent.AncestorsAndSelf.OfType(Of MethodBlockSyntax)().First()

        context.RegisterCodeFix(
            CodeAction.Create(
                title:=title,
                createChangedSolution:=Function(c) AddSuppressMessage(context.Document, declaration, c),
                equivalenceKey:=title),
            diagnostic)
    End Function

    Private Async Function AddSuppressMessage(ByVal document As Document, ByVal methodDeclaration As MethodBlockSyntax, ByVal cancellationToken As CancellationToken) As Task(Of Solution)
        Dim root = Await document.GetSyntaxRootAsync(cancellationToken)

        Dim name = SyntaxFactory.ParseName("System.Diagnostics.CodeAnalysis.SuppressMessage")
        Dim arguments = SyntaxFactory.ParseArgumentList("(""Desenbahia"",""QA0001"")")

        Dim justification = SyntaxFactory.SimpleArgument(SyntaxFactory.NameColonEquals(SyntaxFactory.IdentifierName("Justification")),
                                                         SyntaxFactory.StringLiteralExpression(
                                                         SyntaxFactory.StringLiteralToken(
                                                         """O método não possui problemas de design.""",
                                                         """O método não possui problemas de design.""")))
        arguments = arguments.AddArguments(justification)
        Dim attribute = SyntaxFactory.Attribute(Nothing, name, arguments)
        Dim attributeList = New SeparatedSyntaxList(Of AttributeSyntax)()
        attributeList = attributeList.Add(attribute)
        Dim list = SyntaxFactory.AttributeList(attributeList).WithLeadingTrivia(methodDeclaration.SubOrFunctionStatement.GetLeadingTrivia())
        Dim attributes = methodDeclaration.SubOrFunctionStatement.AttributeLists.Add(list)


        Return document.WithSyntaxRoot(root.ReplaceNode(methodDeclaration.SubOrFunctionStatement, methodDeclaration.SubOrFunctionStatement.WithoutLeadingTrivia().WithAttributeLists(attributes))).Project.Solution

    End Function
End Class
