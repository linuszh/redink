Option Strict On
Option Explicit On

Imports System.Linq

Namespace Agents

    Public NotInheritable Class SubAgentToolScopeInitializer

        Private Sub New()
        End Sub

        Public NotInheritable Class InitializationResult
            Public Property AuthoritativeRegistry As ToolRegistry
            Public Property NarrowedRegistry As ToolRegistry
            Public Property RequestedToolNames As List(Of String)
            Public Property ResolvedTools As List(Of SharedLibrary.ModelConfig)
            Public Property MissingToolNames As List(Of String)
            Public Property SelectedTools As List(Of SharedLibrary.ModelConfig)

            Public ReadOnly Property HasRequestedTools As Boolean
                Get
                    Return RequestedToolNames IsNot Nothing AndAlso RequestedToolNames.Count > 0
                End Get
            End Property

            Public ReadOnly Property ResolvedToolNames As List(Of String)
                Get
                    Return GetToolNames(ResolvedTools, includeLoader:=True)
                End Get
            End Property

            Public ReadOnly Property FinalSelectedToolNames As List(Of String)
                Get
                    Return GetToolNames(SelectedTools, includeLoader:=True)
                End Get
            End Property

            Public ReadOnly Property FinalCallableToolNames As List(Of String)
                Get
                    Return GetToolNames(SelectedTools, includeLoader:=False)
                End Get
            End Property

            Public ReadOnly Property HasMissingRequestedTools As Boolean
                Get
                    Return MissingToolNames IsNot Nothing AndAlso MissingToolNames.Count > 0
                End Get
            End Property

            Public ReadOnly Property MissingFinalToolNames As List(Of String)
                Get
                    Dim result As New List(Of String)()

                    If Not HasRequestedTools Then
                        Return result
                    End If

                    Dim finalSet As New HashSet(Of String)(
                        FinalCallableToolNames,
                        StringComparer.OrdinalIgnoreCase)

                    For Each requestedName In RequestedToolNames
                        If Not finalSet.Contains(requestedName) Then
                            result.Add(requestedName)
                        End If
                    Next

                    Return result
                End Get
            End Property

            Public ReadOnly Property HasMissingFinalToolNames As Boolean
                Get
                    Return MissingFinalToolNames.Count > 0
                End Get
            End Property
        End Class

        Public Shared Function Initialize(authoritativeRegistry As ToolRegistry,
                                  requestedToolNames As IEnumerable(Of String)) As InitializationResult

            Dim requested = NormalizeToolNames(requestedToolNames)
            Dim sourceRegistry As ToolRegistry =
        If(authoritativeRegistry Is Nothing,
           New ToolRegistry(),
           authoritativeRegistry.Snapshot())

            Dim resolvedTools As List(Of SharedLibrary.ModelConfig)

            If requested Is Nothing Then
                resolvedTools = NormalizeToolConfigs(sourceRegistry.MaterializeAll())
            Else
                resolvedTools = NormalizeToolConfigs(sourceRegistry.Materialize(requested))
            End If

            Dim resolvedNameSet As New HashSet(Of String)(
        resolvedTools.
            Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
            Select(Function(t) t.ToolName),
        StringComparer.OrdinalIgnoreCase)

            Dim missingToolNames As New List(Of String)()

            If requested IsNot Nothing Then
                For Each requestedName In requested
                    If Not resolvedNameSet.Contains(requestedName) Then
                        missingToolNames.Add(requestedName)
                    End If
                Next
            End If

            Dim narrowedRegistry As ToolRegistry = sourceRegistry.Narrow(requested)
            Dim selectedTools As New List(Of SharedLibrary.ModelConfig)(resolvedTools)

            Dim loaderManifests = narrowedRegistry.ListManifests().
        Where(Function(m)
                  Return m IsNot Nothing AndAlso
                         Not String.IsNullOrWhiteSpace(m.Name) AndAlso
                         Not m.Name.Equals(ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase)
              End Function).
        OrderBy(Function(m) m.Name, StringComparer.OrdinalIgnoreCase).
        ToList()

            If loaderManifests.Count > 0 Then
                Dim loader = ToolLoaderTool.Build(loaderManifests)

                If loader IsNot Nothing AndAlso
           Not selectedTools.Any(Function(t)
                                     Return t IsNot Nothing AndAlso
                                            Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                            t.ToolName.Equals(loader.ToolName, StringComparison.OrdinalIgnoreCase)
                                 End Function) Then
                    selectedTools.Add(loader)
                End If
            End If

            Return New InitializationResult() With {
        .AuthoritativeRegistry = sourceRegistry,
        .NarrowedRegistry = narrowedRegistry,
        .RequestedToolNames = If(requested, New List(Of String)()),
        .ResolvedTools = resolvedTools,
        .MissingToolNames = missingToolNames,
        .SelectedTools = selectedTools
    }
        End Function

        Private Shared Function NormalizeToolNames(names As IEnumerable(Of String)) As List(Of String)
            If names Is Nothing Then Return Nothing

            Dim result As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each rawName In names
                Dim name As String = If(rawName, "").Trim()
                If name = "" Then Continue For

                If seen.Add(name) Then
                    result.Add(name)
                End If
            Next

            Return result
        End Function

        Private Shared Function NormalizeToolConfigs(tools As IEnumerable(Of SharedLibrary.ModelConfig)) As List(Of SharedLibrary.ModelConfig)
            Dim result As New List(Of SharedLibrary.ModelConfig)()
            If tools Is Nothing Then Return result

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each tool In tools
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For

                If seen.Add(tool.ToolName) Then
                    result.Add(tool)
                End If
            Next

            Return result
        End Function

        Private Shared Function GetToolNames(tools As IEnumerable(Of SharedLibrary.ModelConfig),
                                             includeLoader As Boolean) As List(Of String)

            Dim result As New List(Of String)()
            If tools Is Nothing Then Return result

            For Each tool In tools
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For

                If Not includeLoader AndAlso
                   tool.ToolName.Equals(ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                result.Add(tool.ToolName)
            Next

            Return result
        End Function

    End Class

End Namespace