' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SkillInvokeTool.vb
' Purpose: Implements the universal "skill_use" tool. The model calls
'             skill_use(name, input?)
'          and receives the SKILL.md body (loaded lazily) along with an
'          inventory of the skill's scripts/ and references/ directories.
'
' The model is then expected to follow those instructions in subsequent turns,
' optionally reading individual referenced files via text.read (step 8) or
' executing scripts via js.run (step 8).
'
' Security:
'   - allowed-tools (Claude frontmatter) is communicated to the model in the
'     response so it knows which tools the skill expects. Enforcement of the
'     narrowing happens in the host runner (step 8).
'   - Script bodies are NOT auto-loaded; only an inventory (names + sizes) is
'     returned. The model fetches what it needs.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class SkillInvokeTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "skill_use"

        Public Shared Function Build() As SharedLibrary.ModelConfig
            Dim def =
                "{""name"":""" & ToolName & """," &
                """description"":""Load and apply a Skill (Claude-style SKILL.md). Returns the skill's instructions and an inventory of its scripts/ and references/ files. Read individual files with text_read; execute scripts with js_run. Use this when a relevant skill is offered above and the user's task matches."",""parameters"":{" &
                """type"":""object""," &
                """properties"":{" &
                """name"":{""type"":""string"",""description"":""The skill name (matches the Skill listed above).""}," &
                """input"":{""type"":""string"",""description"":""Optional input or sub-task description for the skill.""}}," &
                """required"":[""name""]}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": Load a Skill's instructions (lazy). Call this once per skill, then follow its directions in subsequent turns.",
                .ModelDescription = "Skill loader",
                .Tool = True,
                .ToolPriority = 940,
                .ToolErrorHandling = "skip"
            }
        End Function

        ''' <summary>
        ''' Executes the skill_use call. Returns a JSON string suitable for the tool response.
        ''' Caller passes the dictionary from ToolCall.Arguments.
        ''' </summary>
        Public Shared Function Execute(arguments As IDictionary(Of String, Object)) As String
            Try
                Dim name As String = GetStr(arguments, "name")
                Dim input As String = GetStr(arguments, "input")

                If String.IsNullOrWhiteSpace(name) Then
                    name = GetStr(arguments, "tool")
                End If

                If String.IsNullOrWhiteSpace(name) Then
                    name = GetStr(arguments, "skill")
                End If

                If name.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then
                    name = name.Substring("skill_".Length)
                End If

                If String.IsNullOrWhiteSpace(name) Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "missing_name"})
                End If

                Dim sk = AgentResources.FindSkill(name)
                If sk Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "skill_not_found", Key .name = name})
                End If

                Dim body As String = sk.LoadBody()
                Dim scripts As List(Of Object) = InventoryDir(sk.ScriptsDir)
                Dim references As List(Of Object) = InventoryDir(sk.ReferencesDir)

                Dim result As New JObject()
                result("name") = sk.Name
                result("description") = If(sk.Description, "")
                result("origin") = If(sk.IsLocal, "local", "central")
                result("dir") = sk.DirectoryPath
                result("network_allowed") = sk.Network
                result("allowed_tools") = If(sk.AllowedTools Is Nothing,
                                             New JArray(),
                                             JArray.FromObject(sk.AllowedTools))
                result("instructions") = body
                result("scripts") = JArray.FromObject(scripts)
                result("references") = JArray.FromObject(references)
                If Not String.IsNullOrWhiteSpace(input) Then result("input") = input

                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {Key .error = "skill_invoke_failed", Key .message = ex.Message})
            End Try
        End Function

        Private Shared Function InventoryDir(dir As String) As List(Of Object)
            Dim list As New List(Of Object)
            If String.IsNullOrWhiteSpace(dir) OrElse Not Directory.Exists(dir) Then Return list
            Try
                For Each f In Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories)
                    Try
                        Dim fi As New FileInfo(f)
                        Dim rel = f.Substring(dir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                        list.Add(New With {
                            Key .path = rel,
                            Key .size = fi.Length
                        })
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return list
        End Function

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

    End Class

End Namespace