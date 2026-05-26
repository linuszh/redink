' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.Parameterhandling.vb
' Purpose: Tool parameter schema handling, placeholder resolution, and signature building.
'          Prepares tool call arguments for external API invocation.
'
' Architecture:
'  - Parameter Schema Extraction:
'      - GetToolParameterSchemas(): Parses ToolDefinition JSON for parameter properties.
'      - GetToolRequiredParameters(): Extracts required parameter names from schema.
'      - GetToolParameterType(): Resolves type string (string/integer/number/boolean/array/object).
'      - GetToolParameterEnumValues(): Extracts enum constraint values.
'  - Placeholder Resolution:
'      - FormatToolValueForPlaceholder(): Type-aware formatting (bool/int/number/array/object/string).
'      - ResolveToolDefaultValue(): Applies ToolParameterDefaults when placeholder is unreplaced.
'      - RemoveToolArgumentPlaceholderProperty(): Strips optional unreplaced placeholders from JSON.
'  - Type-Specific Formatting:
'      - Boolean: Parses "true"/"false"/"1"/"0"/"yes"/"no" via TryParseBooleanLiteral().
'      - Integer: InvariantCulture parsing with Long range support.
'      - Number: InvariantCulture parsing with comma-to-dot normalization.
'      - Array: Validates and preserves JSON arrays; defaults to "[]".
'      - Object: Validates and preserves JSON objects; defaults to "{}".
'      - String: Escapes via EscapeJsonString() for JSON string literal injection.
'  - Signature Building:
'      - BuildToolCallSignature(): Creates deterministic signature for duplicate detection.
'      - BuildToolArgumentsSignature(): Normalized JSON representation of arguments dict.
'      - BuildExecutedToolSignature(): Full signature including response/error state.
'      - NormalizeToolArgumentValue(): Recursive normalization for JToken/JValue/collections.
'  - Placeholder Property Removal:
'      - Strips "{param}:"{param}"" and "{param}":{param} patterns via regex.
'      - Cleans up trailing/leading commas and empty structural remnants.
'
' Key Functions:
'  - GetToolParameterSchemas(): Schema extraction entry point.
'  - FormatToolValueForPlaceholder(): Type-aware value formatting.
'  - ResolveToolDefaultValue(): Default value resolution with schema awareness.
'  - BuildToolCallSignature(): Signature for duplicate detection.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn



    Private Function GetToolParameterSchemas(toolConfig As ModelConfig) As Dictionary(Of String, JToken)
        Dim result As New Dictionary(Of String, JToken)(StringComparer.OrdinalIgnoreCase)

        If toolConfig Is Nothing OrElse String.IsNullOrWhiteSpace(toolConfig.ToolDefinition) Then
            Return result
        End If

        Try
            Dim toolDefinition As JObject = JObject.Parse(toolConfig.ToolDefinition)
            Dim propertiesObject As JObject = TryCast(toolDefinition.SelectToken("parameters.properties"), JObject)

            If propertiesObject Is Nothing Then
                Return result
            End If

            For Each prop As JProperty In propertiesObject.Properties()
                result(prop.Name) = prop.Value
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn(
                "Failed to parse tool parameter schemas.",
                details:=$"ToolName='{If(toolConfig.ToolName, "")}'",
                ex:=ex)
        End Try

        Return result
    End Function

    Private Function GetToolRequiredParameters(toolConfig As ModelConfig) As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If toolConfig Is Nothing OrElse String.IsNullOrWhiteSpace(toolConfig.ToolDefinition) Then
            Return result
        End If

        Try
            Dim toolDefinition As JObject = JObject.Parse(toolConfig.ToolDefinition)
            Dim requiredArray As JArray = TryCast(toolDefinition.SelectToken("parameters.required"), JArray)

            If requiredArray Is Nothing Then
                Return result
            End If

            For Each item As JToken In requiredArray
                Dim name As String = If(item, "").ToString().Trim()
                If name <> "" Then
                    result.Add(name)
                End If
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn(
                "Failed to parse required tool parameters.",
                details:=$"ToolName='{If(toolConfig.ToolName, "")}'",
                ex:=ex)
        End Try

        Return result
    End Function

    Private Function GetToolParameterType(schemaToken As JToken) As String
        If schemaToken Is Nothing Then
            Return "string"
        End If

        Dim typeToken As JToken = schemaToken("type")
        Dim typeName As String = If(typeToken, "").ToString().Trim().ToLowerInvariant()

        If typeName <> "" Then
            Return typeName
        End If

        If schemaToken("enum") IsNot Nothing Then
            Return "string"
        End If

        Return "string"
    End Function

    Private Function GetToolParameterEnumValues(schemaToken As JToken) As List(Of String)
        Dim values As New List(Of String)()

        If schemaToken Is Nothing Then
            Return values
        End If

        Dim enumArray As JArray = TryCast(schemaToken("enum"), JArray)
        If enumArray Is Nothing Then
            Return values
        End If

        For Each item As JToken In enumArray
            Dim value As String = If(item, "").ToString()
            If value <> "" Then
                values.Add(value)
            End If
        Next

        Return values
    End Function

    Private Function TryParseBooleanLiteral(value As String, ByRef result As Boolean) As Boolean
        Dim normalized As String = If(value, "").Trim()

        If Boolean.TryParse(normalized, result) Then
            Return True
        End If

        Select Case normalized.ToLowerInvariant()
            Case "1", "yes", "y"
                result = True
                Return True
            Case "0", "no", "n"
                result = False
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function FormatToolValueForPlaceholder(rawValue As String, schemaToken As JToken) As String
        Dim parameterType As String = GetToolParameterType(schemaToken)
        Dim safeValue As String = If(rawValue, "").Trim()

        Select Case parameterType
            Case "boolean"
                Dim boolValue As Boolean = False
                If TryParseBooleanLiteral(safeValue, boolValue) Then
                    Return If(boolValue, "true", "false")
                End If
                Return "false"

            Case "integer"
                Dim longValue As Long
                If Long.TryParse(safeValue, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, longValue) OrElse
                   Long.TryParse(safeValue, longValue) Then
                    Return longValue.ToString(Globalization.CultureInfo.InvariantCulture)
                End If
                Return "0"

            Case "number"
                Dim doubleValue As Double
                Dim normalized As String = safeValue.Replace(","c, "."c)

                If Double.TryParse(normalized, Globalization.NumberStyles.Float Or Globalization.NumberStyles.AllowThousands,
                                   Globalization.CultureInfo.InvariantCulture, doubleValue) OrElse
                   Double.TryParse(safeValue, doubleValue) Then
                    Return doubleValue.ToString(Globalization.CultureInfo.InvariantCulture)
                End If

                Return "0"

            Case "array"
                If safeValue <> "" Then
                    Try
                        Dim parsed As JToken = JToken.Parse(safeValue)
                        If parsed.Type = JTokenType.Array Then
                            Return parsed.ToString(Formatting.None)
                        End If
                    Catch
                    End Try
                End If
                Return "[]"

            Case "object"
                If safeValue <> "" Then
                    Try
                        Dim parsed As JToken = JToken.Parse(safeValue)
                        If parsed.Type = JTokenType.Object Then
                            Return parsed.ToString(Formatting.None)
                        End If
                    Catch
                    End Try
                End If
                Return "{}"

            Case Else
                Return EscapeJsonString(safeValue)
        End Select
    End Function

    Private Function ResolveToolDefaultValue(placeholderName As String,
                                             toolDefaults As IDictionary(Of String, String),
                                             schemaToken As JToken,
                                             isRequired As Boolean,
                                             ByRef shouldRemoveProperty As Boolean) As String
        shouldRemoveProperty = False

        Dim rawDefault As String = ""

        If toolDefaults IsNot Nothing AndAlso toolDefaults.ContainsKey(placeholderName) Then
            rawDefault = If(toolDefaults(placeholderName), "")
        End If

        If String.IsNullOrWhiteSpace(rawDefault) Then
            If isRequired Then
                Return "{" & placeholderName & "}"
            End If

            shouldRemoveProperty = True
            Return ""
        End If

        Return FormatToolValueForPlaceholder(rawDefault, schemaToken)
    End Function

    Private Function RemoveToolArgumentPlaceholderProperty(apiCall As String, placeholderName As String) As String
        If String.IsNullOrWhiteSpace(apiCall) OrElse String.IsNullOrWhiteSpace(placeholderName) Then
            Return apiCall
        End If

        Dim quotedPropertyName As String = Regex.Escape("""" & placeholderName & """")
        Dim rawPlaceholder As String = Regex.Escape("{" & placeholderName & "}")
        Dim quotedPlaceholder As String = Regex.Escape("""{" & placeholderName & "}""")

        Dim patterns As String() = {
            ",\s*" & quotedPropertyName & "\s*:\s*" & quotedPlaceholder,
            ",\s*" & quotedPropertyName & "\s*:\s*" & rawPlaceholder,
            quotedPropertyName & "\s*:\s*" & quotedPlaceholder & "\s*,",
            quotedPropertyName & "\s*:\s*" & rawPlaceholder & "\s*,",
            quotedPropertyName & "\s*:\s*" & quotedPlaceholder,
            quotedPropertyName & "\s*:\s*" & rawPlaceholder
        }

        Dim result As String = apiCall

        For Each pattern As String In patterns
            result = Regex.Replace(
                result,
                pattern,
                "",
                RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        Next

        result = Regex.Replace(result, ",\s*(\}|\])", "$1", RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        result = Regex.Replace(result, "(\{|\[)\s*,", "$1", RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        Return result
    End Function


    Private Shared Function BuildToolNameSet(names As IEnumerable(Of String)) As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If names Is Nothing Then
            Return result
        End If

        For Each name In names
            If String.IsNullOrWhiteSpace(name) Then Continue For
            result.Add(name.Trim())
        Next

        Return result
    End Function

    Private Function BuildExecutedToolSignature(toolCall As ToolCall, toolResponse As ToolResponse) As String
        Dim toolName As String = If(toolCall?.ToolName, "")
        Dim argsSig As String = BuildToolArgumentsSignature(toolCall?.Arguments)
        Dim successSig As String = If(toolResponse IsNot Nothing AndAlso toolResponse.Success, "ok", "err")
        Dim responseSig As String = If(toolResponse?.Response, "")
        Dim errorSig As String = If(toolResponse?.ErrorMessage, "")

        If responseSig.Length > 500 Then
            responseSig = responseSig.Substring(0, 500)
        End If

        If errorSig.Length > 300 Then
            errorSig = errorSig.Substring(0, 300)
        End If

        Return toolName & "|" & argsSig & "|" & successSig & "|" & responseSig & "|" & errorSig
    End Function


    Private Function BuildToolCallSignature(toolCall As ToolCall) As String
        If toolCall Is Nothing Then Return ""

        Dim parts As New List(Of String)()

        If toolCall.Arguments IsNot Nothing Then
            For Each kvp In toolCall.Arguments.OrderBy(Function(x) x.Key, StringComparer.OrdinalIgnoreCase)
                parts.Add(kvp.Key.ToLowerInvariant() & "=" & NormalizeToolArgumentValue(kvp.Value))
            Next
        End If

        Return toolCall.ToolName & "|" & String.Join(";", parts)
    End Function

    Private Function NormalizeToolArgumentValue(value As Object) As String
        If value Is Nothing Then Return "null"

        If TypeOf value Is JValue Then
            Return DirectCast(value, JValue).ToString(Formatting.None)
        End If

        If TypeOf value Is JToken Then
            Return DirectCast(value, JToken).ToString(Formatting.None)
        End If

        If TypeOf value Is IEnumerable(Of Object) AndAlso Not TypeOf value Is String Then
            Dim items = DirectCast(value, IEnumerable(Of Object)).
                Select(Function(v) NormalizeToolArgumentValue(v))
            Return "[" & String.Join(",", items) & "]"
        End If

        Return System.Convert.ToString(value, Globalization.CultureInfo.InvariantCulture)
    End Function


    Private Function BuildToolArgumentsSignature(arguments As Dictionary(Of String, Object)) As String
        If arguments Is Nothing OrElse arguments.Count = 0 Then
            Return "{}"
        End If

        Try
            Dim normalized As New JObject()

            For Each key In arguments.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
                Dim value = arguments(key)

                If TypeOf value Is JToken Then
                    normalized(key) = DirectCast(value, JToken)
                ElseIf value Is Nothing Then
                    normalized(key) = JValue.CreateNull()
                Else
                    normalized(key) = JToken.FromObject(value)
                End If
            Next

            Return normalized.ToString(Formatting.None)
        Catch
            Try
                Return JsonConvert.SerializeObject(arguments)
            Catch
                Return "{}"
            End Try
        End Try
    End Function




End Class