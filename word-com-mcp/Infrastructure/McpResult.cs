using System.Text.Json;
using System.Text.Json.Serialization;

namespace WordComMcp.Infrastructure;

/// <summary>
/// Unified result/error envelope for every tool (issue 0.9 + Conventions Q4).
/// All tools return a JSON string in the shape:
///   success: { "success": true, "warnings": [...]?, ...data }
///   error:   { "success": false, "error": "&lt;code&gt;", "warnings": [...]? }
/// Errors are structured strings, never raw stack traces.
/// </summary>
public static class McpResult
{
    /// <summary>Well-known error codes shared across tools (issue 0.9 / Conventions Q3).</summary>
    public static class Errors
    {
        public const string WordNotRunning = "Word not running";
        public const string DocumentNotFound = "document not found";
        public const string ReadOnlyOrProtected = "read-only/protected";
        public const string DocumentLocked = "document locked";
        public const string UnknownStyle = "unknown style";
        public const string FindLimitExceeded = "find pattern exceeds 255 characters";
    }

    private static readonly JsonSerializerOptions s_json = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>Build a success envelope. <paramref name="data"/> fields are merged onto the envelope.</summary>
    public static string Ok(object? data = null, IEnumerable<string>? warnings = null)
    {
        var envelope = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["success"] = true,
        };
        MergeWarnings(envelope, warnings);
        MergeData(envelope, data);
        return JsonSerializer.Serialize(envelope, s_json);
    }

    /// <summary>Build a structured error envelope with a stable error code/message.</summary>
    public static string Err(string error, IEnumerable<string>? warnings = null)
    {
        var envelope = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["success"] = false,
            ["error"] = error,
        };
        MergeWarnings(envelope, warnings);
        return JsonSerializer.Serialize(envelope, s_json);
    }

    private static void MergeWarnings(Dictionary<string, object?> envelope, IEnumerable<string>? warnings)
    {
        if (warnings is null)
        {
            return;
        }

        var list = warnings.ToList();
        if (list.Count > 0)
        {
            envelope["warnings"] = list;
        }
    }

    private static void MergeData(Dictionary<string, object?> envelope, object? data)
    {
        if (data is null)
        {
            return;
        }

        // Flatten the caller's data object onto the envelope so callers get
        // { "success": true, "document": "x.docx", ... } rather than a nested "data" node.
        var element = JsonSerializer.SerializeToElement(data, s_json);
        if (element.ValueKind != JsonValueKind.Object)
        {
            envelope["data"] = data;
            return;
        }

        foreach (var property in element.EnumerateObject())
        {
            envelope[property.Name] = property.Value;
        }
    }
}
