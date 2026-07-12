using System.ComponentModel;
using ModelContextProtocol.Server;
using WordComMcp.Infrastructure;

namespace WordComMcp.Tools;

/// <summary>
/// Trivial diagnostics that validate MCP tool discovery and the JSON result envelope
/// (issue 0.3). No COM access, so they work before any Word integration exists.
/// </summary>
[McpServerToolType]
public static class DiagnosticTools
{
    [McpServerTool(Name = "ping")]
    [Description("Health check. Returns a success envelope with a 'pong' message.")]
    public static string Ping() => McpResult.Ok(new { message = "pong" });

    [McpServerTool(Name = "echo")]
    [Description("Echoes the supplied message back inside a success envelope.")]
    public static string Echo(
        [Description("The message to echo back.")] string message) =>
        McpResult.Ok(new { message });
}
