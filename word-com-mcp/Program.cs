using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using WordComMcp.Markdown;

var builder = Host.CreateApplicationBuilder(args);

// 0.10 — the stdio transport owns stdout for JSON-RPC framing. Route ALL logs to stderr
// so nothing corrupts the protocol stream. (Tool code must never Console.Write either.)
builder.Logging.AddConsole(options =>
{
    options.LogToStandardErrorThreshold = LogLevel.Trace;
});

// 0.8 — configuration from environment (author/initials/style-map override).
builder.Services.AddSingleton(_ => ServerConfig.FromEnvironment());
builder.Services.AddSingleton(sp => StyleMap.Load(sp.GetRequiredService<ServerConfig>().StyleMapPath));

// 0.4 / 0.5 — COM primitives shared by every future tool. The dispatcher owns the single
// STA thread; the connection caches the resolved Word application on it.
builder.Services.AddSingleton<StaDispatcher>();
builder.Services.AddSingleton<WordConnection>();

// 0.3 — MCP server over stdio; discover [McpServerTool] methods in this assembly.
builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync();
