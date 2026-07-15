using System.Reflection;
using WordComMcp.Tools;
using Xunit;

namespace WordComMcp.Tests;

/// <summary>
/// Word-free checks on the MCP tool surface: every Tier-1 tool is discoverable under its
/// <c>word_live_*</c> name, and — per the Conventions "tracked changes are human-owned" rule —
/// the server exposes no accept/reject revision tool.
/// </summary>
public class ToolSurfaceTests
{
    private static IReadOnlyList<string> ToolNames()
    {
        var assembly = typeof(DiagnosticTools).Assembly;
        var names = new List<string>();
        foreach (var type in assembly.GetTypes())
        {
            if (type.GetCustomAttributes().All(a => a.GetType().Name != "McpServerToolTypeAttribute"))
            {
                continue;
            }

            foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
            {
                var toolAttr = method.GetCustomAttributes()
                    .FirstOrDefault(a => a.GetType().Name == "McpServerToolAttribute");
                if (toolAttr is null)
                {
                    continue;
                }

                var name = toolAttr.GetType().GetProperty("Name")?.GetValue(toolAttr) as string;
                if (!string.IsNullOrEmpty(name))
                {
                    names.Add(name);
                }
            }
        }

        return names;
    }

    [Theory]
    [InlineData("word_live_toggle_track_changes")]
    [InlineData("word_live_list_styles")]
    [InlineData("word_live_get_markdown")]
    [InlineData("word_live_find_text")]
    [InlineData("word_live_insert_markdown")]
    [InlineData("word_live_replace_text")]
    [InlineData("word_live_delete_text")]
    [InlineData("word_live_save")]
    [InlineData("word_live_save_as")]
    [InlineData("word_live_get_info")]
    [InlineData("word_live_set_block_style")]
    [InlineData("word_live_apply_markdown")]
    [InlineData("word_live_add_comment")]
    [InlineData("word_live_get_comments")]
    [InlineData("word_live_reply_to_comment")]
    [InlineData("word_live_resolve_comment")]
    [InlineData("word_live_delete_comment")]
    [InlineData("word_live_list_revisions")]
    [InlineData("word_live_list_open")]
    [InlineData("word_live_list_headings")]
    [InlineData("word_live_goto")]
    [InlineData("word_live_set_core_properties")]
    public void AllTierOneTools_AreDiscoverable(string expected)
    {
        // Act / Assert
        Assert.Contains(expected, ToolNames());
    }

    [Fact]
    public void NoTool_ExposesAcceptOrRejectRevisions()
    {
        // Assert — tracked changes are human-owned; the MCP never accepts/rejects revisions.
        Assert.DoesNotContain(ToolNames(), n =>
            n.Contains("accept", StringComparison.OrdinalIgnoreCase) ||
            n.Contains("reject", StringComparison.OrdinalIgnoreCase));
    }
}
