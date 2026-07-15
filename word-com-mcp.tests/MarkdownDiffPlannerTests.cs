using WordComMcp.Markdown;
using Xunit;

namespace WordComMcp.Tests;

public class MarkdownDiffPlannerTests
{
    private readonly StyleMap m_styles = StyleMap.Default;

    [Fact]
    public void IdenticalMarkdown_HasNoChanges()
    {
        var document = MarkdownParser.Parse("# Titel\n\nEin Absatz.");

        var plan = MarkdownDiffPlanner.Plan(document, document, this.m_styles);

        Assert.True(plan.CanApply);
        Assert.Empty(plan.Changes);
    }

    [Fact]
    public void SingleWordEdit_IsOneReplacement()
    {
        var current = MarkdownParser.Parse("# Titel\n\nEin alter Absatz.");
        var requested = MarkdownParser.Parse("# Titel\n\nEin neuer Absatz.");

        var plan = MarkdownDiffPlanner.Plan(current, requested, this.m_styles);

        var change = Assert.Single(plan.Changes);
        Assert.Equal("replace", change.Kind);
        Assert.Equal(1, change.OldIndex);
        Assert.Equal(1, change.NewIndex);
    }

    [Fact]
    public void InlineFormattingOnly_IsAStyleChange()
    {
        var current = MarkdownParser.Parse("Ein wichtiger Absatz.");
        var requested = MarkdownParser.Parse("Ein **wichtiger** Absatz.");

        var plan = MarkdownDiffPlanner.Plan(current, requested, this.m_styles);

        Assert.Equal("style", Assert.Single(plan.Changes).Kind);
    }

    [Fact]
    public void ParagraphInsertionAndDeletion_AreTargeted()
    {
        var current = MarkdownParser.Parse("Eins\n\nZwei");
        var inserted = MarkdownParser.Parse("Eins\n\nNeu\n\nZwei");
        var deleted = MarkdownParser.Parse("Zwei");

        var insertPlan = MarkdownDiffPlanner.Plan(current, inserted, this.m_styles);
        var deletePlan = MarkdownDiffPlanner.Plan(current, deleted, this.m_styles);

        Assert.Equal("insert", Assert.Single(insertPlan.Changes).Kind);
        Assert.Equal("delete", Assert.Single(deletePlan.Changes).Kind);
    }

    [Fact]
    public void ExistingListRenumber_ReturnsWarningAndNoChanges()
    {
        var current = MarkdownParser.Parse("- Eins\n- Zwei");
        var requested = MarkdownParser.Parse("1. Eins\n2. Zwei");

        var plan = MarkdownDiffPlanner.Plan(current, requested, this.m_styles);

        Assert.False(plan.CanApply);
        Assert.Empty(plan.Changes);
        Assert.Contains(plan.Warnings, warning => warning.Contains("list", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ExistingListReorder_ReturnsWarningAndNoChanges()
    {
        var current = MarkdownParser.Parse("- Eins\n- Zwei");
        var requested = MarkdownParser.Parse("- Zwei\n- Eins");

        var plan = MarkdownDiffPlanner.Plan(current, requested, this.m_styles);

        Assert.False(plan.CanApply);
        Assert.Empty(plan.Changes);
        Assert.Contains(plan.Warnings, warning => warning.Contains("reordered", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AmbiguousMultiBlockReplacement_ReturnsWarningAndNoChanges()
    {
        var current = MarkdownParser.Parse("Alt eins\n\nAlt zwei");
        var requested = MarkdownParser.Parse("Neu eins\n\nNeu zwei");

        var plan = MarkdownDiffPlanner.Plan(current, requested, this.m_styles);

        Assert.False(plan.CanApply);
        Assert.Empty(plan.Changes);
        Assert.Contains(plan.Warnings, warning => warning.Contains("ambiguous", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ChangedTable_ReturnsWarningAndNoChanges()
    {
        var current = MarkdownParser.Parse("| A | B |\n| --- | --- |\n| 1 | 2 |");
        var requested = MarkdownParser.Parse("| A | B |\n| --- | --- |\n| 1 | 3 |");

        var plan = MarkdownDiffPlanner.Plan(current, requested, this.m_styles);

        Assert.False(plan.CanApply);
        Assert.Empty(plan.Changes);
        Assert.Contains(plan.Warnings, warning => warning.Contains("table", StringComparison.OrdinalIgnoreCase));
    }
}
