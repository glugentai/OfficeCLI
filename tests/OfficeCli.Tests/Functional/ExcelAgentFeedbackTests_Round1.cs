// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for Round 1 bugs reported by Agent A's Excel exploration.
///
/// CONFIRMED BUGS (all tests here are expected to FAIL until fixed):
///
///   Bug 4 (HIGH) — Get outputs duplicate keys: both canonical and legacy alias
///            Example: Get("/Sheet1/A1") returns both "bold" AND "font.bold" in Format.
///            CLAUDE.md rule: "Use one canonical key per semantic value in DocumentNode.Format.
///            Do not store duplicate aliases for the same property."
///            Root cause: ExcelHandler.Helpers.cs line 426 explicitly writes both:
///              node.Format["font.bold"] = true; node.Format["bold"] = true;
///            Same pattern repeats for italic (lines 429-430), superscript (lines 446-447),
///            and subscript (lines 451-452).
///            Fix: Only emit the canonical key. For cell-level font properties, the canonical
///            key should be one of: "bold", "italic", "font.strike", "font.underline",
///            "font.color", "font.size", "font.name", "superscript", "subscript".
///            Remove the duplicate "font.bold", "font.italic", "font.superscript",
///            "font.subscript" entries (or vice versa — pick one canonical form and stick to it).
///
///   Bug 5 (HIGH) — Set with font.strikethrough=true silently fails (no error, no effect)
///            Root cause: ExcelStyleManager.IsStyleKey matches any key starting with "font."
///            so "font.strikethrough" passes the style-key filter. Then the prefix is stripped
///            to "strikethrough". But GetOrCreateFont (line 321) only checks for key "strike",
///            not "strikethrough". The unrecognized key is silently ignored.
///            The shorthand mapping (line 109) also only maps "strike", not "strikethrough".
///            Fix: In ExcelStyleManager.GetOrCreateFont, normalize "strikethrough" to "strike"
///            or add it as an alias. Same should apply to "font.strikethrough" in the
///            shorthand mapping.
///
///   Bug 9 (MEDIUM) — Remove("/Sheet1/autofilter") throws "Cell autofilter not found"
///            Root cause: ExcelHandler.Remove.cs has no dispatch block for "autofilter" path
///            segments. The segment falls through to FindCell() which treats "autofilter" as
///            a cell reference and throws "Cell autofilter not found".
///            Fix: Add a dispatch block in Remove for autofilter, similar to the existing
///            comment[N], validation[N], cf[N] blocks. When path is /SheetName/autofilter,
///            find and remove the AutoFilter element from the worksheet.
/// </summary>
public class ExcelAgentFeedbackTests_Round1 : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelAgentFeedbackTests_Round1()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
    }

    // =====================================================================
    // Bug 4: Get outputs duplicate keys (canonical + legacy alias)
    // =====================================================================

    [Fact]
    public void Bug4_Bold_Get_ShouldNotContainDuplicateKeys()
    {
        // Set bold on a cell
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Hello", ["bold"] = "true" });

        var node = _handler.Get("/Sheet1/A1");

        // The Format dictionary should contain exactly ONE key for bold, not both
        var boldKeys = node.Format.Keys.Where(k =>
            k.Equals("bold", StringComparison.OrdinalIgnoreCase) ||
            k.Equals("font.bold", StringComparison.OrdinalIgnoreCase)).ToList();

        boldKeys.Should().HaveCount(1,
            "CLAUDE.md mandates one canonical key per semantic value — " +
            "both 'bold' and 'font.bold' should not appear simultaneously");
    }

    [Fact]
    public void Bug4_Italic_Get_ShouldNotContainDuplicateKeys()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Hello", ["italic"] = "true" });

        var node = _handler.Get("/Sheet1/A1");

        var italicKeys = node.Format.Keys.Where(k =>
            k.Equals("italic", StringComparison.OrdinalIgnoreCase) ||
            k.Equals("font.italic", StringComparison.OrdinalIgnoreCase)).ToList();

        italicKeys.Should().HaveCount(1,
            "CLAUDE.md mandates one canonical key per semantic value — " +
            "both 'italic' and 'font.italic' should not appear simultaneously");
    }

    [Fact]
    public void Bug4_Superscript_Get_ShouldNotContainDuplicateKeys()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Hello", ["superscript"] = "true" });

        var node = _handler.Get("/Sheet1/A1");

        var superKeys = node.Format.Keys.Where(k =>
            k.Equals("superscript", StringComparison.OrdinalIgnoreCase) ||
            k.Equals("font.superscript", StringComparison.OrdinalIgnoreCase)).ToList();

        superKeys.Should().HaveCount(1,
            "CLAUDE.md mandates one canonical key per semantic value — " +
            "both 'superscript' and 'font.superscript' should not appear simultaneously");
    }

    [Fact]
    public void Bug4_Subscript_Get_ShouldNotContainDuplicateKeys()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Hello", ["subscript"] = "true" });

        var node = _handler.Get("/Sheet1/A1");

        var subKeys = node.Format.Keys.Where(k =>
            k.Equals("subscript", StringComparison.OrdinalIgnoreCase) ||
            k.Equals("font.subscript", StringComparison.OrdinalIgnoreCase)).ToList();

        subKeys.Should().HaveCount(1,
            "CLAUDE.md mandates one canonical key per semantic value — " +
            "both 'subscript' and 'font.subscript' should not appear simultaneously");
    }

    [Fact]
    public void Bug4_Bold_PersistsWithSingleKey()
    {
        // Set bold, reopen, verify only one key survives
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Test", ["bold"] = "true" });
        Reopen();

        var node = _handler.Get("/Sheet1/A1");

        var boldKeys = node.Format.Keys.Where(k =>
            k.Equals("bold", StringComparison.OrdinalIgnoreCase) ||
            k.Equals("font.bold", StringComparison.OrdinalIgnoreCase)).ToList();

        boldKeys.Should().HaveCount(1,
            "after reopen, only one canonical bold key should exist in Format");
    }

    // =====================================================================
    // Bug 5: font.strikethrough=true silently fails
    // =====================================================================

    [Fact]
    public void Bug5_FontStrikethrough_Set_ShouldApplyStrike()
    {
        // "font.strikethrough" is a reasonable alias that users would try.
        // It should work the same as "font.strike" or "strike".
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Struck", ["font.strikethrough"] = "true" });

        var node = _handler.Get("/Sheet1/A1");

        // The cell should have strikethrough applied — check both possible canonical keys
        var hasStrike = node.Format.ContainsKey("font.strike") ||
                        node.Format.ContainsKey("strike") ||
                        node.Format.ContainsKey("font.strikethrough") ||
                        node.Format.ContainsKey("strikethrough");

        hasStrike.Should().BeTrue(
            "font.strikethrough=true should apply strikethrough formatting, " +
            "but it is silently ignored because ExcelStyleManager only recognizes 'strike'");
    }

    [Fact]
    public void Bug5_FontStrikethrough_Set_Persists()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Struck", ["font.strikethrough"] = "true" });
        Reopen();

        var node = _handler.Get("/Sheet1/A1");

        var hasStrike = node.Format.ContainsKey("font.strike") ||
                        node.Format.ContainsKey("strike") ||
                        node.Format.ContainsKey("font.strikethrough") ||
                        node.Format.ContainsKey("strikethrough");

        hasStrike.Should().BeTrue(
            "font.strikethrough should persist after reopen");
    }

    [Fact]
    public void Bug5_Strikethrough_Shorthand_ShouldAlsoWork()
    {
        // "strikethrough" (without font. prefix) should also be accepted as input
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Struck", ["strikethrough"] = "true" });

        var node = _handler.Get("/Sheet1/A1");

        var hasStrike = node.Format.ContainsKey("font.strike") ||
                        node.Format.ContainsKey("strike") ||
                        node.Format.ContainsKey("font.strikethrough") ||
                        node.Format.ContainsKey("strikethrough");

        hasStrike.Should().BeTrue(
            "'strikethrough' shorthand should be treated as alias for 'strike'");
    }

    // =====================================================================
    // Bug 9: Remove("/Sheet1/autofilter") fails
    // =====================================================================

    [Fact]
    public void Bug9_Remove_AutoFilter_ShouldWork()
    {
        // Add autofilter, then remove it
        _handler.Add("/Sheet1/A1", "cell", null, new() { ["value"] = "Header1" });
        _handler.Add("/Sheet1/B1", "cell", null, new() { ["value"] = "Header2" });
        _handler.Add("/Sheet1/A2", "cell", null, new() { ["value"] = "Data1" });
        _handler.Add("/Sheet1/B2", "cell", null, new() { ["value"] = "Data2" });
        _handler.Add("/Sheet1", "autofilter", null, new() { ["range"] = "A1:B2" });

        // Verify autofilter exists via Get (sheet-level node includes autoFilter in Format)
        var sheetNode = _handler.Get("/Sheet1");
        sheetNode.Format.Should().ContainKey("autoFilter");

        // This should not throw — but it does because Remove has no autofilter dispatch
        var act = () => _handler.Remove("/Sheet1/autofilter");
        act.Should().NotThrow(
            "Remove should support /SheetName/autofilter path, " +
            "but currently falls through to cell lookup and throws 'Cell autofilter not found'");
    }

    [Fact]
    public void Bug9_Remove_AutoFilter_ActuallyRemovesIt()
    {
        _handler.Add("/Sheet1/A1", "cell", null, new() { ["value"] = "Header1" });
        _handler.Add("/Sheet1/B1", "cell", null, new() { ["value"] = "Header2" });
        _handler.Add("/Sheet1", "autofilter", null, new() { ["range"] = "A1:B2" });

        _handler.Remove("/Sheet1/autofilter");

        // After removal, Get should not include autoFilter in Format
        var sheetNode = _handler.Get("/Sheet1");
        sheetNode.Format.Should().NotContainKey("autoFilter",
            "autofilter should be gone after Remove");
    }

    [Fact]
    public void Bug9_Remove_AutoFilter_Persists()
    {
        _handler.Add("/Sheet1/A1", "cell", null, new() { ["value"] = "Header1" });
        _handler.Add("/Sheet1/B1", "cell", null, new() { ["value"] = "Header2" });
        _handler.Add("/Sheet1", "autofilter", null, new() { ["range"] = "A1:B2" });

        _handler.Remove("/Sheet1/autofilter");
        Reopen();

        var sheetNode = _handler.Get("/Sheet1");
        sheetNode.Format.Should().NotContainKey("autoFilter",
            "autofilter removal should persist after reopen");
    }
}
