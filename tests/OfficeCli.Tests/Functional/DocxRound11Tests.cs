// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Failing tests for DOCX bugs found in rounds 11-30.
/// Each test documents a specific bug and should fail against the unfixed code.
/// </summary>
public class DocxRound11Tests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private (string path, WordHandler handler) CreateDoc()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return (path, new WordHandler(path, editable: true));
    }

    private WordHandler Reopen(string path)
        => new WordHandler(path, editable: true);

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R11-1: SDT sdtType always reads back as "text" for SdtRun
    //
    // In WordHandler.Query.cs Query("sdt"), the code at lines 700-717 iterates
    // ALL SdtBlock and SdtRun descendants and assigns them sequential paths
    // /body/sdt[1], /body/sdt[2], etc. regardless of their actual position.
    // For inline SDTs (SdtRun inside a <w:p>), the path /body/sdt[N] is wrong —
    // NavigateToElement only looks at direct body children for 'sdt' elements,
    // so an inline SDT inside a paragraph is not found there.
    //
    // Additionally, in WordHandler.Navigation.cs ElementToNode for SdtRun,
    // SdtContentText is checked FIRST (line 643), unlike SdtBlock which checks
    // specific types first (line 608). This means any SdtRun that also carries
    // SdtContentText (e.g., Word may add it for certain SDT configurations)
    // would incorrectly report type "text".
    //
    // Fix: Query must return correct paths for inline SDTs (/body/p[N]/sdt[M]),
    // and NavigateToElement must support the /body/p[N]/sdt[M] path segment.
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R11_1a_QuerySdt_InlineSdtPath_IsNavigable()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Create a paragraph, then add an inline (SdtRun) SDT inside it
        h.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Choose: " });
        h.Add("/body/p[1]", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "dropdown",
            ["alias"] = "Role",
            ["items"] = "Admin,User,Guest",
            ["text"] = "User"
        });

        // Query returns paths for all SDTs
        var sdtNodes = h.Query("sdt");
        sdtNodes.Should().HaveCountGreaterThanOrEqualTo(1,
            "one inline SDT was added and should be queryable");

        var sdtPath = sdtNodes[0].Path;

        // The path returned by Query must be navigable via Get
        // Bug: Query returns /body/sdt[1] for inline SDTs, but NavigateToElement
        // looks for SdtBlock|SdtRun as DIRECT body children — misses inline SDTs
        // that live inside paragraphs.
        var act = () => h.Get(sdtPath);
        act.Should().NotThrow(
            $"Get('{sdtPath}') must succeed — the path was returned by Query. " +
            "Currently fails because Query assigns /body/sdt[N] to inline SDTs " +
            "but NavigateToElement cannot find SdtRun inside a paragraph via that path.");
    }

    [Fact]
    public void Bug_R11_1b_QuerySdt_InlineSdtNode_HasCorrectType()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Date: " });
        h.Add("/body/p[1]", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "date",
            ["alias"] = "BirthDate",
            ["text"] = "2024-06-15"
        });

        var sdtNodes = h.Query("sdt");
        sdtNodes.Should().HaveCountGreaterThanOrEqualTo(1);

        // The path from Query must be usable with Get
        var sdtPath = sdtNodes[0].Path;
        var act = () => h.Get(sdtPath);
        act.Should().NotThrow($"Get('{sdtPath}') must not throw for inline SDT path from Query");

        var node = h.Get(sdtPath);
        node.Format.Should().ContainKey("sdtType",
            "SDT node returned by Get should expose sdtType");
        node.Format["sdtType"].Should().Be("date",
            "inline date SDT should report sdtType=date, not 'text'");
    }

    [Fact]
    public void Bug_R11_1c_QuerySdt_MixedBlockAndInline_AllPathsNavigable()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Block-level SDT
        h.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "dropdown",
            ["alias"] = "Status",
            ["items"] = "Active,Inactive",
            ["text"] = "Active"
        });

        // Inline SDT inside a paragraph
        h.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Combo: " });
        h.Add("/body/p[1]", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "combobox",
            ["alias"] = "Country",
            ["items"] = "US,UK,DE",
            ["text"] = "US"
        });

        var sdtNodes = h.Query("sdt");
        sdtNodes.Should().HaveCountGreaterThanOrEqualTo(2,
            "two SDTs were added (one block, one inline)");

        // EVERY path returned by Query must be navigable via Get
        foreach (var sdt in sdtNodes)
        {
            var act = () => h.Get(sdt.Path);
            act.Should().NotThrow(
                $"Get('{sdt.Path}') must succeed for every path returned by Query — " +
                "inline SDTs (SdtRun inside paragraphs) are currently assigned wrong " +
                "paths /body/sdt[N] that NavigateToElement cannot resolve.");
        }
    }

    [Fact]
    public void Bug_R11_1d_InlineSdtDropdown_DirectGet_WorksAndReturnsDropdown()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Choose: " });

        var addedPath = h.Add("/body/p[1]", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "dropdown",
            ["alias"] = "Role",
            ["items"] = "Admin,User,Guest",
            ["text"] = "User"
        });

        // Direct Get via the Add-returned path should work
        addedPath.Should().Contain("sdt", "Add should return a path containing 'sdt'");
        var node = h.Get(addedPath);
        node.Format.Should().ContainKey("sdtType");
        node.Format["sdtType"].Should().Be("dropdown",
            "inline dropdown SDT accessed directly should return sdtType=dropdown");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R11-2: field type property not used as field-type discriminator
    //
    // When Add("..", "field", ..) is called with properties["type"]="date"
    // (using the generic "type" key instead of "fieldType"), the AddField
    // implementation only checks for "fieldType"/"fieldtype" property keys,
    // not "type". So properties["type"]="date" silently falls through to the
    // default PAGE instruction.
    //
    // This is a usability bug: the "type" property key is a natural alias that
    // users try first. The discriminator should also accept properties["type"]
    // for field type dispatch, consistent with how other handlers accept
    // property aliases.
    //
    // Additionally, the bug affects any field variant not matching the switch —
    // "pagenum" in fieldType works, but "pageNumber" (camelCase) does not.
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R11_2a_FieldWithTypePropertyDate_InsertsDateNotPage()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>());
        // Use "type" property key (not "fieldType") — a natural user expectation
        h.Add("/body/p[1]", "field", null, new Dictionary<string, string>
        {
            ["type"] = "date"
        });

        var fields = h.Query("field");
        fields.Should().HaveCountGreaterThanOrEqualTo(1,
            "a field was added to the document");

        var fieldNode = h.Get(fields[0].Path);
        var instrText = fieldNode.Format.ContainsKey("instruction")
            ? fieldNode.Format["instruction"]?.ToString() ?? ""
            : fieldNode.Text ?? "";

        instrText.ToUpperInvariant().Should().Contain("DATE",
            "properties[\"type\"]=\"date\" should insert a DATE field instruction. " +
            "Bug: AddField only checks 'fieldType'/'fieldtype' property keys, not 'type', " +
            "so this silently defaults to PAGE");
        instrText.ToUpperInvariant().Should().NotContain("PAGE",
            "a date field must not produce PAGE instruction");
    }

    [Fact]
    public void Bug_R11_2b_FieldWithTypePropertyNumPages_InsertsNumPagesNotPage()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>());
        h.Add("/body/p[1]", "field", null, new Dictionary<string, string>
        {
            ["type"] = "numpages"
        });

        var fields = h.Query("field");
        fields.Should().HaveCountGreaterThanOrEqualTo(1);

        var fieldNode = h.Get(fields[0].Path);
        var instrText = fieldNode.Format.ContainsKey("instruction")
            ? fieldNode.Format["instruction"]?.ToString() ?? ""
            : fieldNode.Text ?? "";

        instrText.ToUpperInvariant().Should().Contain("NUMPAGES",
            "properties[\"type\"]=\"numpages\" should insert NUMPAGES instruction, not PAGE");
    }

    [Fact]
    public void Bug_R11_2c_FieldWithTypePropertyAuthor_InsertsAuthorInstruction()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>());
        // Use "type" property with a value that does NOT accidentally match the default PAGE
        h.Add("/body/p[1]", "field", null, new Dictionary<string, string>
        {
            ["type"] = "author"
        });

        var fields = h.Query("field");
        fields.Should().HaveCountGreaterThanOrEqualTo(1,
            "an author field was added");

        var fieldNode = h.Get(fields[0].Path);
        var instrText = fieldNode.Format.ContainsKey("instruction")
            ? fieldNode.Format["instruction"]?.ToString() ?? ""
            : fieldNode.Text ?? "";

        instrText.ToUpperInvariant().Should().Contain("AUTHOR",
            "properties[\"type\"]=\"author\" should produce AUTHOR field instruction. " +
            "Bug: AddField only checks 'fieldType'/'fieldtype' keys, not 'type', " +
            "so this falls through to PAGE default.");
        instrText.ToUpperInvariant().Should().NotContain("PAGE",
            "author field must not default to PAGE");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R11-3: Get /body/oMathPara[N] throws "path not found"
    //
    // Query("equation") correctly returns paths like /body/oMathPara[1].
    // However, Get("/body/oMathPara[1]") fails because NavigateToElement
    // processes segment "oMathPara" via the generic fallback:
    //   current.ChildElements.Where(e => e.LocalName == "oMathPara")
    // When the oMathPara is wrapped inside a <w:p> paragraph, the body's
    // direct children don't include oMathPara, so navigation returns null.
    //
    // Fix: NavigateToElement must handle "oMathPara" as a special case,
    // extracting oMathPara elements from paragraphs (as Query does).
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R11_3a_GetOmathPara_ReturnsEquationNode()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Add a display equation (produces oMathPara wrapped in w:p)
        h.Add("/body", "equation", null, new Dictionary<string, string>
        {
            ["formula"] =@"E = mc^2"
        });

        // Query returns the correct path
        var equations = h.Query("equation");
        equations.Should().HaveCountGreaterThanOrEqualTo(1,
            "the equation should be queryable");

        var equationPath = equations[0].Path;
        equationPath.Should().Contain("oMathPara",
            "display equations should be referenced by oMathPara path");

        // Get should not throw and should return the equation node
        var act = () => h.Get(equationPath);
        act.Should().NotThrow(
            $"Get('{equationPath}') should succeed — path was returned by Query");

        var node = h.Get(equationPath);
        node.Should().NotBeNull();
        node.Type.Should().Be("equation",
            "the node type should be 'equation'");
        node.Text.Should().NotBeNullOrEmpty(
            "the node text should contain the LaTeX representation");
    }

    [Fact]
    public void Bug_R11_3b_GetOmathPara_SecondEquation_ReturnsCorrectNode()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "equation", null, new Dictionary<string, string> { ["formula"] =@"a + b = c" });
        h.Add("/body", "equation", null, new Dictionary<string, string> { ["formula"] =@"x^2 + y^2 = r^2" });

        var equations = h.Query("equation");
        equations.Should().HaveCountGreaterThanOrEqualTo(2,
            "two equations were added");

        // Both paths must be gettable
        foreach (var eq in equations)
        {
            var act = () => h.Get(eq.Path);
            act.Should().NotThrow($"Get('{eq.Path}') should succeed");
        }

        var node2 = h.Get(equations[1].Path);
        node2.Type.Should().Be("equation");
        node2.Text.Should().NotBeNullOrEmpty();
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R11-4: Remove /header[N] crashes with NullReferenceException
    //
    // WordHandler.Remove() calls NavigateToElement("/header[1]") which returns
    // the Header OpenXML element (the root of the HeaderPart XML tree).
    // This element has no Parent (it is the root of a package part), so
    // element.Remove() either throws a NullReferenceException or silently
    // does nothing — leaving the header relationship intact.
    //
    // Fix: Remove("/header[N]") should delete the HeaderPart and its
    // HeaderReference from the SectionProperties, not call element.Remove().
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R11_4a_RemoveHeader_DoesNotThrow()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "header", null, new Dictionary<string, string>
        {
            ["text"] = "My Header"
        });

        // Verify the header was added
        var root = h.Get("/");
        root.Children.Should().Contain(c => c.Type == "header",
            "header was added and should appear as a child of root");

        // Remove should not throw
        var act = () => h.Remove("/header[1]");
        act.Should().NotThrow(
            "Remove('/header[1]') must not throw a NullReferenceException");
    }

    [Fact]
    public void Bug_R11_4b_RemoveHeader_HeaderIsActuallyGone()
    {
        var (path, handler) = CreateDoc();
        handler.Add("/body", "header", null, new Dictionary<string, string>
        {
            ["text"] = "Temporary Header"
        });
        handler.Remove("/header[1]");
        handler.Dispose();

        using var h2 = Reopen(path);
        var root = h2.Get("/");
        root.Children.Should().NotContain(c => c.Type == "header",
            "after Remove('/header[1]') the header must not appear on reopen");
    }

    [Fact]
    public void Bug_R11_4c_RemoveFooter_DoesNotThrow()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "footer", null, new Dictionary<string, string>
        {
            ["text"] = "Page {PAGE}"
        });

        var act = () => h.Remove("/footer[1]");
        act.Should().NotThrow(
            "Remove('/footer[1]') must not throw a NullReferenceException");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R11-5: SDT body paragraphs counted in /body/p[N] index
    //
    // GetBodyElements() flattens paragraphs inside SdtBlock containers and
    // counts them together with normal body paragraphs. This means:
    //   - A normal paragraph after an SdtBlock (which contains 1 paragraph)
    //     gets index /body/p[2] from navigation.
    //   - But Query("/body", "paragraph") sees the SdtBlock in body.ChildElements,
    //     skips it (not a Paragraph), and assigns the following normal paragraph
    //     index 1 → path /body/p[1].
    //
    // The mismatch means: Query says /body/p[1] but Get("/body/p[1]") returns
    // the SDT's inner paragraph (index 1 in the flattened list), not the normal
    // paragraph that Query found.
    //
    // Fix: GetBodyElements() must NOT include SDT inner paragraphs when counting
    // /body/p[N] indices, OR Query must use the same flattening as navigation.
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R11_5a_ParagraphAfterSdt_HasCorrectIndex()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Add a block-level SDT (it will contain one inner paragraph)
        h.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "text",
            ["text"] = "SDT content paragraph"
        });

        // Add a normal paragraph AFTER the SDT
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Normal paragraph after SDT"
        });

        // Query should find exactly one paragraph (the normal one, not the SDT inner one)
        var paragraphs = h.Query("paragraph");
        // There may be 1 or 2 depending on blank doc; find the one we added
        var normalParaNode = paragraphs.FirstOrDefault(p =>
            p.Text?.Contains("Normal paragraph after SDT") == true);

        normalParaNode.Should().NotBeNull(
            "Query should find the normal paragraph we added after the SDT");

        var normalParaPath = normalParaNode!.Path;

        // Get the node via the path returned by Query — must return the same paragraph
        var getNode = h.Get(normalParaPath);
        getNode.Should().NotBeNull();
        getNode.Text.Should().Contain("Normal paragraph after SDT",
            $"Get('{normalParaPath}') must return the normal paragraph, " +
            "not an SDT inner paragraph — indices must be consistent between Query and Get");
    }

    [Fact]
    public void Bug_R11_5b_QueryAndGetParagraphIndex_AreConsistent()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Add two SDTs with inner paragraphs
        h.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "text",
            ["text"] = "First SDT"
        });
        h.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "text",
            ["text"] = "Second SDT"
        });

        // Add a distinguishable paragraph after both SDTs
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Paragraph after two SDTs"
        });

        var paragraphs = h.Query("paragraph");
        var targetNode = paragraphs.FirstOrDefault(p =>
            p.Text?.Contains("Paragraph after two SDTs") == true);

        targetNode.Should().NotBeNull(
            "Query must find the paragraph added after two SDTs");

        var gotten = h.Get(targetNode!.Path);
        gotten.Text.Should().Contain("Paragraph after two SDTs",
            $"Get('{targetNode.Path}') must return the same paragraph Query found — " +
            "SDT inner paragraphs must not shift the /body/p[N] index");
    }
}
