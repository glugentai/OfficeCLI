// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for DOCX: each test creates a blank file, adds elements,
/// queries them, and modifies them — exercising the full Create→Add→Get→Set lifecycle.
/// </summary>
public class WordFunctionalTests : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public WordFunctionalTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== DOCX Hyperlinks ====================

    [Fact]
    public void Hyperlink_Lifecycle()
    {
        // 1. Add paragraph + hyperlink
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>());
        var path = _handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://first.com",
            ["text"] = "Click here"
        });
        path.Should().Be("/body/p[1]/hyperlink[1]");

        // 2. Get + Verify type, url, text
        var node = _handler.Get("/body/p[1]/hyperlink[1]");
        node.Type.Should().Be("hyperlink");
        node.Text.Should().Be("Click here");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://first.com");

        // 3. Verify paragraph text contains link text
        var para = _handler.Get("/body/p[1]");
        para.Text.Should().Contain("Click here");

        // 4. Query + Verify
        var results = _handler.Query("hyperlink");
        results.Should().Contain(n => n.Type == "hyperlink" && n.Text == "Click here");

        // 5. Set (update URL via run) + Verify
        _handler.Set("/body/p[1]/r[1]", new Dictionary<string, string> { ["link"] = "https://updated.com" });
        node = _handler.Get("/body/p[1]/hyperlink[1]");
        ((string)node.Format["link"]).Should().StartWith("https://updated.com");
    }

    // ==================== DOCX Numbering / Lists ====================

    [Fact]
    public void ListStyle_Bullet_Lifecycle()
    {
        // 1. Add paragraph with bullet list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Bullet item 1",
            ["liststyle"] = "bullet"
        });

        // 2. Get + Verify all numbering properties
        var node = _handler.Get("/body/p[1]");
        node.Text.Should().Be("Bullet item 1");
        node.Format.Should().ContainKey("numid");
        node.Format.Should().ContainKey("numlevel");
        node.Format.Should().ContainKey("listStyle");
        node.Format.Should().ContainKey("numFmt");
        node.Format.Should().ContainKey("start");
        ((int)node.Format["numlevel"]).Should().Be(0);
        ((string)node.Format["listStyle"]).Should().Be("bullet");
        ((string)node.Format["numFmt"]).Should().Be("bullet");
        ((int)node.Format["start"]).Should().Be(1);

        // 3. Set — change numlevel
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["numlevel"] = "1" });

        // 4. Get + Verify level changed
        node = _handler.Get("/body/p[1]");
        ((int)node.Format["numlevel"]).Should().Be(1);

        // 5. Persist + Verify
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        node.Text.Should().Be("Bullet item 1");
        ((string)node.Format["listStyle"]).Should().Be("bullet");
        ((int)node.Format["numlevel"]).Should().Be(1);
    }

    [Fact]
    public void ListStyle_Ordered_Lifecycle()
    {
        // 1. Add paragraph with ordered list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Step 1",
            ["liststyle"] = "numbered"
        });

        // 2. Get + Verify
        var node = _handler.Get("/body/p[1]");
        node.Text.Should().Be("Step 1");
        node.Format.Should().ContainKey("numid");
        node.Format.Should().ContainKey("listStyle");
        node.Format.Should().ContainKey("numFmt");
        ((string)node.Format["listStyle"]).Should().Be("ordered");
        ((string)node.Format["numFmt"]).Should().Be("decimal");

        // 3. Set — change to bullet
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["liststyle"] = "bullet" });

        // 4. Get + Verify changed
        node = _handler.Get("/body/p[1]");
        ((string)node.Format["listStyle"]).Should().Be("bullet");

        // 5. Persist + Verify
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        ((string)node.Format["listStyle"]).Should().Be("bullet");
    }

    [Fact]
    public void ListStyle_None_RemovesNumbering()
    {
        // 1. Add paragraph with bullet list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Will lose numbering",
            ["liststyle"] = "bullet"
        });
        var node = _handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("numid");

        // 2. Set listStyle=none to remove numbering
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["liststyle"] = "none" });

        // 3. Get + Verify numbering removed
        node = _handler.Get("/body/p[1]");
        node.Text.Should().Be("Will lose numbering");
        node.Format.Should().NotContainKey("numid");
        node.Format.Should().NotContainKey("listStyle");

        // 4. Persist + Verify still removed
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        node.Format.Should().NotContainKey("numid");
    }

    [Fact]
    public void ListStyle_Continuation_SharesNumId()
    {
        // 1. Add first bullet paragraph — creates new numbering
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Item A",
            ["liststyle"] = "bullet"
        });
        var numId1 = (int)_handler.Get("/body/p[1]").Format["numid"];

        // 2. Add second consecutive bullet paragraph — should reuse same numId
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Item B",
            ["liststyle"] = "bullet"
        });
        var numId2 = (int)_handler.Get("/body/p[2]").Format["numid"];

        numId2.Should().Be(numId1, "consecutive same-type list items should share numId");

        // 3. Persist + Verify continuation survives reopen
        var handler2 = Reopen();
        var n1 = handler2.Get("/body/p[1]");
        var n2 = handler2.Get("/body/p[2]");
        ((int)n1.Format["numid"]).Should().Be((int)n2.Format["numid"]);
    }

    [Fact]
    public void ListStyle_StartValue_Lifecycle()
    {
        // 1. Add ordered list starting from 5
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Step 5",
            ["liststyle"] = "numbered",
            ["start"] = "5"
        });

        // 2. Get + Verify start value
        var node = _handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("start");
        ((int)node.Format["start"]).Should().Be(5);
        ((string)node.Format["listStyle"]).Should().Be("ordered");

        // 3. Set — change start value via Set
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["start"] = "10" });

        // 4. Get + Verify
        node = _handler.Get("/body/p[1]");
        ((int)node.Format["start"]).Should().Be(10);

        // 5. Persist + Verify
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        ((int)node.Format["start"]).Should().Be(10);
    }

    [Fact]
    public void ListStyle_NumId_RawAccess()
    {
        // 1. Add paragraph with listStyle
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Raw item",
            ["liststyle"] = "bullet"
        });

        // 2. Get the numid back
        var numId = (int)_handler.Get("/body/p[1]").Format["numid"];
        numId.Should().BeGreaterThan(0);

        // 3. Add another paragraph using the raw numid
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Same list",
            ["numid"] = numId.ToString(),
            ["numlevel"] = "0"
        });

        // 4. Get + Verify shared numid
        var node2 = _handler.Get("/body/p[2]");
        ((int)node2.Format["numid"]).Should().Be(numId);
        ((int)node2.Format["numlevel"]).Should().Be(0);

        // 5. Persist + Verify
        var handler2 = Reopen();
        node2 = handler2.Get("/body/p[2]");
        ((int)node2.Format["numid"]).Should().Be(numId);
    }

    [Fact]
    public void ListStyle_NineLevels_Supported()
    {
        // 1. Add a bullet list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Deep nesting",
            ["liststyle"] = "bullet"
        });

        // 2. Set numlevel to 8 (0-based, 9th level)
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["numlevel"] = "8" });

        // 3. Get + Verify level 8 works
        var node = _handler.Get("/body/p[1]");
        ((int)node.Format["numlevel"]).Should().Be(8);
        ((string)node.Format["listStyle"]).Should().Be("bullet");

        // 4. Persist + Verify
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        ((int)node.Format["numlevel"]).Should().Be(8);
    }

    [Fact]
    public void ListStyle_NumFmt_ReturnsSpecificFormat()
    {
        // 1. Add ordered list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Level 0",
            ["liststyle"] = "numbered"
        });

        // 2. Verify level 0 = decimal
        var node = _handler.Get("/body/p[1]");
        ((string)node.Format["numFmt"]).Should().Be("decimal");

        // 3. Set to level 1
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["numlevel"] = "1" });

        // 4. Verify level 1 = lowerLetter
        node = _handler.Get("/body/p[1]");
        ((string)node.Format["numFmt"]).Should().Be("lowerLetter");

        // 5. Set to level 2
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["numlevel"] = "2" });

        // 6. Verify level 2 = lowerRoman
        node = _handler.Get("/body/p[1]");
        ((string)node.Format["numFmt"]).Should().Be("lowerRoman");
    }

    [Fact]
    public void ListStyle_Query_FilterByListStyle()
    {
        // 1. Add mixed paragraphs
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Normal paragraph"
        });
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Bullet item",
            ["liststyle"] = "bullet"
        });
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Ordered item",
            ["liststyle"] = "numbered"
        });

        // 2. Query + Verify filtering
        var bullets = _handler.Query("paragraph[liststyle=bullet]");
        bullets.Should().HaveCount(1);
        bullets[0].Text.Should().Be("Bullet item");

        var ordered = _handler.Query("paragraph[liststyle=ordered]");
        ordered.Should().HaveCount(1);
        ordered[0].Text.Should().Be("Ordered item");

        // 3. Query by numid
        var numId = (int)_handler.Get("/body/p[2]").Format["numid"];
        var byNumId = _handler.Query($"paragraph[numid={numId}]");
        byNumId.Should().ContainSingle(n => n.Text == "Bullet item");
    }

    [Fact]
    public void Hyperlink_Persist_SurvivesReopenFile()
    {
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>());
        _handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://original.com",
            ["text"] = "My link"
        });
        _handler.Set("/body/p[1]/r[1]", new Dictionary<string, string> { ["link"] = "https://persist.com" });

        var handler2 = Reopen();
        var node = handler2.Get("/body/p[1]/hyperlink[1]");
        node.Text.Should().Be("My link");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://persist.com");
    }

    // ==================== Table Row Add Lifecycle ====================

    [Fact]
    public void AddRow_FullLifecycle()
    {
        // 1. Create table
        _handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });

        // 2. Add row with cell text
        var path = _handler.Add("/body/tbl[1]", "row", null, new() { ["c1"] = "Hello", ["c2"] = "World" });
        path.Should().Be("/body/tbl[1]/tr[2]");

        // 3. Get + Verify
        var cell1 = _handler.Get("/body/tbl[1]/tr[2]/tc[1]");
        cell1.Text.Should().Be("Hello");
        var cell2 = _handler.Get("/body/tbl[1]/tr[2]/tc[2]");
        cell2.Text.Should().Be("World");

        // 4. Set (modify cell text and formatting)
        _handler.Set("/body/tbl[1]/tr[2]/tc[1]", new() { ["text"] = "Modified", ["bold"] = "true" });

        // 5. Get + Verify again
        cell1 = _handler.Get("/body/tbl[1]/tr[2]/tc[1]");
        cell1.Text.Should().Be("Modified");

        // 6. Persistence: Reopen + Verify
        Reopen();
        cell1 = _handler.Get("/body/tbl[1]/tr[2]/tc[1]");
        cell1.Text.Should().Be("Modified");
        _handler.Get("/body/tbl[1]/tr[2]/tc[2]").Text.Should().Be("World");
    }

    [Fact]
    public void AddRow_AtIndex_FullLifecycle()
    {
        // 1. Create table with 2 rows
        _handler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "1" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "First" });
        _handler.Set("/body/tbl[1]/tr[2]/tc[1]", new() { ["text"] = "Last" });

        // 2. Add row at index 1 (between First and Last)
        var path = _handler.Add("/body/tbl[1]", "row", 1, new() { ["c1"] = "Middle" });
        path.Should().Be("/body/tbl[1]/tr[2]");

        // 3. Get + Verify insertion position
        _handler.Get("/body/tbl[1]/tr[1]/tc[1]").Text.Should().Be("First");
        _handler.Get("/body/tbl[1]/tr[2]/tc[1]").Text.Should().Be("Middle");
        _handler.Get("/body/tbl[1]/tr[3]/tc[1]").Text.Should().Be("Last");

        // 4. Set (modify inserted row)
        _handler.Set("/body/tbl[1]/tr[2]/tc[1]", new() { ["text"] = "Center" });

        // 5. Get + Verify
        _handler.Get("/body/tbl[1]/tr[2]/tc[1]").Text.Should().Be("Center");

        // 6. Persistence
        Reopen();
        _handler.Get("/body/tbl[1]/tr[2]/tc[1]").Text.Should().Be("Center");
    }

    // ==================== Table Cell Add Lifecycle ====================

    [Fact]
    public void AddCell_FullLifecycle()
    {
        // 1. Create table
        _handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        // 2. Add cell
        var path = _handler.Add("/body/tbl[1]/tr[1]", "cell", null, new() { ["text"] = "NewCell" });
        path.Should().Be("/body/tbl[1]/tr[1]/tc[2]");

        // 3. Get + Verify
        var cell = _handler.Get("/body/tbl[1]/tr[1]/tc[2]");
        cell.Text.Should().Be("NewCell");

        // 4. Set (modify)
        _handler.Set("/body/tbl[1]/tr[1]/tc[2]", new() { ["text"] = "Updated", ["shd"] = "FF0000" });

        // 5. Get + Verify
        cell = _handler.Get("/body/tbl[1]/tr[1]/tc[2]");
        cell.Text.Should().Be("Updated");

        // 6. Persistence
        Reopen();
        cell = _handler.Get("/body/tbl[1]/tr[1]/tc[2]");
        cell.Text.Should().Be("Updated");
    }

    [Fact]
    public void AddCell_AtIndex_FullLifecycle()
    {
        // 1. Create table with 2 cells
        _handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "A" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[2]", new() { ["text"] = "C" });

        // 2. Add cell at index 1 (between A and C)
        var path = _handler.Add("/body/tbl[1]/tr[1]", "cell", 1, new() { ["text"] = "B" });
        path.Should().Be("/body/tbl[1]/tr[1]/tc[2]");

        // 3. Get + Verify order
        _handler.Get("/body/tbl[1]/tr[1]/tc[1]").Text.Should().Be("A");
        _handler.Get("/body/tbl[1]/tr[1]/tc[2]").Text.Should().Be("B");
        _handler.Get("/body/tbl[1]/tr[1]/tc[3]").Text.Should().Be("C");

        // 4. Set (modify inserted cell)
        _handler.Set("/body/tbl[1]/tr[1]/tc[2]", new() { ["text"] = "Beta" });

        // 5. Get + Verify
        _handler.Get("/body/tbl[1]/tr[1]/tc[2]").Text.Should().Be("Beta");

        // 6. Persistence
        Reopen();
        _handler.Get("/body/tbl[1]/tr[1]/tc[1]").Text.Should().Be("A");
        _handler.Get("/body/tbl[1]/tr[1]/tc[2]").Text.Should().Be("Beta");
        _handler.Get("/body/tbl[1]/tr[1]/tc[3]").Text.Should().Be("C");
    }

    // ==================== Document Core Properties Lifecycle ====================

    [Fact]
    public void CoreProperties_FullLifecycle()
    {
        // 1. Set properties
        _handler.Set("/", new() { ["title"] = "My Document", ["author"] = "Test User", ["subject"] = "Testing" });

        // 2. Get + Verify
        var root = _handler.Get("/");
        ((string)root.Format["title"]).Should().Be("My Document");
        ((string)root.Format["author"]).Should().Be("Test User");
        ((string)root.Format["subject"]).Should().Be("Testing");

        // 3. Set (modify)
        _handler.Set("/", new() { ["title"] = "Updated Title", ["keywords"] = "test,docx" });

        // 4. Get + Verify
        root = _handler.Get("/");
        ((string)root.Format["title"]).Should().Be("Updated Title");
        ((string)root.Format["keywords"]).Should().Be("test,docx");

        // 5. Persistence
        Reopen();
        root = _handler.Get("/");
        ((string)root.Format["title"]).Should().Be("Updated Title");
        ((string)root.Format["author"]).Should().Be("Test User");
        ((string)root.Format["keywords"]).Should().Be("test,docx");
    }

    // ==================== Paragraph Indent Lifecycle ====================

    [Fact]
    public void ParagraphIndent_FullLifecycle()
    {
        // 1. Add paragraph with left indent
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Indented", ["leftindent"] = "720" });

        // 2. Get + Verify
        var node = _handler.Get("/body/p[1]");
        ((string)node.Format["leftIndent"]).Should().Be("720");

        // 3. Set (modify + add right indent and hanging)
        _handler.Set("/body/p[1]", new() { ["leftindent"] = "1440", ["rightindent"] = "720", ["hanging"] = "360" });

        // 4. Get + Verify
        node = _handler.Get("/body/p[1]");
        ((string)node.Format["leftIndent"]).Should().Be("1440");
        ((string)node.Format["rightIndent"]).Should().Be("720");
        ((string)node.Format["hangingIndent"]).Should().Be("360");

        // 5. Persistence
        Reopen();
        node = _handler.Get("/body/p[1]");
        ((string)node.Format["leftIndent"]).Should().Be("1440");
        ((string)node.Format["rightIndent"]).Should().Be("720");
        ((string)node.Format["hangingIndent"]).Should().Be("360");
    }

    // ==================== Superscript/Subscript Lifecycle ====================

    [Fact]
    public void Superscript_FullLifecycle()
    {
        // 1. Add paragraph + run with superscript
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "E=mc" });
        _handler.Add("/body/p[1]", "run", null, new() { ["text"] = "2", ["superscript"] = "true" });

        // 2. Get + Verify
        var run = _handler.Get("/body/p[1]/r[2]");
        run.Text.Should().Be("2");
        ((bool)run.Format["superscript"]).Should().BeTrue();

        // 3. Set (change to subscript)
        _handler.Set("/body/p[1]/r[2]", new() { ["superscript"] = "false", ["subscript"] = "true" });

        // 4. Get + Verify (subscript on, superscript gone)
        run = _handler.Get("/body/p[1]/r[2]");
        run.Format.Should().NotContainKey("superscript");
        ((bool)run.Format["subscript"]).Should().BeTrue();

        // 5. Persistence
        Reopen();
        run = _handler.Get("/body/p[1]/r[2]");
        ((bool)run.Format["subscript"]).Should().BeTrue();
    }

    // ==================== Paragraph Flow Control Lifecycle ====================

    [Fact]
    public void ParagraphFlowControl_FullLifecycle()
    {
        // 1. Add paragraph with flow control
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Keep with next", ["keepnext"] = "true" });

        // 2. Get + Verify
        var node = _handler.Get("/body/p[1]");
        ((bool)node.Format["keepNext"]).Should().BeTrue();

        // 3. Set (add more flow controls)
        _handler.Set("/body/p[1]", new() { ["keeplines"] = "true", ["pagebreakbefore"] = "true", ["widowcontrol"] = "true" });

        // 4. Get + Verify
        node = _handler.Get("/body/p[1]");
        ((bool)node.Format["keepNext"]).Should().BeTrue();
        ((bool)node.Format["keepLines"]).Should().BeTrue();
        ((bool)node.Format["pageBreakBefore"]).Should().BeTrue();
        ((bool)node.Format["widowControl"]).Should().BeTrue();

        // 5. Set (remove)
        _handler.Set("/body/p[1]", new() { ["keepnext"] = "false" });

        // 6. Verify removed
        node = _handler.Get("/body/p[1]");
        node.Format.Should().NotContainKey("keepNext");

        // 7. Persistence
        Reopen();
        node = _handler.Get("/body/p[1]");
        ((bool)node.Format["keepLines"]).Should().BeTrue();
        ((bool)node.Format["pageBreakBefore"]).Should().BeTrue();
    }

    // ==================== Section Break Lifecycle ====================

    [Fact]
    public void SectionBreak_FullLifecycle()
    {
        // 1. Add content + section break
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Section 1" });
        var path = _handler.Add("/body", "section", null, new() { ["type"] = "nextPage" });
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Section 2" });

        // 2. Get + Verify
        path.Should().Be("/section[1]");
        var sec = _handler.Get("/section[1]");
        sec.Type.Should().Be("section");
        ((string)sec.Format["type"]).Should().Be("nextPage");
        sec.Format.Should().ContainKey("pageWidth");
        sec.Format.Should().ContainKey("pageHeight");

        // 3. Set (modify section properties)
        _handler.Set("/section[1]", new() { ["type"] = "continuous", ["margintop"] = "720" });

        // 4. Get + Verify
        sec = _handler.Get("/section[1]");
        ((string)sec.Format["type"]).Should().Be("continuous");
        ((int)sec.Format["marginTop"]).Should().Be(720);

        // 5. Persistence
        Reopen();
        sec = _handler.Get("/section[1]");
        ((string)sec.Format["type"]).Should().Be("continuous");
        ((int)sec.Format["marginTop"]).Should().Be(720);
    }

    // ==================== Footnote Lifecycle ====================

    [Fact]
    public void Footnote_FullLifecycle()
    {
        // 1. Add paragraph
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Some text" });

        // 2. Add footnote
        var path = _handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "This is a footnote" });
        path.Should().Be("/footnote[1]");

        // 3. Get + Verify
        var fn = _handler.Get("/footnote[1]");
        fn.Type.Should().Be("footnote");
        fn.Text.Should().Contain("This is a footnote");

        // 4. Set (modify text)
        _handler.Set("/footnote[1]", new() { ["text"] = "Updated footnote" });

        // 5. Get + Verify (new text present, old text gone)
        fn = _handler.Get("/footnote[1]");
        fn.Text.Should().Contain("Updated footnote");
        fn.Text.Should().NotContain("This is a footnote");

        // 6. Persistence
        Reopen();
        fn = _handler.Get("/footnote[1]");
        fn.Type.Should().Be("footnote");
        fn.Text.Should().Contain("Updated footnote");
    }

    // ==================== Endnote Lifecycle ====================

    [Fact]
    public void Endnote_FullLifecycle()
    {
        // 1. Add paragraph
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Some text" });

        // 2. Add endnote
        var path = _handler.Add("/body/p[1]", "endnote", null, new() { ["text"] = "This is an endnote" });
        path.Should().Be("/endnote[1]");

        // 3. Get + Verify
        var en = _handler.Get("/endnote[1]");
        en.Type.Should().Be("endnote");
        en.Text.Should().Contain("This is an endnote");

        // 4. Set (modify text)
        _handler.Set("/endnote[1]", new() { ["text"] = "Updated endnote" });

        // 5. Get + Verify (new text present, old text gone)
        en = _handler.Get("/endnote[1]");
        en.Text.Should().Contain("Updated endnote");
        en.Text.Should().NotContain("This is an endnote");

        // 6. Persistence
        Reopen();
        en = _handler.Get("/endnote[1]");
        en.Type.Should().Be("endnote");
        en.Text.Should().Contain("Updated endnote");
    }

    // ==================== TOC Lifecycle ====================

    [Fact]
    public void TOC_FullLifecycle()
    {
        // 1. Add headings
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Chapter 1", ["style"] = "Heading1" });
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Chapter 2", ["style"] = "Heading1" });

        // 2. Add TOC
        var path = _handler.Add("/body", "toc", null, new() { ["levels"] = "1-3" });
        path.Should().Be("/toc[1]");

        // 3. Get + Verify
        var toc = _handler.Get("/toc[1]");
        toc.Type.Should().Be("toc");
        ((string)toc.Format["levels"]).Should().Be("1-3");
        ((bool)toc.Format["hyperlinks"]).Should().BeTrue();
        ((bool)toc.Format["pageNumbers"]).Should().BeTrue();

        // 4. Set (modify)
        _handler.Set("/toc[1]", new() { ["levels"] = "1-2", ["pagenumbers"] = "false" });

        // 5. Get + Verify
        toc = _handler.Get("/toc[1]");
        ((string)toc.Format["levels"]).Should().Be("1-2");
        ((bool)toc.Format["pageNumbers"]).Should().BeFalse();

        // 6. Persistence
        Reopen();
        toc = _handler.Get("/toc[1]");
        toc.Type.Should().Be("toc");
        ((string)toc.Format["levels"]).Should().Be("1-2");
    }

    // ==================== Style Creation Lifecycle ====================

    [Fact]
    public void StyleCreation_FullLifecycle()
    {
        // 1. Create style
        var path = _handler.Add("/body", "style", null, new()
        {
            ["name"] = "MyCustomStyle", ["id"] = "MyCustomStyle",
            ["font"] = "Arial", ["size"] = "14", ["bold"] = "true", ["color"] = "FF0000",
            ["alignment"] = "center", ["spacebefore"] = "240"
        });
        path.Should().Be("/styles/MyCustomStyle");

        // 2. Get + Verify style properties
        var style = _handler.Get("/styles/MyCustomStyle");
        style.Type.Should().Be("style");
        ((string)style.Format["font"]).Should().Be("Arial");
        ((int)style.Format["size"]).Should().Be(14);
        ((bool)style.Format["bold"]).Should().BeTrue();
        ((string)style.Format["color"]).Should().Be("FF0000");
        ((string)style.Format["alignment"]).Should().Be("center");
        ((string)style.Format["spaceBefore"]).Should().Be("240");

        // 3. Set (modify style)
        _handler.Set("/styles/MyCustomStyle", new() { ["font"] = "Calibri", ["size"] = "12", ["bold"] = "false" });

        // 4. Get + Verify
        style = _handler.Get("/styles/MyCustomStyle");
        ((string)style.Format["font"]).Should().Be("Calibri");
        ((int)style.Format["size"]).Should().Be(12);
        style.Format.Should().NotContainKey("bold");

        // 5. Apply style to paragraph + verify
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Styled text", ["style"] = "MyCustomStyle" });
        var node = _handler.Get("/body/p[1]");
        node.Style.Should().Be("MyCustomStyle");

        // 6. Persistence
        Reopen();
        style = _handler.Get("/styles/MyCustomStyle");
        ((string)style.Format["font"]).Should().Be("Calibri");
        node = _handler.Get("/body/p[1]");
        node.Style.Should().Be("MyCustomStyle");
    }
}
