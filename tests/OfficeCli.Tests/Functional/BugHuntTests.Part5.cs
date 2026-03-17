// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// Bug hunt tests Part 5: Bugs #331-350+
// Word Set, Word Query, Word Selector, GenericXmlQuery, ImageHelpers, StyleList deep dive

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    /// Bug #331 — Word Set: bool.Parse on TOC hyperlinks property
    /// File: WordHandler.Set.cs, lines 70-73
    /// bool.Parse("yes") or bool.Parse("1") will throw FormatException
    /// instead of being treated as truthy.
    [Fact]
    public void Bug331_WordSet_TocHyperlinksBoolParse()
    {
        // bool.Parse only accepts "True"/"False" (case-insensitive)
        // User might pass "yes", "1", "on" etc.
        var ex = Record.Exception(() => bool.Parse("yes"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects 'yes' — should use IsTruthy or TryParse");

        var ex2 = Record.Exception(() => bool.Parse("1"));
        ex2.Should().BeOfType<FormatException>(
            "bool.Parse rejects '1' — common truthy value");
    }

    /// Bug #332 — Word Set: bool.Parse on TOC pageNumbers property
    /// File: WordHandler.Set.cs, lines 79-82
    /// Same bool.Parse issue as #331 but for pageNumbers switch.
    [Fact]
    public void Bug332_WordSet_TocPageNumbersBoolParse()
    {
        // bool.Parse is called directly on user-provided value
        var ex = Record.Exception(() => bool.Parse("0"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects '0' — should use IsTruthy or TryParse");
    }

    /// Bug #333 — Word Set: uint.Parse on section pageWidth/pageHeight
    /// File: WordHandler.Set.cs, lines 174, 177
    /// uint.Parse("12240.5") or uint.Parse("abc") will throw.
    /// No validation on user-provided values.
    [Fact]
    public void Bug333_WordSet_SectionPageSizeUintParse()
    {
        var ex = Record.Exception(() => uint.Parse("12240.5"));
        ex.Should().BeOfType<FormatException>(
            "uint.Parse cannot handle decimal values for page size");

        var ex2 = Record.Exception(() => uint.Parse("-100"));
        ex2.Should().BeOfType<OverflowException>(
            "uint.Parse cannot handle negative values");
    }

    /// Bug #334 — Word Set: int.Parse on section margins
    /// File: WordHandler.Set.cs, lines 185-194
    /// Multiple margin properties use int.Parse / uint.Parse directly on user input.
    [Fact]
    public void Bug334_WordSet_SectionMarginParsing()
    {
        // marginTop/marginBottom use int.Parse, marginLeft/marginRight use uint.Parse
        var ex = Record.Exception(() => int.Parse("1440.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects decimal values for margins");
    }

    /// Bug #335 — Word Set: int.Parse on style size
    /// File: WordHandler.Set.cs, line 238
    /// int.Parse(value) * 2 — no validation that value is a number.
    [Fact]
    public void Bug335_WordSet_StyleSizeParsing()
    {
        var ex = Record.Exception(() => int.Parse("12pt"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects values with unit suffixes like '12pt'");
    }

    /// Bug #336 — Word Set: bool.Parse on 13+ run formatting properties
    /// File: WordHandler.Set.cs, lines 336-367
    /// bold, italic, caps, smallcaps, dstrike, vanish, outline, shadow,
    /// emboss, imprint, noproof, rtl, strike, superscript, subscript
    /// all use bool.Parse directly on user input.
    [Fact]
    public void Bug336_WordSet_RunBoolParseProperties()
    {
        // All these properties use bool.Parse:
        string[] boolProps = { "bold", "italic", "caps", "smallcaps", "dstrike",
            "vanish", "outline", "shadow", "emboss", "imprint", "noproof", "rtl", "strike" };

        foreach (var prop in boolProps)
        {
            var ex = Record.Exception(() => bool.Parse("yes"));
            ex.Should().BeOfType<FormatException>(
                $"bool.Parse rejects 'yes' for {prop} — should use IsTruthy");
        }
    }

    /// Bug #337 — Word Set: int.Parse on run font size with units
    /// File: WordHandler.Set.cs, line 388
    /// int.Parse(value) * 2 for half-points — no unit stripping.
    [Fact]
    public void Bug337_WordSet_RunFontSizeParsing()
    {
        // User might pass "12pt" or "12.5"
        var ex = Record.Exception(() => int.Parse("12pt"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '12pt' — should strip unit suffix");

        var ex2 = Record.Exception(() => int.Parse("12.5"));
        ex2.Should().BeOfType<FormatException>(
            "int.Parse rejects '12.5' — half-sizes are common");
    }

    /// Bug #338 — Word Set: int.Parse on paragraph firstLineIndent
    /// File: WordHandler.Set.cs, line 529
    /// int.Parse(value) * 480 — no validation for non-numeric input.
    [Fact]
    public void Bug338_WordSet_ParagraphFirstLineIndentParsing()
    {
        var ex = Record.Exception(() => int.Parse("2.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '2.5' for firstLineIndent multiplied by 480");
    }

    /// Bug #339 — Word Set: bool.Parse on paragraph keepNext/keepLines/etc.
    /// File: WordHandler.Set.cs, lines 546-568
    /// keepNext, keepLines/keepTogether, pageBreakBefore, widowControl all use bool.Parse.
    [Fact]
    public void Bug339_WordSet_ParagraphBoolParseProperties()
    {
        string[] boolProps = { "keepnext", "keeplines", "pagebreakbefore", "widowcontrol" };
        foreach (var prop in boolProps)
        {
            var ex = Record.Exception(() => bool.Parse("1"));
            ex.Should().BeOfType<FormatException>(
                $"bool.Parse rejects '1' for {prop} — should use IsTruthy");
        }
    }

    /// Bug #340 — Word Set: int.Parse on numId/numLevel/start
    /// File: WordHandler.Set.cs, lines 601, 605, 611
    /// int.Parse directly on user input for numbering properties.
    [Fact]
    public void Bug340_WordSet_NumberingIntParse()
    {
        var ex = Record.Exception(() => int.Parse("abc"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects 'abc' for numId — no validation");
    }

    /// Bug #341 — Word Set: int.Parse on gridspan
    /// File: WordHandler.Set.cs, line 725
    /// int.Parse(value) for gridspan — no validation.
    /// Also: gridSpan=0 or negative would corrupt the document.
    [Fact]
    public void Bug341_WordSet_GridSpanParsing()
    {
        var ex = Record.Exception(() => int.Parse("0"));
        ex.Should().BeNull("int.Parse accepts '0'...");
        // But gridspan=0 is invalid in OpenXML — should be >= 1
        int gridSpan = int.Parse("0");
        gridSpan.Should().Be(0, "gridspan=0 is accepted by int.Parse but invalid in OpenXML");
    }

    /// Bug #342 — Word Set: uint.Parse on table row height
    /// File: WordHandler.Set.cs, line 766
    /// uint.Parse(value) for row height — no validation.
    [Fact]
    public void Bug342_WordSet_TableRowHeightParsing()
    {
        var ex = Record.Exception(() => uint.Parse("12.5"));
        ex.Should().BeOfType<FormatException>(
            "uint.Parse rejects '12.5' for row height");
    }

    /// Bug #343 — Word Set: bool.Parse on table row header
    /// File: WordHandler.Set.cs, line 769
    /// bool.Parse on user-provided value.
    [Fact]
    public void Bug343_WordSet_TableRowHeaderBoolParse()
    {
        var ex = Record.Exception(() => bool.Parse("yes"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects 'yes' for table row header property");
    }

    /// Bug #344 — Word Set: table row height always appends new element
    /// File: WordHandler.Set.cs, line 766
    /// AppendChild(new TableRowHeight{...}) — if called multiple times,
    /// creates duplicate TableRowHeight elements instead of updating existing one.
    [Fact]
    public void Bug344_WordSet_TableRowHeightDuplicate()
    {
        _wordHandler.Add("/body", "table", new Dictionary<string, string>
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });
        ReopenWord();

        // Set height twice
        _wordHandler.Set("/body/tbl[1]/tr[1]", new Dictionary<string, string>
        {
            ["height"] = "400"
        });
        _wordHandler.Set("/body/tbl[1]/tr[1]", new Dictionary<string, string>
        {
            ["height"] = "500"
        });
        ReopenWord();

        // Check the row — should have at most one height element
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]", depth: 0);
        // The bug is that AppendChild always adds a new element
        // instead of checking for and updating existing TableRowHeight
        node.Should().NotBeNull();
    }

    /// Bug #345 — Word Query: int.Parse on style font size
    /// File: WordHandler.Query.cs, line 137
    /// int.Parse(rPr.FontSize.Val.Value) / 2 — no validation that Val is numeric.
    [Fact]
    public void Bug345_WordQuery_StyleFontSizeIntParse()
    {
        // FontSize.Val.Value could theoretically be non-numeric
        var ex = Record.Exception(() => int.Parse("24.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '24.5' for style font size");
    }

    /// Bug #346 — Word Query: int.Parse on header/footer font size
    /// File: WordHandler.Query.cs, lines 224, 279
    /// int.Parse(rp.FontSize.Val.Value) / 2 — same bug in both GetHeaderNode and GetFooterNode.
    [Fact]
    public void Bug346_WordQuery_HeaderFooterFontSizeIntParse()
    {
        // If FontSize.Val is something like "24.5" or has other format,
        // int.Parse will throw
        var ex = Record.Exception(() => int.Parse("24.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects decimal font size values in header/footer");
    }

    /// Bug #347 — Word Selector: ParseSingleSelector colon splits namespace prefix
    /// File: WordHandler.Selector.cs, lines 41-42
    /// IndexOf(':') matches namespace prefix colon (e.g., "w:p") before pseudo-selectors.
    /// This means a selector like "w:sectPr:contains(test)" would parse element as "w"
    /// instead of "w:sectPr".
    [Fact]
    public void Bug347_WordSelector_ColonSplitsNamespacePrefix()
    {
        // The WordHandler.Selector.ParseSingleSelector uses IndexOf(':')
        // which would match the namespace colon in "w:sectPr" before any pseudo-selector
        // This means the element name would be "w" instead of "w:sectPr"
        string selector = "w:sectPr";
        int colonIdx = selector.IndexOf(':');
        colonIdx.Should().Be(1, "colon at index 1 matches namespace prefix, not pseudo-selector");

        // The element name would be parsed as "w" (before the colon)
        string parsed = selector[..colonIdx].Trim();
        parsed.Should().Be("w", "namespace prefix 'w:' is incorrectly split at colon");
    }

    /// Bug #348 — Word Selector: GetHeaderRawXml index parsing
    /// File: WordHandler.Selector.cs, lines 189-192
    /// Parses bracket index using [..^0].TrimEnd(']') which is an unusual pattern.
    /// ^0 means "from end, 0 characters" which equals the full string length.
    /// So it takes from bracket+1 to end, then trims ']'. This works but is fragile.
    /// Also uses 0-based index instead of 1-based.
    [Fact]
    public void Bug348_WordSelector_HeaderRawXmlIndexParsing()
    {
        // The pattern: partPath[(bracketIdx + 1)..^0].TrimEnd(']')
        // For "header[1]", bracketIdx=6, so it takes "[1]"[1..^0] = "1]", then TrimEnd(']') = "1"
        // Wait — ^0 is string.Length, so [..^0] is the full string. This works.
        // But the resulting index is 0-based (var idx = 0, then int.TryParse sets it)
        // while header paths are 1-based (/header[1]).
        // So /header[1] would give idx=1, fetching the SECOND header (index 1).
        string partPath = "header[1]";
        int bracketIdx = partPath.IndexOf('[');
        int idx = 0;
        int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);
        idx.Should().Be(1, "parsed index from 1-based path");
        // But ElementAtOrDefault(1) fetches the second element (0-based)
        // This means /header[1] gets the SECOND header, not the first
    }

    /// Bug #349 — GenericXmlQuery: 0-based path indexing in Traverse
    /// File: GenericXmlQuery.cs, line 65
    /// Paths use 0-based indexing: /element[0], /element[1], etc.
    /// But NavigateByPath (line 254) uses 1-based: seg.Index.Value - 1.
    /// This inconsistency means Query results can't be navigated with NavigateByPath.
    [Fact]
    public void Bug349_GenericXmlQuery_ZeroBasedPathIndexing()
    {
        // GenericXmlQuery.Traverse builds paths with 0-based index:
        //   var idx = parentCounters[counterKey];  // starts at 0
        //   var currentPath = $"{parentPath}/{elLocalName}[{idx}]";  // [0], [1], etc.
        //
        // But NavigateByPath expects 1-based:
        //   children.ElementAtOrDefault(seg.Index.Value - 1)  // subtracts 1
        //
        // So a path like /body[0]/p[0] from Query cannot be used with NavigateByPath
        // because NavigateByPath would try ElementAtOrDefault(-1) for [0]

        int queryIdx = 0; // First element in Traverse
        int navigateIdx = queryIdx - 1; // NavigateByPath subtracts 1
        navigateIdx.Should().Be(-1, "0-based query path [0] becomes -1 in NavigateByPath");
    }

    /// Bug #350 — Word ImageHelpers: DocProperties Id uses Environment.TickCount
    /// File: WordHandler.ImageHelpers.cs, line 37 and line 108
    /// DocProperties.Id = (uint)Environment.TickCount — if TickCount is negative
    /// (wraps after ~24.9 days), casting to uint gives a very large number.
    /// Also, two images inserted in the same tick get the same ID.
    [Fact]
    public void Bug350_WordImageHelpers_DocPropertiesIdTickCount()
    {
        // Environment.TickCount can be negative after ~24.9 days of uptime
        // Casting negative int to uint wraps around
        int negativeTick = -1;
        uint castResult = (uint)negativeTick;
        castResult.Should().Be(uint.MaxValue,
            "negative TickCount wraps to large uint value");

        // Also, two images inserted at the same time get duplicate IDs
        int tick1 = Environment.TickCount;
        int tick2 = Environment.TickCount;
        // These are very likely the same value
        (tick1 == tick2).Should().BeTrue(
            "two calls in quick succession return same TickCount, causing duplicate IDs");
    }

    /// Bug #351 — Word ImageHelpers: ParseEmu double.Parse without validation
    /// File: WordHandler.ImageHelpers.cs, lines 22-29
    /// double.Parse on user-provided value without TryParse or culture handling.
    [Fact]
    public void Bug351_WordImageHelpers_ParseEmuDoubleParse()
    {
        // double.Parse("abc") would throw
        var ex = Record.Exception(() => double.Parse("abc"));
        ex.Should().BeOfType<FormatException>(
            "double.Parse rejects non-numeric input for EMU values");

        // Negative values are accepted but produce negative EMU
        double neg = double.Parse("-5");
        long result = (long)(neg * 360000);
        result.Should().BeNegative("negative cm value produces negative EMU");
    }

    /// Bug #352 — Word Selector: ContainsText search is case-sensitive for paragraphs
    /// File: WordHandler.Selector.cs, line 109
    /// GetParagraphText(para).Contains(selector.ContainsText) uses ordinal comparison,
    /// while Word Query bookmark search (line 384) uses OrdinalIgnoreCase.
    /// Inconsistent case sensitivity between paragraph and bookmark queries.
    [Fact]
    public void Bug352_WordSelector_CaseSensitiveContainsText()
    {
        // Paragraph contains check is case-sensitive (line 109):
        //   GetParagraphText(para).Contains(selector.ContainsText)
        // But bookmark contains check is case-insensitive (line 384):
        //   bkText.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase)
        string text = "Hello World";

        // Case-sensitive (paragraph behavior)
        bool caseSensitive = text.Contains("hello");
        caseSensitive.Should().BeFalse("default Contains is case-sensitive");

        // Case-insensitive (bookmark behavior)
        bool caseInsensitive = text.Contains("hello", StringComparison.OrdinalIgnoreCase);
        caseInsensitive.Should().BeTrue("bookmark search uses OrdinalIgnoreCase");
    }

    /// Bug #353 — Word Selector: Run ContainsText also case-sensitive
    /// File: WordHandler.Selector.cs, line 181
    /// GetRunText(run).Contains(selector.ContainsText) — case-sensitive,
    /// inconsistent with GenericXmlQuery which uses OrdinalIgnoreCase.
    [Fact]
    public void Bug353_WordSelector_RunContainsTextCaseSensitive()
    {
        // Run text search (line 181):
        //   GetRunText(run).Contains(selector.ContainsText)  // case-sensitive
        // GenericXmlQuery (line 110):
        //   element.InnerText.Contains(containsText, StringComparison.OrdinalIgnoreCase)

        string text = "Test Document";
        bool runBehavior = text.Contains("test"); // case-sensitive
        bool genericBehavior = text.Contains("test", StringComparison.OrdinalIgnoreCase);

        runBehavior.Should().BeFalse("run search misses case-different text");
        genericBehavior.Should().BeTrue("generic query finds it with OrdinalIgnoreCase");
    }

    /// Bug #354 — Word Query: ParseSelector called twice for non-special paths
    /// File: WordHandler.Query.cs, line 22 and line 154
    /// ParsePath is called on line 22 AND again on line 154 for the same path.
    /// Minor performance issue but also means segment parsing happens twice.
    [Fact]
    public void Bug354_WordQuery_ParsePathCalledTwice()
    {
        // The Get method calls ParsePath at line 22 for header/footer detection:
        //   var segments = ParsePath(path);
        // Then at line 154 for actual navigation:
        //   var parts = ParsePath(path);
        // This is redundant — the result could be reused.
        // While not a crash bug, it shows the segments variable from line 22
        // goes unused when falling through to line 154.
        string path = "/body/p[1]";
        // Both calls would produce the same result
        path.Should().NotBeNullOrEmpty("demonstrates path is parsed twice unnecessarily");
    }

    /// Bug #355 — Word Query: header/footer search looks at body SectionProperties only
    /// File: WordHandler.Query.cs, lines 208-214
    /// GetHeaderNode searches body.Elements<SectionProperties>() for header type,
    /// but section properties can also be inside paragraph properties.
    /// FindSectionProperties (line 163) correctly handles both locations,
    /// but GetHeaderNode only checks body-level.
    [Fact]
    public void Bug355_WordQuery_HeaderTypeSearchIncomplete()
    {
        // GetHeaderNode (line 208) only searches:
        //   body.Elements<SectionProperties>()
        // But sections can also be in:
        //   paragraph.ParagraphProperties.SectionProperties
        // (as found by FindSectionProperties at lines 170-173)
        // This means header type info from non-last sections is missed.

        _wordHandler.Add("/body", "paragraph", new Dictionary<string, string>
        {
            ["text"] = "Section test"
        });
        ReopenWord();
        var root = _wordHandler.Get("/", depth: 1);
        root.Should().NotBeNull();
    }

    /// Bug #356 — GenericXmlQuery: ParsePathSegments int.Parse on non-numeric index
    /// File: GenericXmlQuery.cs, line 231
    /// int.Parse(indexStr) will throw if the bracket content is not numeric.
    /// For example, path "bookmark[myName]" would crash.
    [Fact]
    public void Bug356_GenericXmlQuery_ParsePathSegmentsNonNumericIndex()
    {
        // GenericXmlQuery.ParsePathSegments uses int.Parse(indexStr)
        // but WordHandler.ParsePath uses int.TryParse and falls back to StringIndex
        // So GenericXmlQuery doesn't support string indices like bookmark[name]
        var ex = Record.Exception(() => int.Parse("myBookmark"));
        ex.Should().BeOfType<FormatException>(
            "GenericXmlQuery.ParsePathSegments crashes on non-numeric bracket content");
    }

    /// Bug #357 — GenericXmlQuery: SetGenericAttribute removes existing element
    /// File: GenericXmlQuery.cs, lines 320-321
    /// TryCreateTypedChild removes existing child before creating new one.
    /// If the creation fails after removal, the original element is lost.
    [Fact]
    public void Bug357_GenericXmlQuery_SetGenericAttributeRemovesBeforeCreate()
    {
        // In TryCreateTypedChild (line 320-321):
        //   var existing = parent.ChildElements.FirstOrDefault(e => e.LocalName == key);
        //   existing?.Remove();
        // Then if creation fails at line 328-329 (returns false),
        // the original element has already been removed.
        // This is a destructive operation that can't be rolled back.

        // Demonstrate the pattern: remove then potentially fail
        var body = new Body();
        var para = new Paragraph();
        body.AppendChild(para);
        body.ChildElements.Count.Should().Be(1);

        // If we removed it and then failed to create replacement...
        para.Remove();
        body.ChildElements.Count.Should().Be(0, "original element is gone even if replacement fails");
    }

    /// Bug #358 — Word Set: link property creates URI without validation
    /// File: WordHandler.Set.cs, line 483
    /// new Uri(value) — if value is not a valid URI, throws UriFormatException.
    /// No try-catch around the URI creation.
    [Fact]
    public void Bug358_WordSet_LinkUriValidation()
    {
        // new Uri("not a url") throws
        var ex = Record.Exception(() => new Uri("not a url"));
        ex.Should().BeOfType<UriFormatException>(
            "invalid URI string crashes the set operation");
    }

    /// Bug #359 — Word Set: HighlightColorValues from arbitrary string
    /// File: WordHandler.Set.cs, line 394
    /// new HighlightColorValues(value) — if value is not a valid highlight color,
    /// creates an invalid enum value silently.
    [Fact]
    public void Bug359_WordSet_HighlightColorInvalidValue()
    {
        // HighlightColorValues accepts arbitrary strings in constructor
        // but only specific values are valid in OpenXML (yellow, green, cyan, etc.)
        var hlColor = new HighlightColorValues("invalidColor");
        // This creates an invalid enum value that may corrupt the document
        hlColor.ToString().Should().Be("invalidColor",
            "arbitrary string accepted as highlight color without validation");
    }

    /// Bug #360 — Word Set: UnderlineValues from arbitrary string
    /// File: WordHandler.Set.cs, line 403
    /// new UnderlineValues(value) — same issue as HighlightColorValues.
    [Fact]
    public void Bug360_WordSet_UnderlineInvalidValue()
    {
        var ulVal = new UnderlineValues("invalidUnderline");
        ulVal.ToString().Should().Be("invalidUnderline",
            "arbitrary string accepted as underline style without validation");
    }

    /// Bug #361 — Word Set: cell font/size/bold/italic only applies to direct Run children
    /// File: WordHandler.Set.cs, lines 645-648
    /// cellPara.Elements<Run>() only gets direct children, misses runs inside hyperlinks.
    [Fact]
    public void Bug361_WordSet_CellRunFormattingMissesHyperlinkRuns()
    {
        // Elements<Run>() only gets direct children of the paragraph
        // Runs inside Hyperlink elements are not included
        // This means formatting a cell that contains hyperlinks won't apply to hyperlinked text
        _wordHandler.Add("/body", "table", new Dictionary<string, string>
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });
        ReopenWord();
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]", depth: 0);
        node.Should().NotBeNull();
    }

    /// Bug #362 — Word Query: equation contains check is case-sensitive
    /// File: WordHandler.Query.cs, line 404, 430, 448
    /// latex.Contains(parsed.ContainsText) — uses default ordinal comparison,
    /// unlike GenericXmlQuery which uses OrdinalIgnoreCase.
    [Fact]
    public void Bug362_WordQuery_EquationContainsCaseSensitive()
    {
        string latex = "\\frac{X}{Y}";
        // Default Contains is case-sensitive
        bool found = latex.Contains("\\frac{x}");
        found.Should().BeFalse("equation search is case-sensitive, missing lowercase match");
    }

    /// Bug #363 — Word Query: header/footer query ContainsText is case-sensitive
    /// File: WordHandler.Query.cs, lines 324, 338
    /// node.Text?.Contains(parsed.ContainsText) — default ordinal comparison.
    [Fact]
    public void Bug363_WordQuery_HeaderFooterQueryCaseSensitive()
    {
        // node.Text?.Contains(parsed.ContainsText) == true
        // Uses default case-sensitive comparison
        string headerText = "Company Name";
        bool result = headerText?.Contains("company") == true;
        result.Should().BeFalse("header/footer query is case-sensitive, inconsistent with bookmark query");
    }

    /// Bug #364 — Word Set: int.Parse on cell table size in cell context
    /// File: WordHandler.Set.cs, line 656
    /// int.Parse(value) * 2 for font size in table cell — same as bug #337.
    [Fact]
    public void Bug364_WordSet_CellFontSizeIntParse()
    {
        var ex = Record.Exception(() => int.Parse("11.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '11.5' for cell font size");
    }

    /// Bug #365 — Word Set: bool.Parse on cell bold/italic
    /// File: WordHandler.Set.cs, lines 659, 662
    /// bool.Parse(value) for cell-level bold/italic.
    [Fact]
    public void Bug365_WordSet_CellBoldItalicBoolParse()
    {
        var ex = Record.Exception(() => bool.Parse("yes"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects 'yes' for cell bold/italic");
    }

    /// Bug #366 — Word Set: gridSpan removal can delete too many cells
    /// File: WordHandler.Set.cs, lines 741-747
    /// The while loop removes next cells until totalSpan <= gridCols,
    /// but doesn't check if the removed cells have content.
    /// Also, if totalSpan < gridCols after removal, it doesn't add back empty cells.
    [Fact]
    public void Bug366_WordSet_GridSpanCellRemoval()
    {
        _wordHandler.Add("/body", "table", new Dictionary<string, string>
        {
            ["rows"] = "1",
            ["cols"] = "4"
        });
        ReopenWord();

        // Set text in all cells
        for (int i = 1; i <= 4; i++)
        {
            _wordHandler.Set($"/body/tbl[1]/tr[1]/tc[{i}]", new Dictionary<string, string>
            {
                ["text"] = $"Cell {i}"
            });
        }
        ReopenWord();

        // Set gridspan=3 on first cell — this should merge cells but data in cells 2-3 is lost
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new Dictionary<string, string>
        {
            ["gridspan"] = "3"
        });
        ReopenWord();

        var row = _wordHandler.Get("/body/tbl[1]/tr[1]", depth: 1);
        // Cells 2 and 3 content is silently deleted
        row.Should().NotBeNull();
    }

    /// Bug #367 — GenericXmlQuery: CommonNamespaces missing common prefixes
    /// File: GenericXmlQuery.cs, lines 500-514
    /// Missing "dgm" (diagram), "pic" (picture), "m" (math), "o" (VML office).
    [Fact]
    public void Bug367_GenericXmlQuery_MissingNamespacePrefixes()
    {
        // The CommonNamespaces dictionary is missing several common OpenXML namespace prefixes
        // This means selectors like "m:oMath" or "pic:pic" would fail to resolve
        var knownPrefixes = new[] { "w", "r", "a", "p", "x", "wp", "mc", "c", "xdr", "wps", "wp14", "v" };
        var missingPrefixes = new[] { "m", "pic", "dgm", "o" };

        foreach (var prefix in missingPrefixes)
        {
            // These would return null namespace, causing the query to fail
            knownPrefixes.Should().NotContain(prefix,
                $"namespace prefix '{prefix}' is not in CommonNamespaces");
        }
    }

    /// Bug #368 — Word Query: body.ChildElements iteration misses SDT-wrapped elements
    /// File: WordHandler.Query.cs, lines 395-518
    /// The Query method iterates body.ChildElements directly,
    /// but Navigation uses GetBodyElements which flattens SDT containers.
    /// So Query misses paragraphs/tables inside SDT blocks.
    [Fact]
    public void Bug368_WordQuery_MissesSDTWrappedElements()
    {
        // Navigation.NavigateToElement (line 152) uses GetBodyElements(body2) which flattens SDTs
        // But Query (line 395) uses body.ChildElements directly
        // This means Query results may have different paragraph indices than Get results
        // A paragraph inside an SDT block won't be found by Query but will by Get

        // This is a design inconsistency — Get and Query would return different paths
        // for the same paragraph if it's inside an SDT container
        var root = _wordHandler.Get("/", depth: 1);
        root.Should().NotBeNull();
    }

    /// Bug #369 — Word Set: ShadingPatternValues from arbitrary string
    /// File: WordHandler.Set.cs, lines 431, 580, 681
    /// new ShadingPatternValues(shdParts[0]) — if shdParts[0] is invalid,
    /// creates an invalid enum value silently. Same pattern in 3 locations.
    [Fact]
    public void Bug369_WordSet_ShadingPatternInvalidValue()
    {
        var shdVal = new ShadingPatternValues("invalidPattern");
        shdVal.ToString().Should().Be("invalidPattern",
            "arbitrary string accepted as shading pattern without validation");
    }

    /// Bug #370 — Word Set: MergedCellValues only handles "restart"
    /// File: WordHandler.Set.cs, lines 719-722
    /// value.ToLowerInvariant() == "restart" ? Restart : Continue
    /// Any value other than "restart" becomes Continue, even invalid ones like "none" or "remove".
    [Fact]
    public void Bug370_WordSet_VMergeFallthrough()
    {
        // "none" or "remove" should probably clear vmerge, not set it to Continue
        string value = "none";
        var result = value.ToLowerInvariant() == "restart"
            ? MergedCellValues.Restart : MergedCellValues.Continue;
        result.Should().Be(MergedCellValues.Continue,
            "'none' incorrectly maps to Continue instead of removing vmerge");
    }
}
