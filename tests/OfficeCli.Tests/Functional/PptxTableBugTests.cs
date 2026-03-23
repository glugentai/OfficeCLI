// PptxTableBugTests.cs — Failing tests for table cell XML schema violations and related issues.
//
// Bug: Setting table row cell text via c1=/c2=/c3= shorthand produces invalid XML.
// Root cause: The c-shorthand handler removes Run/Break children from the paragraph but
// leaves any existing EndParagraphRunProperties in place, then AppendChild(newRun) puts
// the Run AFTER EndParagraphRunProperties. The DrawingML CT_TextParagraph schema requires
// EndParagraphRunProperties to be the very last child — any Run after it is a schema error.
//
// Additional issues found while reviewing table code:
//   - Table style Set: AppendChild(new TableStyleId) without first removing the old one
//     produces duplicate TableStyleId children, violating the TableProperties schema.
//   - Cell fill Set: fill appended with AppendChild without regard to TableCellProperties
//     child-element order (lnL/lnR/lnT/lnB/lnTlToBr/lnBlToTr must come before fills).

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

public class PptxTableBugTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTempPptx()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { System.IO.File.Delete(f); } catch { }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug: c1=/c2= shorthand puts Run AFTER EndParagraphRunProperties
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// When a table is created with blank cells (EndParagraphRunProperties only),
    /// setting text via c-shorthand on the row must place the Run BEFORE any
    /// EndParagraphRunProperties in the paragraph. This test verifies the schema
    /// order so PowerPoint does not silently discard text.
    /// </summary>
    [Fact]
    public void Set_CellShorthand_RunMustComeBeforeEndParagraphRunProperties()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "3" });

        // Blank cells created in AddTable have EndParagraphRunProperties (no Run).
        // Setting c1=Region via the row shorthand should produce a valid paragraph:
        //   <a:p> <a:r>...</a:r> <a:endParaRPr/> </a:p>
        // NOT:
        //   <a:p> <a:endParaRPr/> <a:r>...</a:r> </a:p>  <- schema violation
        handler.Set("/slide[1]/table[1]/tr[1]", new()
        {
            ["c1"] = "Region",
            ["c2"] = "Revenue",
            ["c3"] = "Growth",
        });

        // Reopen to force serialization round-trip
        handler.Dispose();
        using var h2 = new PowerPointHandler(path, editable: false);

        var row1 = h2.Get("/slide[1]/table[1]/tr[1]");
        row1.Should().NotBeNull();

        // Text must round-trip correctly
        var c1 = h2.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        var c2 = h2.Get("/slide[1]/table[1]/tr[1]/tc[2]");
        var c3 = h2.Get("/slide[1]/table[1]/tr[1]/tc[3]");
        c1.Text.Should().Be("Region",  "c1 text must survive serialization round-trip");
        c2.Text.Should().Be("Revenue", "c2 text must survive serialization round-trip");
        c3.Text.Should().Be("Growth",  "c3 text must survive serialization round-trip");
    }

    /// <summary>
    /// Directly inspects the raw XML paragraph structure inside a table cell after
    /// c-shorthand is used. The Run element must appear before EndParagraphRunProperties.
    /// </summary>
    [Fact]
    public void Set_CellShorthand_SchemaOrder_RunBeforeEndParagraphRunProperties()
    {
        var path = CreateTempPptx();
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });
            handler.Set("/slide[1]/table[1]/tr[1]", new() { ["c1"] = "Hello", ["c2"] = "World" });
        }

        // Inspect raw XML after close/save
        using var pkg = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false);
        var table = pkg.PresentationPart!.SlideParts.First()
                       .Slide.Descendants<Drawing.Table>().First();

        var row = table.Elements<Drawing.TableRow>().First();
        var cells = row.Elements<Drawing.TableCell>().ToList();

        foreach (var cell in cells)
        {
            var para = cell.TextBody!.Elements<Drawing.Paragraph>().First();
            var children = para.ChildElements.ToList();

            var runIdx     = children.FindIndex(c => c is Drawing.Run);
            var endParaIdx = children.FindIndex(c => c is Drawing.EndParagraphRunProperties);

            runIdx.Should().BeGreaterOrEqualTo(0,
                "paragraph must contain a Run after setting cell text via c-shorthand");
            endParaIdx.Should().BeGreaterOrEqualTo(0,
                "paragraph should still contain EndParagraphRunProperties");
            runIdx.Should().BeLessThan(endParaIdx,
                "Run must appear BEFORE EndParagraphRunProperties in CT_TextParagraph schema");
        }
    }

    /// <summary>
    /// Regression: setting text via c-shorthand multiple times on the same row
    /// must not accumulate EndParagraphRunProperties or produce duplicate Runs.
    /// </summary>
    [Fact]
    public void Set_CellShorthand_SecondSet_DoesNotDuplicateEndParaRPr()
    {
        var path = CreateTempPptx();
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

            // First set
            handler.Set("/slide[1]/table[1]/tr[1]", new() { ["c1"] = "First" });
            // Second set — should replace, not accumulate
            handler.Set("/slide[1]/table[1]/tr[1]", new() { ["c1"] = "Second" });
        }

        using var pkg = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false);
        var table = pkg.PresentationPart!.SlideParts.First()
                       .Slide.Descendants<Drawing.Table>().First();

        var para = table.Elements<Drawing.TableRow>().First()
                        .Elements<Drawing.TableCell>().First()
                        .TextBody!.Elements<Drawing.Paragraph>().First();

        var runCount     = para.Elements<Drawing.Run>().Count();
        var endParaCount = para.Elements<Drawing.EndParagraphRunProperties>().Count();

        runCount.Should().Be(1,
            "exactly one Run should exist after two successive c-shorthand sets");
        endParaCount.Should().BeLessThanOrEqualTo(1,
            "EndParagraphRunProperties must not be duplicated");

        var runIdx     = para.ChildElements.ToList().FindIndex(c => c is Drawing.Run);
        var endParaIdx = para.ChildElements.ToList().FindIndex(c => c is Drawing.EndParagraphRunProperties);
        if (endParaIdx >= 0)
            runIdx.Should().BeLessThan(endParaIdx,
                "Run must precede EndParagraphRunProperties even after multiple Set calls");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug: Table style Set duplicates TableStyleId on repeated calls
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Setting a table style twice should result in exactly one TableStyleId child
    /// inside TableProperties, not two.
    /// </summary>
    [Fact]
    public void Set_TableStyle_RepeatedSet_DoesNotDuplicateTableStyleId()
    {
        var path = CreateTempPptx();
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

            // Apply style twice
            handler.Set("/slide[1]/table[1]", new() { ["style"] = "medium1" });
            handler.Set("/slide[1]/table[1]", new() { ["style"] = "medium2" });
        }

        using var pkg = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false);
        var table = pkg.PresentationPart!.SlideParts.First()
                       .Slide.Descendants<Drawing.Table>().First();

        var tblPr = table.GetFirstChild<Drawing.TableProperties>()!;
        var styleIdCount = tblPr.Elements<Drawing.TableStyleId>().Count();

        styleIdCount.Should().Be(1,
            "TableProperties must contain exactly one TableStyleId after repeated Set calls");

        // The last-applied style should win
        tblPr.Elements<Drawing.TableStyleId>().First().Text.Should().Contain("{F5AB1C69",
            "the final style (medium2) GUID should be stored");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug: Cell fill AppendChild violates TableCellProperties child element order
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// The CT_TableCellProperties schema requires border line elements (lnL, lnR,
    /// lnT, lnB, lnTlToBr, lnBlToTr) to come BEFORE fill elements (solidFill,
    /// gradFill, blipFill, pattFill, noFill). AppendChild adds fill after whatever
    /// is already present, which can violate ordering when borders are set first.
    /// </summary>
    [Fact]
    public void Set_CellFillAfterBorder_SchemaOrder_BorderBeforeFill()
    {
        var path = CreateTempPptx();
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

            // Set border first, then fill — the schema requires border lines before fill
            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            {
                ["border.left"] = "2pt solid FF0000",
            });
            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            {
                ["fill"] = "0000FF",
            });
        }

        using var pkg = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false);
        var table = pkg.PresentationPart!.SlideParts.First()
                       .Slide.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.TableCellProperties!;

        var children = tcPr.ChildElements.ToList();
        var borderIdx = children.FindIndex(c =>
            c is Drawing.LeftBorderLineProperties
            or Drawing.RightBorderLineProperties
            or Drawing.TopBorderLineProperties
            or Drawing.BottomBorderLineProperties
            or Drawing.TopLeftToBottomRightBorderLineProperties
            or Drawing.BottomLeftToTopRightBorderLineProperties);
        var fillIdx = children.FindIndex(c =>
            c is Drawing.SolidFill
            or Drawing.GradientFill
            or Drawing.NoFill
            or Drawing.BlipFill);

        borderIdx.Should().BeGreaterOrEqualTo(0, "border line element should be present");
        fillIdx.Should().BeGreaterOrEqualTo(0,   "fill element should be present");
        borderIdx.Should().BeLessThan(fillIdx,
            "border line elements must appear before fill in CT_TableCellProperties schema");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Merge / colspan — basic correctness after Set
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Setting gridSpan (colspan) on a cell via Set should be readable back via Get
    /// and round-trip through file close/reopen.
    /// </summary>
    [Fact]
    public void Set_CellColspan_PersistsAfterReopen()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "3" });

        // Merge first row: cell 1 spans 2 columns
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"]    = "Merged",
            ["colspan"] = "2",
        });
        handler.Set("/slide[1]/table[1]/tr[1]/tc[2]", new()
        {
            ["hmerge"] = "true",
        });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("gridSpan",
            "Get should return gridSpan after Set colspan");
        node.Format["gridSpan"]!.ToString().Should().Be("2",
            "gridSpan value should match the set colspan");

        // Reopen and verify persistence
        handler.Dispose();
        using var h2 = new PowerPointHandler(path, editable: false);
        var node2 = h2.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node2.Format.Should().ContainKey("gridSpan",
            "gridSpan must persist after file close/reopen");
        node2.Format["gridSpan"]!.ToString().Should().Be("2");
    }

    /// <summary>
    /// Setting rowSpan on a cell via Set should be readable back via Get
    /// and round-trip through file close/reopen.
    /// </summary>
    [Fact]
    public void Set_CellRowspan_PersistsAfterReopen()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });

        // Merge first column: cell [1][1] spans 2 rows
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"]    = "Tall",
            ["rowspan"] = "2",
        });
        handler.Set("/slide[1]/table[1]/tr[2]/tc[1]", new()
        {
            ["vmerge"] = "true",
        });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("rowSpan",
            "Get should return rowSpan after Set rowspan");
        node.Format["rowSpan"]!.ToString().Should().Be("2");

        handler.Dispose();
        using var h2 = new PowerPointHandler(path, editable: false);
        var node2 = h2.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node2.Format.Should().ContainKey("rowSpan",
            "rowSpan must persist after file close/reopen");
        node2.Format["rowSpan"]!.ToString().Should().Be("2");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Border handling — basic round-trip
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Setting a cell border and then reading it back should return the border color.
    /// </summary>
    [Fact]
    public void Set_CellBorder_RoundTrips_Color()
    {
        var path = CreateTempPptx();
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"]         = "Bordered",
            ["border.left"]  = "2pt solid FF0000",
            ["border.right"] = "1pt solid 0000FF",
        });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Bordered");
        node.Format.Should().ContainKey("border.left",
            "Get should return border.left after Set");
    }

    /// <summary>
    /// Setting border then clearing it with "none" should remove the solid fill
    /// from the line properties element.
    /// </summary>
    [Fact]
    public void Set_CellBorder_None_RemovesBorderFill()
    {
        var path = CreateTempPptx();
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["border.top"] = "2pt solid FF0000" });
            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["border.top"] = "none" });
        }

        using var pkg = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false);
        var table = pkg.PresentationPart!.SlideParts.First()
                       .Slide.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First().Elements<Drawing.TableCell>().First();
        var tcPr = cell.TableCellProperties!;

        // After clearing, the top border line element should contain NoFill, not SolidFill
        var topBorder = tcPr.GetFirstChild<Drawing.TopBorderLineProperties>();
        topBorder.Should().NotBeNull("top border line element should still be present");
        topBorder!.Elements<Drawing.NoFill>().Should().ContainSingle(
            "top border should contain NoFill after being set to 'none'");
        topBorder.Elements<Drawing.SolidFill>().Should().BeEmpty(
            "SolidFill must be removed when border is set to 'none'");
    }
}
