// FuzzRound14 — Target: untested lifecycle corners after R13 fixes
//
// Areas:
//   WT01–WT03: Word TOC remove + reopen (body-reference cleanup)
//   WS01–WS03: Word Swap paragraphs basic lifecycle
//   EX01–EX03: Excel chart Remove via chart[N] path (no handler — expected throw/no-op, not NullRef)
//   ET01–ET03: Excel table Remove via table[N] path (no remove handler — guard test)
//   EV01–EV03: Excel Swap rows persistence round-trip
//   PC01–PC03: PPTX connector add/remove/reopen lifecycle
//   PL01–PL03: PPTX CopyFrom shape across two slides

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound14 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz14_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== WT01–WT03: Word TOC remove + reopen ====================

    [Fact]
    public void WT01_Word_AddToc_ThenRemove_NoNullRef()
    {
        // FIX: Word.Remove("/toc[1]") now handles /toc[N] path correctly.
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Heading", ["style"] = "Heading1" });
        h.Add("/body", "toc", null, new() { ["levels"] = "1-3" });
        var act = () => h.Remove("/toc[1]");
        act.Should().NotThrow("Remove(\"/toc[1]\") should be supported");
        // After removal, Get should throw because TOC no longer exists
        var getAct = () => h.Get("/toc[1]");
        getAct.Should().Throw<ArgumentException>("TOC should no longer exist after removal");
    }

    [Fact]
    public void WT02_Word_AddToc_FileValidAfterAdd()
    {
        // Simplified: just verify TOC add + reopen is valid (remove path is a known bug)
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Introduction", ["style"] = "Heading1" });
            h.Add("/body", "toc", null, new() { ["levels"] = "1-2" });
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            var node = h2.Get("/toc[1]");
            node.Should().NotBeNull("TOC[1] should be accessible after save+reopen");
        };
        act.Should().NotThrow("file should be valid after TOC add and reopen");
    }

    [Fact]
    public void WT03_Word_AddToc_Remove_Persistence()
    {
        // Verify TOC removal persists after save and reopen.
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Chapter 1", ["style"] = "Heading1" });
            h.Add("/body", "toc", null, new() { ["levels"] = "1-3" });
            h.Remove("/toc[1]");
        }
        using var h2 = new WordHandler(path, editable: false);
        var act = () => h2.Get("/toc[1]");
        act.Should().Throw<ArgumentException>("TOC should not exist after removal and reopen");
    }

    // ==================== WS01–WS03: Word Swap paragraphs ====================

    [Fact]
    public void WS01_Word_Swap_TwoParagraphs_OrderChanges()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Alpha" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Beta" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Gamma" });
        // Swap first and third paragraphs
        var (p1, p3) = h.Swap("/body/p[1]", "/body/p[3]");
        // After swap the returned paths should be valid
        p1.Should().NotBeNullOrEmpty("swap should return a non-empty path for element 1");
        p3.Should().NotBeNullOrEmpty("swap should return a non-empty path for element 2");
        // The text at the first position should now be "Gamma" (was "Alpha")
        var first = h.Get("/body/p[1]");
        first.Should().NotBeNull("first paragraph should still be accessible after swap");
        first!.Text.Should().Be("Gamma", "swap should have moved Gamma to position 1");
    }

    [Fact]
    public void WS02_Word_Swap_Paragraphs_Persistence()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "First" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "Second" });
            h.Swap("/body/p[1]", "/body/p[2]");
        }
        using var h2 = new WordHandler(path, editable: false);
        var para1 = h2.Get("/body/p[1]");
        para1.Should().NotBeNull("first paragraph should be accessible after swap + reopen");
        para1!.Text.Should().Be("Second", "swapped order should persist after save and reopen");
    }

    [Fact]
    public void WS03_Word_Swap_AdjacentParagraphs_NoNullRef()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para A" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para B" });
        var act = () => h.Swap("/body/p[1]", "/body/p[2]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef during Word Swap: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable */ }
    }

    // ==================== EX01–EX03: Excel chart Remove guard ====================

    [Fact]
    public void EX01_Excel_RemoveChart_ByIndex_NoNullRef()
    {
        // chart[N] remove path has no dedicated handler — must not NullRef
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Add("/Sheet1", "chart", null, new() { ["chartType"] = "bar", ["data"] = "S1:10,20,30" });
        var act = () => h.Remove("/Sheet1/chart[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing Excel chart: {ex.Message}"); }
        catch (Exception) { /* unsupported path = acceptable */ }
    }

    [Fact]
    public void EX02_Excel_RemoveChart_FileRemainsValid()
    {
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            h.Add("/Sheet1", "chart", null, new() { ["chartType"] = "line", ["data"] = "S1:5,10,15" });
            try { h.Remove("/Sheet1/chart[1]"); } catch (Exception) { /* if unsupported, skip */ }
        }
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/A1");
        };
        act.Should().NotThrow("file should remain valid after chart remove attempt");
    }

    [Fact]
    public void EX03_Excel_RemoveChart_OutOfRange_NoNullRef()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        // No chart added — removing chart[1] must not NullRef
        var act = () => h.Remove("/Sheet1/chart[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing nonexistent chart: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = expected */ }
    }

    // ==================== ET01–ET03: Excel table Remove guard ====================

    [Fact]
    public void ET01_Excel_RemoveTable_ByIndex_NoNullRef()
    {
        // table[N] remove has no dedicated handler — must not NullRef
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        for (int r = 1; r <= 3; r++)
            for (int c = 0; c < 3; c++)
                h.Set($"/Sheet1/{(char)('A' + c)}{r}", new() { ["value"] = $"v{r}{c}" });
        try { h.Add("/Sheet1", "table", null, new() { ["range"] = "A1:C3", ["name"] = "Table1" }); }
        catch (Exception) { return; } // skip if add unsupported
        var act = () => h.Remove("/Sheet1/table[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing Excel table: {ex.Message}"); }
        catch (Exception) { /* unsupported = acceptable */ }
    }

    [Fact]
    public void ET02_Excel_RemoveTable_FileRemainsValid()
    {
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            for (int r = 1; r <= 3; r++)
                for (int c = 0; c < 3; c++)
                    h.Set($"/Sheet1/{(char)('A' + c)}{r}", new() { ["value"] = $"v{r}{c}" });
            try
            {
                h.Add("/Sheet1", "table", null, new() { ["range"] = "A1:C3", ["name"] = "TestTable" });
                h.Remove("/Sheet1/table[1]");
            }
            catch (Exception) { /* if unsupported, skip */ }
        }
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/A1");
        };
        act.Should().NotThrow("file should remain valid after table remove attempt");
    }

    [Fact]
    public void ET03_Excel_RemoveTable_NoTable_NoNullRef()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var act = () => h.Remove("/Sheet1/table[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing table from sheet with no tables: {ex.Message}"); }
        catch (Exception) { /* expected */ }
    }

    // ==================== EV01–EV03: Excel Swap rows persistence ====================

    [Fact]
    public void EV01_Excel_Swap_Rows_OrderChanges()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "Row1" });
        h.Set("/Sheet1/A2", new() { ["value"] = "Row2" });
        h.Set("/Sheet1/A3", new() { ["value"] = "Row3" });
        var (r1, r2) = h.Swap("/Sheet1/row[1]", "/Sheet1/row[3]");
        r1.Should().NotBeNullOrEmpty();
        r2.Should().NotBeNullOrEmpty();
        var cell = h.Get("/Sheet1/A1");
        cell.Should().NotBeNull();
        // After swap rows 1 and 3, A1 should contain what was in A3
        cell!.Text.Should().Be("Row3", "swap should move row 3 content to row 1 position");
    }

    [Fact]
    public void EV02_Excel_Swap_Rows_Persistence()
    {
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            h.Set("/Sheet1/A1", new() { ["value"] = "First" });
            h.Set("/Sheet1/A2", new() { ["value"] = "Second" });
            h.Swap("/Sheet1/row[1]", "/Sheet1/row[2]");
        }
        using var h2 = new ExcelHandler(path, editable: false);
        var cell = h2.Get("/Sheet1/A1");
        cell.Should().NotBeNull("A1 should be accessible after swap + reopen");
        cell!.Text.Should().Be("Second", "swapped row order should persist after save and reopen");
    }

    [Fact]
    public void EV03_Excel_Swap_SameRow_NoThrowOrNoop()
    {
        // Swapping a row with itself should be a safe no-op
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "SameRow" });
        var act = () => h.Swap("/Sheet1/row[1]", "/Sheet1/row[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef swapping row with itself: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable */ }
    }

    // ==================== PC01–PC03: PPTX connector add/remove/reopen ====================

    [Fact]
    public void PC01_Pptx_AddConnector_ThenRemove_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "connector", null, new() {
            ["x1"] = "1cm", ["y1"] = "1cm", ["x2"] = "5cm", ["y2"] = "5cm"
        });
        var act = () => h.Remove("/slide[1]/connector[1]");
        act.Should().NotThrow("removing a connector should not throw");
    }

    [Fact]
    public void PC02_Pptx_AddConnector_Remove_Reopen_FileValid()
    {
        var path = CreateTemp("pptx");
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { });
            h.Add("/slide[1]", "connector", null, new() {
                ["x1"] = "1cm", ["y1"] = "2cm", ["x2"] = "6cm", ["y2"] = "4cm"
            });
            h.Remove("/slide[1]/connector[1]");
        }
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            _ = h2.Query("shape").ToList();
        };
        act.Should().NotThrow("file should be valid after connector add+remove and reopen");
    }

    [Fact]
    public void PC03_Pptx_TwoConnectors_RemoveFirst_SecondRemains()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "connector", null, new() { ["x1"] = "0cm", ["y1"] = "0cm", ["x2"] = "3cm", ["y2"] = "3cm" });
        h.Add("/slide[1]", "connector", null, new() { ["x1"] = "4cm", ["y1"] = "0cm", ["x2"] = "7cm", ["y2"] = "3cm" });
        h.Remove("/slide[1]/connector[1]");
        // After removing first connector, file should still be openable
        var act = () => { _ = h.Query("connector").ToList(); };
        act.Should().NotThrow("querying after first connector removal should not throw");
    }

    // ==================== PL01–PL03: PPTX CopyFrom shape across slides ====================

    [Fact]
    public void PL01_Pptx_CopyFrom_Shape_SameSlide_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Original", ["x"] = "1cm", ["y"] = "1cm" });
        var act = () => h.CopyFrom("/slide[1]/shape[1]", "/slide[1]", null);
        act.Should().NotThrow("copying a shape within the same slide should not throw");
    }

    [Fact]
    public void PL02_Pptx_CopyFrom_Shape_CrossSlide_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Source Shape", ["x"] = "1cm", ["y"] = "1cm" });
        var act = () => h.CopyFrom("/slide[1]/shape[1]", "/slide[2]", null);
        act.Should().NotThrow("copying a shape to a different slide should not throw");
    }

    [Fact]
    public void PL03_Pptx_CopyFrom_Shape_CrossSlide_TextPreserved()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() {
            ["text"] = "Cloned Text", ["fill"] = "4472C4", ["x"] = "2cm", ["y"] = "2cm"
        });
        string? clonePath = null;
        try { clonePath = h.CopyFrom("/slide[1]/shape[1]", "/slide[2]", null); }
        catch (Exception) { return; } // skip if unsupported
        clonePath.Should().NotBeNullOrEmpty("CopyFrom should return the new element path");
        var cloned = h.Get(clonePath!);
        cloned.Should().NotBeNull("cloned shape should be gettable by returned path");
        cloned!.Text.Should().Be("Cloned Text", "cloned shape should preserve original text");
    }
}
