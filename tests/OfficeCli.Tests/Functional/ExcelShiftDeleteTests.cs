// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for true-shift row/column deletion.
/// Deleting a row or column must renumber all subsequent rows/columns
/// so that there are no gaps — matching Excel's interactive behavior.
/// </summary>
public class ExcelShiftDeleteTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelShiftDeleteTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
        Seed3x3();
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    /// <summary>Fill a 3×3 grid: A1=r1c1, B2=r2c2, etc.</summary>
    private void Seed3x3()
    {
        for (int r = 1; r <= 3; r++)
            for (int c = 1; c <= 3; c++)
            {
                var col = (char)('A' + c - 1);
                _handler.Set($"/Sheet1/{col}{r}", new() { ["value"] = $"r{r}c{c}" });
            }
    }

    // ==================== Row delete ====================

    [Fact]
    public void RemoveRow_MiddleRow_ShiftsSubsequentRowsUp()
    {
        _handler.Remove("/Sheet1/row[2]");

        // row[1] unchanged
        _handler.Get("/Sheet1/A1").Text.Should().Be("r1c1");
        _handler.Get("/Sheet1/B1").Text.Should().Be("r1c2");

        // old row[3] is now row[2]
        _handler.Get("/Sheet1/A2").Text.Should().Be("r3c1");
        _handler.Get("/Sheet1/B2").Text.Should().Be("r3c2");
        _handler.Get("/Sheet1/C2").Text.Should().Be("r3c3");

        // no row[3] anymore
        _handler.Get("/Sheet1/A3").Text.Should().Be("(empty)");
    }

    [Fact]
    public void RemoveRow_FirstRow_ShiftsAll()
    {
        _handler.Remove("/Sheet1/row[1]");

        _handler.Get("/Sheet1/A1").Text.Should().Be("r2c1");
        _handler.Get("/Sheet1/A2").Text.Should().Be("r3c1");
        _handler.Get("/Sheet1/A3").Text.Should().Be("(empty)");
    }

    [Fact]
    public void RemoveRow_NativePath_Works()
    {
        _handler.Remove("Sheet1!row[2]");

        _handler.Get("Sheet1!A2").Text.Should().Be("r3c1");
    }

    [Fact]
    public void RemoveRow_PersistsAfterReopen()
    {
        _handler.Remove("/Sheet1/row[2]");
        Reopen();

        _handler.Get("/Sheet1/A1").Text.Should().Be("r1c1");
        _handler.Get("/Sheet1/A2").Text.Should().Be("r3c1");
        _handler.Get("/Sheet1/A3").Text.Should().Be("(empty)");
    }

    [Fact]
    public void RemoveRow_SetOnShiftedRow_WritesCorrectly()
    {
        _handler.Remove("/Sheet1/row[2]");

        // After deletion, row[2] is the old row[3]. Writing to row[2] should work.
        _handler.Set("/Sheet1/A2", new() { ["value"] = "updated" });
        _handler.Get("/Sheet1/A2").Text.Should().Be("updated");
    }

    // ==================== Column delete ====================

    [Fact]
    public void RemoveCol_MiddleCol_ShiftsSubsequentColsLeft()
    {
        _handler.Remove("/Sheet1/col[B]");

        // col A unchanged
        _handler.Get("/Sheet1/A1").Text.Should().Be("r1c1");
        _handler.Get("/Sheet1/A2").Text.Should().Be("r2c1");

        // old col C is now col B
        _handler.Get("/Sheet1/B1").Text.Should().Be("r1c3");
        _handler.Get("/Sheet1/B2").Text.Should().Be("r2c3");
        _handler.Get("/Sheet1/B3").Text.Should().Be("r3c3");

        // no col C anymore
        _handler.Get("/Sheet1/C1").Text.Should().Be("(empty)");
    }

    [Fact]
    public void RemoveCol_FirstCol_ShiftsAll()
    {
        _handler.Remove("/Sheet1/col[A]");

        _handler.Get("/Sheet1/A1").Text.Should().Be("r1c2");
        _handler.Get("/Sheet1/B1").Text.Should().Be("r1c3");
        _handler.Get("/Sheet1/C1").Text.Should().Be("(empty)");
    }

    [Fact]
    public void RemoveCol_NativePath_Works()
    {
        _handler.Remove("Sheet1!col[B]");

        _handler.Get("Sheet1!B1").Text.Should().Be("r1c3");
    }

    [Fact]
    public void RemoveCol_PersistsAfterReopen()
    {
        _handler.Remove("/Sheet1/col[B]");
        Reopen();

        _handler.Get("/Sheet1/A1").Text.Should().Be("r1c1");
        _handler.Get("/Sheet1/B1").Text.Should().Be("r1c3");
        _handler.Get("/Sheet1/C1").Text.Should().Be("(empty)");
    }

    [Fact]
    public void RemoveCol_SetOnShiftedCol_WritesCorrectly()
    {
        _handler.Remove("/Sheet1/col[B]");

        // After deletion, col B is old col C. Writing to B1 should work.
        _handler.Set("/Sheet1/B1", new() { ["value"] = "updated" });
        _handler.Get("/Sheet1/B1").Text.Should().Be("updated");
    }

    // ==================== Combined ====================

    [Fact]
    public void RemoveRowThenCol_BothShiftCorrectly()
    {
        _handler.Remove("/Sheet1/row[2]");
        _handler.Remove("/Sheet1/col[B]");

        // After row[2] delete: row[1]=r1cx, row[2]=r3cx
        // After col[B] delete: col A=c1, col B=c3
        _handler.Get("/Sheet1/A1").Text.Should().Be("r1c1");
        _handler.Get("/Sheet1/B1").Text.Should().Be("r1c3");
        _handler.Get("/Sheet1/A2").Text.Should().Be("r3c1");
        _handler.Get("/Sheet1/B2").Text.Should().Be("r3c3");
    }
}
