// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests that Remove returns a warning when formula cells reference shifted rows/columns.
/// No formula content is modified — the warning is purely informational.
/// </summary>
public class ExcelFormulaWarningTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelFormulaWarningTests()
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

    // ==================== Row delete warnings ====================

    [Fact]
    public void RemoveRow_NoFormulas_ReturnsNull()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "x" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "y" });

        var warning = _handler.Remove("/Sheet1/row[1]");

        warning.Should().BeNull();
    }

    [Fact]
    public void RemoveRow_FormulaReferencesShiftedRow_ReturnsWarning()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "30" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "SUM(A1:A3)" });

        var warning = _handler.Remove("/Sheet1/row[2]");

        warning.Should().NotBeNull();
        warning.Should().Contain("Warning:");
        warning.Should().Contain("B1");
    }

    [Fact]
    public void RemoveRow_FormulaInDeletedRow_NoWarning()
    {
        // The formula itself is in the deleted row — no warning needed (it's gone)
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/A2", new() { ["formula"] = "A1*2" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "30" });

        var warning = _handler.Remove("/Sheet1/row[2]");

        // A2 was in deleted row 2, so no surviving formula cell to warn about
        warning.Should().BeNull();
    }

    [Fact]
    public void RemoveRow_MultipleAffectedFormulas_AllMentioned()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "1" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "2" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "3" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "SUM(A1:A3)" });
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "A3*10" });

        var warning = _handler.Remove("/Sheet1/row[2]");

        warning.Should().NotBeNull();
        warning.Should().Contain("B1");
        warning.Should().Contain("C1");
    }

    // ==================== Column delete warnings ====================

    [Fact]
    public void RemoveCol_NoFormulas_ReturnsNull()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "x" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "y" });

        var warning = _handler.Remove("/Sheet1/col[A]");

        warning.Should().BeNull();
    }

    [Fact]
    public void RemoveCol_FormulaReferencesShiftedCol_ReturnsWarning()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "SUM(A1:B1)" });

        var warning = _handler.Remove("/Sheet1/col[A]");

        warning.Should().NotBeNull();
        warning.Should().Contain("Warning:");
        warning.Should().Contain("C1");
    }

    [Fact]
    public void RemoveCol_NativePath_WarningWorks()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "A1+B1" });

        var warning = _handler.Remove("Sheet1!col[B]");

        warning.Should().NotBeNull();
        warning.Should().Contain("C1");
    }

    // ==================== Warning content format ====================

    [Fact]
    public void Warning_ContainsCellCount()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "1" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "2" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "3" });
        // B3 has a formula referencing rows that will shift after deleting row 1
        _handler.Set("/Sheet1/B3", new() { ["formula"] = "SUM(A1:A3)" });

        var warning = _handler.Remove("/Sheet1/row[1]");

        warning.Should().NotBeNull();
        warning.Should().Contain("1 formula cell(s)");
        warning.Should().Contain("B3");
    }
}
