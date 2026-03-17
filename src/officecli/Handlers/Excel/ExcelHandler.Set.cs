// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;


namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        // Parse path: /SheetName, /SheetName/A1, /SheetName/A1:D1, /SheetName/col[A], /SheetName/row[1], /SheetName/autofilter
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null)
            throw new ArgumentException($"Sheet not found: {sheetName}");

        // Sheet-level Set (path is just /SheetName)
        if (segments.Length < 2)
        {
            return SetSheetLevel(worksheet, properties);
        }

        var cellRef = segments[1];

        // Handle /SheetName/autofilter
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            return SetAutoFilter(worksheet, properties);
        }

        // Handle /SheetName/cf[N]
        var cfSetMatch = Regex.Match(cellRef, @"^cf\[(\d+)\]$");
        if (cfSetMatch.Success)
        {
            var cfIdx = int.Parse(cfSetMatch.Groups[1].Value);
            var ws = GetSheet(worksheet);
            var cfElements = ws.Elements<ConditionalFormatting>().ToList();
            if (cfIdx < 1 || cfIdx > cfElements.Count)
                throw new ArgumentException($"CF {cfIdx} not found (total: {cfElements.Count})");

            var cf = cfElements[cfIdx - 1];
            var unsup = new List<string>();
            var rule = cf.Elements<ConditionalFormattingRule>().FirstOrDefault();

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "sqref":
                        cf.SequenceOfReferences = new ListValue<StringValue>(
                            value.Split(' ').Select(s => new StringValue(s)));
                        break;
                    case "color":
                        var dbColor = rule?.GetFirstChild<DataBar>()?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                        if (dbColor != null) dbColor.Rgb = (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant();
                        else unsup.Add(key);
                        break;
                    case "mincolor":
                        var csColors = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                        if (csColors != null && csColors.Count >= 2)
                            csColors[0].Rgb = (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant();
                        else unsup.Add(key);
                        break;
                    case "maxcolor":
                        var csColors2 = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                        if (csColors2 != null && csColors2.Count >= 2)
                            csColors2[^1].Rgb = (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant();
                        else unsup.Add(key);
                        break;
                    case "iconset":
                        var iconSetEl = rule?.GetFirstChild<IconSet>();
                        if (iconSetEl != null)
                            iconSetEl.IconSetValue = new EnumValue<IconSetValues>(new IconSetValues(value));
                        else unsup.Add(key);
                        break;
                    case "reverse":
                        var isEl = rule?.GetFirstChild<IconSet>();
                        if (isEl != null) isEl.Reverse = bool.Parse(value);
                        else unsup.Add(key);
                        break;
                    case "showvalue":
                        var isEl2 = rule?.GetFirstChild<IconSet>();
                        if (isEl2 != null) isEl2.ShowValue = bool.Parse(value);
                        else unsup.Add(key);
                        break;
                    default:
                        unsup.Add(key);
                        break;
                }
            }
            ws.Save();
            return unsup;
        }

        // Handle /SheetName/col[X]
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Z]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colName = colMatch.Groups[1].Value.ToUpperInvariant();
            return SetColumn(worksheet, colName, properties);
        }

        // Handle /SheetName/row[N]
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            return SetRow(worksheet, rowIdx, properties);
        }

        // Handle /SheetName/chart[N]
        var chartMatch = Regex.Match(cellRef, @"^chart\[(\d+)\]$");
        if (chartMatch.Success)
        {
            var chartIdx = int.Parse(chartMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart;
            if (drawingsPart == null)
                throw new ArgumentException("No charts in this sheet");
            var chartParts = drawingsPart.ChartParts.ToList();
            if (chartIdx < 1 || chartIdx > chartParts.Count)
                throw new ArgumentException($"Chart {chartIdx} not found");
            var chartPart = chartParts[chartIdx - 1];

            var unsup = new List<string>();
            var chart = chartPart.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            if (chart == null) return unsup;

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "title":
                        var titleEl = chart.Title;
                        if (titleEl != null)
                        {
                            var titleRun = titleEl.Descendants<DocumentFormat.OpenXml.Drawing.Run>().FirstOrDefault();
                            if (titleRun?.Text != null) titleRun.Text.Text = value;
                        }
                        break;
                    default:
                        unsup.Add(key);
                        break;
                }
            }
            chartPart.ChartSpace?.Save();
            return unsup;
        }

        // Handle /SheetName/A1:D1 (range — merge/unmerge)
        if (cellRef.Contains(':'))
        {
            var firstPartRange = cellRef.Split(':')[0];
            bool isRangeRef = Regex.IsMatch(firstPartRange, @"^[A-Z]+\d+$", RegexOptions.IgnoreCase);
            if (isRangeRef)
            {
                return SetRange(worksheet, cellRef.ToUpperInvariant(), properties);
            }
        }

        // Check if path is a cell reference or generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = Regex.IsMatch(firstPart, @"^[A-Z]+\d+", RegexOptions.IgnoreCase);
        if (!isCellRef)
        {
            // Generic XML fallback: navigate to element and set attributes
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                throw new ArgumentException($"Element not found: {cellRef}");
            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            GetSheet(worksheet).Save();
            return unsup;
        }

        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            GetSheet(worksheet).Append(sheetData);
        }

        var cell = FindOrCreateCell(sheetData, cellRef);

        // Separate content props from style props
        var styleProps = new Dictionary<string, string>();
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            if (ExcelStyleManager.IsStyleKey(key))
            {
                styleProps[key] = value;
                continue;
            }

            switch (key.ToLowerInvariant())
            {
                case "value":
                    cell.CellValue = new CellValue(value);
                    // Auto-detect type
                    if (double.TryParse(value, out _))
                        cell.DataType = null; // Number is default
                    else
                    {
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                    break;
                case "formula":
                    cell.CellFormula = new CellFormula(value);
                    cell.CellValue = null;
                    break;
                case "type":
                    cell.DataType = value.ToLowerInvariant() switch
                    {
                        "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                        _ => cell.DataType
                    };
                    break;
                case "clear":
                    cell.CellValue = null;
                    cell.CellFormula = null;
                    break;
                case "link":
                {
                    var ws = GetSheet(worksheet);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        hyperlinksEl?.Elements<Hyperlink>()
                            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                            .ToList().ForEach(h => h.Remove());
                    }
                    else
                    {
                        var hlRel = worksheet.AddHyperlinkRelationship(new Uri(value), isExternal: true);
                        if (hyperlinksEl == null)
                        {
                            hyperlinksEl = new Hyperlinks();
                            ws.AppendChild(hyperlinksEl);
                        }
                        hyperlinksEl.Elements<Hyperlink>()
                            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                            .ToList().ForEach(h => h.Remove());
                        hyperlinksEl.AppendChild(new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id });
                    }
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                        unsupported.Add(key);
                    break;
            }
        }

        // Apply style properties if any
        if (styleProps.Count > 0)
        {
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var styleManager = new ExcelStyleManager(workbookPart);
            cell.StyleIndex = styleManager.ApplyStyle(cell, styleProps);
        }

        GetSheet(worksheet).Save();
        return unsupported;
    }

    // ==================== Sheet-level Set (freeze panes) ====================

    private List<string> SetSheetLevel(WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "freeze":
                {
                    var sheetViews = ws.GetFirstChild<SheetViews>();
                    if (sheetViews == null)
                    {
                        sheetViews = new SheetViews();
                        ws.InsertAt(sheetViews, 0);
                    }
                    var sheetView = sheetViews.GetFirstChild<SheetView>();
                    if (sheetView == null)
                    {
                        sheetView = new SheetView { WorkbookViewId = 0 };
                        sheetViews.AppendChild(sheetView);
                    }

                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove freeze
                        var existingPane = sheetView.GetFirstChild<Pane>();
                        existingPane?.Remove();
                    }
                    else
                    {
                        // Parse cell reference for freeze position
                        // "A2" = freeze row 1, "B1" = freeze col A, "B2" = freeze row 1 + col A
                        var (col, row) = ParseCellReference(value.ToUpperInvariant());
                        var colSplit = ColumnNameToIndex(col) - 1; // 0-based: B=1 means split at 1
                        var rowSplit = row - 1; // 0-based: 2 means split at 1

                        // Remove existing pane
                        var existingPane = sheetView.GetFirstChild<Pane>();
                        existingPane?.Remove();

                        var activePane = (colSplit > 0 && rowSplit > 0) ? PaneValues.BottomRight
                            : (rowSplit > 0) ? PaneValues.BottomLeft
                            : PaneValues.TopRight;

                        var pane = new Pane
                        {
                            TopLeftCell = value.ToUpperInvariant(),
                            State = PaneStateValues.Frozen,
                            ActivePane = activePane
                        };
                        if (rowSplit > 0) pane.VerticalSplit = rowSplit;
                        if (colSplit > 0) pane.HorizontalSplit = colSplit;

                        sheetView.InsertAt(pane, 0);
                    }
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ws.Save();
        return unsupported;
    }

    // ==================== Range Set (merge/unmerge) ====================

    private List<string> SetRange(WorksheetPart worksheet, string rangeRef, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "merge":
                {
                    bool doMerge = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);

                    if (doMerge)
                    {
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        if (mergeCells == null)
                        {
                            mergeCells = new MergeCells();
                            // MergeCells must be after SheetData, before Hyperlinks/Drawing
                            var sheetData = ws.GetFirstChild<SheetData>();
                            if (sheetData != null)
                                sheetData.InsertAfterSelf(mergeCells);
                            else
                                ws.AppendChild(mergeCells);
                        }

                        // Avoid duplicate
                        var existing = mergeCells.Elements<MergeCell>()
                            .FirstOrDefault(m => m.Reference?.Value?.Equals(rangeRef, StringComparison.OrdinalIgnoreCase) == true);
                        if (existing == null)
                        {
                            mergeCells.AppendChild(new MergeCell { Reference = rangeRef });
                        }
                    }
                    else
                    {
                        // Unmerge: remove the MergeCell for this range
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        if (mergeCells != null)
                        {
                            var mc = mergeCells.Elements<MergeCell>()
                                .FirstOrDefault(m => m.Reference?.Value?.Equals(rangeRef, StringComparison.OrdinalIgnoreCase) == true);
                            mc?.Remove();

                            // Remove empty MergeCells element
                            if (!mergeCells.HasChildren)
                                mergeCells.Remove();
                        }
                    }
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ws.Save();
        return unsupported;
    }

    // ==================== Column Set (width, hidden) ====================

    private List<string> SetColumn(WorksheetPart worksheet, string colName, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);
        var colIdx = (uint)ColumnNameToIndex(colName);

        var columns = ws.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData != null)
                ws.InsertBefore(columns, sheetData);
            else
                ws.AppendChild(columns);
        }

        // Find existing column definition or create one
        var col = columns.Elements<Column>()
            .FirstOrDefault(c => c.Min?.Value <= colIdx && c.Max?.Value >= colIdx);
        if (col == null)
        {
            col = new Column { Min = colIdx, Max = colIdx, Width = 8.43, CustomWidth = true };
            columns.AppendChild(col);
        }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "width":
                    col.Width = double.Parse(value);
                    col.CustomWidth = true;
                    break;
                case "hidden":
                    col.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ws.Save();
        return unsupported;
    }

    // ==================== Row Set (height, hidden) ====================

    private List<string> SetRow(WorksheetPart worksheet, uint rowIdx, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
            throw new ArgumentException("Sheet has no data");

        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
        if (row == null)
        {
            // Create the row
            row = new Row { RowIndex = rowIdx };
            var afterRow = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < rowIdx);
            if (afterRow != null)
                afterRow.InsertAfterSelf(row);
            else
                sheetData.InsertAt(row, 0);
        }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "height":
                    row.Height = double.Parse(value);
                    row.CustomHeight = true;
                    break;
                case "hidden":
                    row.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ws.Save();
        return unsupported;
    }

    // ==================== AutoFilter Set ====================

    private List<string> SetAutoFilter(WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "range":
                {
                    var autoFilter = ws.GetFirstChild<AutoFilter>();
                    if (autoFilter == null)
                    {
                        autoFilter = new AutoFilter();
                        // AutoFilter goes after SheetData (after MergeCells if present)
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        var sheetData = ws.GetFirstChild<SheetData>();
                        if (mergeCells != null)
                            mergeCells.InsertAfterSelf(autoFilter);
                        else if (sheetData != null)
                            sheetData.InsertAfterSelf(autoFilter);
                        else
                            ws.AppendChild(autoFilter);
                    }
                    autoFilter.Reference = value.ToUpperInvariant();
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ws.Save();
        return unsupported;
    }
}
