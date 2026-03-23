// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt rounds 56-75: Three targeted bugs found via white-box code review.
///
/// Bug A — headEnd="open" is silently mapped to "triangle" (filled arrow).
///   Root cause: ParseLineEndType() has no "open" case; the fallback maps every
///   unrecognized token to Drawing.LineEndValues.Triangle. In OOXML the open
///   (non-filled) arrowhead is LineEndValues.Arrow, whose InnerText is "arrow".
///   The existing aliases "triangle" and "arrow" both map to Triangle, which
///   leaves no way to request the open style by name.
///   Fix: add "open" => Drawing.LineEndValues.Arrow to ParseLineEndType().
///
/// Bug B — Set("/slide[N]/connector[M]", {headEnd/tailEnd}) is silently rejected.
///   Root cause: the connector Set switch (PowerPointHandler.Set.cs ~1240) has
///   no case for "headend" / "tailend". The Add path handles them correctly.
///   Fix: add case "headend"/"tailend" to the connector Set switch that finds
///   or creates the Outline element and upserts HeadEnd/TailEnd children.
///
/// Bug C — Set("/slide[N]/picture[M]", {rotation}) is silently rejected.
///   Root cause: the picture Set switch (PowerPointHandler.Set.cs ~737) has no
///   "rotation"/"rotate" case, yet the unsupported-property error message lists
///   "rotation" as a valid prop — a misleading promise with no implementation.
///   Fix: add case "rotation"/"rotate" to the picture Set switch, mirroring the
///   connector and group Set implementations.
/// </summary>
public class PptxRound56Tests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext = ".pptx")
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // =========================================================================
    // Bug A — ParseLineEndType: "open" alias is missing, falls through to Triangle
    //
    // "open" should produce LineEndValues.Arrow (OOXML InnerText = "arrow"),
    // which is the open (non-filled) arrowhead distinct from the filled Triangle.
    // Currently "open" hits the default fallback => Triangle => InnerText "triangle".
    // =========================================================================

    [Fact]
    public void BugA_Connector_Add_HeadEndOpen_StoresArrowNotTriangle()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["headEnd"] = "open",
            ["tailEnd"] = "none",
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Should().NotBeNull();

        // "open" must NOT resolve to "triangle" (the filled arrowhead).
        // After fix it should resolve to the open arrowhead whose OOXML
        // InnerText is "arrow".
        node!.Format.Should().ContainKey("headEnd");
        node.Format["headEnd"].Should().NotBe("triangle",
            "headEnd='open' must produce the open (non-filled) arrowhead, not a filled triangle");
        node.Format["headEnd"].Should().Be("arrow",
            "the OOXML open arrowhead has LineEndValues.Arrow with InnerText 'arrow'");
    }

    [Fact]
    public void BugA_Connector_Add_HeadEndOpen_IsDistinctFromTriangle()
    {
        // Verify that "open" and "triangle" produce different stored values,
        // proving they map to different LineEndValues.
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new() { ["headEnd"] = "open" });
        handler.Add("/slide[1]", "connector", null, new() { ["headEnd"] = "triangle" });

        var openNode     = handler.Get("/slide[1]/connector[1]");
        var triangleNode = handler.Get("/slide[1]/connector[2]");

        openNode!.Format["headEnd"].Should().NotBe(
            triangleNode!.Format["headEnd"],
            "open arrowhead and triangle arrowhead are distinct OOXML types");
    }

    [Fact]
    public void BugA_Connector_Add_HeadEndOpen_Persists()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Add("/slide[1]", "connector", null, new() { ["headEnd"] = "open" });
        }

        // Reopen and verify persistence
        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/connector[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("headEnd");
        node.Format["headEnd"].Should().Be("arrow");
    }

    // =========================================================================
    // Bug B — Set connector headEnd/tailEnd is not implemented
    //
    // Add supports headEnd/tailEnd. Set silently rejects them with an
    // "unsupported property" error. After fix, Set should upsert HeadEnd/TailEnd
    // on the connector's outline element.
    // =========================================================================

    [Fact]
    public void BugB_Connector_Set_HeadEnd_IsAccepted()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new() { ["headEnd"] = "none" });

        // Change headEnd from "none" to "triangle" via Set
        var unsupported = handler.Set("/slide[1]/connector[1]", new() { ["headEnd"] = "triangle" });

        unsupported.Should().BeEmpty(
            "headEnd is a documented connector property and Set must support it");

        var node = handler.Get("/slide[1]/connector[1]");
        node!.Format.Should().ContainKey("headEnd");
        node.Format["headEnd"].Should().Be("triangle");
    }

    [Fact]
    public void BugB_Connector_Set_TailEnd_IsAccepted()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new() { ["tailEnd"] = "none" });

        var unsupported = handler.Set("/slide[1]/connector[1]", new() { ["tailEnd"] = "diamond" });

        unsupported.Should().BeEmpty(
            "tailEnd is a documented connector property and Set must support it");

        var node = handler.Get("/slide[1]/connector[1]");
        node!.Format.Should().ContainKey("tailEnd");
        node.Format["tailEnd"].Should().Be("diamond");
    }

    [Fact]
    public void BugB_Connector_Set_HeadEnd_UpdatesExistingValue()
    {
        // Verify round-trip: Add with one type, Set to another, Get returns new type.
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new() { ["headEnd"] = "triangle" });

        // Verify initial value
        var before = handler.Get("/slide[1]/connector[1]");
        before!.Format["headEnd"].Should().Be("triangle");

        // Change to stealth
        handler.Set("/slide[1]/connector[1]", new() { ["headEnd"] = "stealth" });

        var after = handler.Get("/slide[1]/connector[1]");
        after!.Format["headEnd"].Should().Be("stealth",
            "Set should overwrite the existing headEnd type on the connector outline");
    }

    [Fact]
    public void BugB_Connector_Set_HeadEnd_Persists()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Add("/slide[1]", "connector", null, new() { ["headEnd"] = "none" });
            h.Set("/slide[1]/connector[1]", new() { ["headEnd"] = "triangle" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/connector[1]");
        node!.Format["headEnd"].Should().Be("triangle",
            "headEnd set via Set must survive save/reopen");
    }

    // =========================================================================
    // Bug C — Set picture rotation is not implemented (misleading error message)
    //
    // The unsupported-property error message for pictures includes "rotation"
    // in its hint string, implying it is supported, but the switch has no case
    // for it. After fix, Set picture rotation must update the Transform2D.Rotation
    // attribute and round-trip correctly via Get.
    // =========================================================================

    [Fact]
    public void BugC_Picture_Set_Rotation_IsAccepted()
    {
        // Use a tiny 1x1 transparent PNG as the picture source to avoid file-not-found.
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        _tempFiles.Add(pngPath);
        WriteTiny1x1Png(pngPath);

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "picture", null, new()
        {
            ["path"] = pngPath,
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
        });

        // Set rotation (45 degrees)
        var unsupported = handler.Set("/slide[1]/picture[1]", new() { ["rotation"] = "45" });

        unsupported.Should().BeEmpty(
            "rotation is listed in the valid-picture-props hint and must be implemented");
    }

    [Fact]
    public void BugC_Picture_Set_Rotation_RoundTrips()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        _tempFiles.Add(pngPath);
        WriteTiny1x1Png(pngPath);

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "picture", null, new()
        {
            ["path"] = pngPath,
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
        });

        handler.Set("/slide[1]/picture[1]", new() { ["rotation"] = "90" });

        var node = handler.Get("/slide[1]/picture[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("rotation",
            "Get must expose the rotation set via Set");
        // Rotation should round-trip as 90 (degrees)
        var rot = node.Format["rotation"]?.ToString();
        rot.Should().NotBeNullOrEmpty();
        double.Parse(rot!).Should().BeApproximately(90.0, 1.0,
            "rotation should round-trip as 90 degrees");
    }

    [Fact]
    public void BugC_Picture_Set_Rotation_Persists()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        _tempFiles.Add(pngPath);
        WriteTiny1x1Png(pngPath);

        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Add("/slide[1]", "picture", null, new()
            {
                ["path"] = pngPath,
                ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "3cm",
            });
            h.Set("/slide[1]/picture[1]", new() { ["rotation"] = "45" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/picture[1]");
        node!.Format.Should().ContainKey("rotation");
        double.Parse(node.Format["rotation"]!.ToString()!).Should().BeApproximately(45.0, 1.0,
            "rotation set via Set must survive save/reopen");
    }

    // =========================================================================
    // Bug D — Slide background image requires "image:" prefix; bare path fails
    //
    // ApplySlideBackground() only routes to ApplyBackgroundImageFill() when the
    // value starts with "image:". A raw file path (e.g. "/tmp/foo.png") falls
    // through to BuildSolidFill() which calls SanitizeColorForOoxml() on the
    // path string, producing garbage hex and silently corrupting the background.
    //
    // Fix: in ApplySlideBackground(), before the solid-fill fallback, check
    // whether the value is an existing file path (or has a known image extension)
    // and auto-route it to ApplyBackgroundImageFill(), or alternatively document
    // and enforce the "image:" prefix with a clear error message rather than
    // silently misinterpreting the path as a color.
    //
    // The tests below verify the two user-visible symptoms:
    //   D1 — "image:/path/to/file.png" roundtrips to Format["background"]=="image"
    //   D2 — A bare path without "image:" prefix should either also work (if
    //         auto-detection is added) OR throw a meaningful ArgumentException
    //         rather than silently storing a corrupted color value.
    // =========================================================================

    [Fact]
    public void BugD_Slide_Background_ImagePrefix_RoundTrips()
    {
        // "image:/path" form must be accepted and stored as a blip fill.
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        _tempFiles.Add(pngPath);
        WriteTiny1x1Png(pngPath);

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        var unsupported = handler.Set("/slide[1]", new() { ["background"] = $"image:{pngPath}" });

        unsupported.Should().BeEmpty("background=image:path is a documented slide property");

        var node = handler.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("background",
            "a slide with a background image should expose Format['background']");
        node.Format["background"].Should().Be("image",
            "ReadSlideBackground returns 'image' for blip-filled backgrounds");
    }

    [Fact]
    public void BugD_Slide_Background_ImagePrefix_Persists()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        _tempFiles.Add(pngPath);
        WriteTiny1x1Png(pngPath);

        var path = CreateTemp();
        BlankDocCreator.Create(path);

        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new());
            h.Set("/slide[1]", new() { ["background"] = $"image:{pngPath}" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]");
        node!.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("image",
            "background image set via Set must survive save/reopen");
    }

    [Fact]
    public void BugD_Slide_Background_BarePath_DoesNotSilentlyCorrupt()
    {
        // A bare file path without "image:" prefix must NOT silently store a
        // garbage hex value. The expected behaviour post-fix is either:
        //   (a) auto-detect and treat it as an image path, or
        //   (b) throw an ArgumentException with a helpful message.
        // This test enforces that the stored background value is never a
        // hex-looking string derived from path characters, which is the
        // current broken behavior.
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        _tempFiles.Add(pngPath);
        WriteTiny1x1Png(pngPath);

        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        // The call should either succeed (auto-detect) or throw ArgumentException.
        // In either case, it must NOT silently store a hex color derived from
        // path characters like "000000" (the silent corruption symptom).
        var threwExpectedException = false;
        try
        {
            handler.Set("/slide[1]", new() { ["background"] = pngPath });
        }
        catch (ArgumentException)
        {
            threwExpectedException = true;
        }

        if (!threwExpectedException)
        {
            // If it did not throw, it must have stored an "image" background,
            // not a garbage hex color.
            var node = handler.Get("/slide[1]");
            node!.Format.Should().ContainKey("background");
            node.Format["background"].Should().Be("image",
                "auto-detecting a bare file path should store a blip fill, not a hex color");
        }
        // else: threw ArgumentException — acceptable. Test passes.
    }

    // =========================================================================
    // Helper — minimal 1×1 PNG (67 bytes, valid PNG header + IDAT + IEND)
    // =========================================================================

    private static void WriteTiny1x1Png(string filePath)
    {
        // A valid 1×1 transparent PNG generated offline.
        var pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
        File.WriteAllBytes(filePath, pngBytes);
    }
}
