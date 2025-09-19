// File: Program.cs
// Build: dotnet build
// Run:   dotnet run -- --input deck.pptx --out report.json --media-dir media --embed-images false
// Requires: dotnet add package DocumentFormat.OpenXml
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace PptStructureInspector
{
    public class Program
    {
        public static int Main(string[] args)
        {
            var inputOpt = new Option<FileInfo?>(name: "--input", description: "Path to .pptx", getDefaultValue: () => null);
            var outOpt   = new Option<FileInfo?>(name: "--out", description: "Path to output JSON", getDefaultValue: () => new FileInfo("report.json"));
            var mediaDirOpt = new Option<DirectoryInfo?>(name: "--media-dir", description: "Directory to export images to (created if missing)", getDefaultValue: () => new DirectoryInfo("media"));
            var embedImagesOpt = new Option<bool>(name: "--embed-images", description: "Embed images as base64 in JSON (may be large)", getDefaultValue: () => false);

            var root = new RootCommand("PowerPoint Structure Inspector (placeholders, charts, tables, images)");
            root.AddOption(inputOpt);
            root.AddOption(outOpt);
            root.AddOption(mediaDirOpt);
            root.AddOption(embedImagesOpt);

            root.SetHandler((FileInfo? input, FileInfo? output, DirectoryInfo? mediaDir, bool embedImages) =>
            {
                if (input is null || !input.Exists) throw new FileNotFoundException("Input .pptx not found", input?.FullName);
                if (output is null) throw new ArgumentNullException("--out");
                if (mediaDir is null) throw new ArgumentNullException("--media-dir");

                var report = PptReader.Read(input.FullName, mediaDir.FullName, embedImages);
                var json = JsonSerializer.Serialize(report, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                });
                File.WriteAllText(output.FullName, json);
                Console.WriteLine($"Wrote {output.FullName}");
            },
            inputOpt, outOpt, mediaDirOpt, embedImagesOpt);

            return root.Invoke(args);
        }
    }

    public static class PptReader
    {
        public static DeckReport Read(string pptxPath, string mediaDir, bool embedImages)
        {
            Directory.CreateDirectory(mediaDir);
            using var doc = PresentationDocument.Open(pptxPath, false);
            var pres = doc.PresentationPart ?? throw new InvalidOperationException("Invalid deck.");

            var slideIds = pres.Presentation!.SlideIdList!.Elements<SlideId>().ToList();
            var deck = new DeckReport
            {
                File = Path.GetFileName(pptxPath),
                SlideCount = slideIds.Count,
                Slides = new()
            };

            for (int i = 0; i < slideIds.Count; i++)
            {
                var sid = slideIds[i];
                var sp = (SlidePart)pres.GetPartById(sid.RelationshipId!);

                var layoutMap = BuildLayoutPlaceholderMap(sp.SlideLayoutPart?.SlideLayout);
                var ctx = new WalkContext
                {
                    SlidePart = sp,
                    LayoutPlaceholderMap = layoutMap,
                    MediaDir = mediaDir,
                    EmbedImages = embedImages,
                    SlideIndex = i + 1
                };

                var slideInfo = new SlideReport { SlideNumber = i + 1, Placeholders = new(), OtherShapes = new() };
                var tree = sp.Slide.CommonSlideData!.ShapeTree!;

                // Flatten top-level shapes; recurse into groups
                foreach (var child in tree.ChildElements)
                {
                    WalkShapeElement(child, slideInfo, ctx);
                }

                // Best-effort title
                slideInfo.Title = slideInfo.Placeholders
                    .Where(p => EqualsCI(p.PlaceholderType, "Title") || EqualsCI(p.PlaceholderType, "CtrTitle"))
                    .Select(p => p.Text).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));

                deck.Slides.Add(slideInfo);
            }

            return deck;
        }

        static void WalkShapeElement(DocumentFormat.OpenXml.OpenXmlElement elem, SlideReport slideInfo, WalkContext ctx)
        {
            switch (elem)
            {
                case Shape shp:
                    HandleShape(shp, slideInfo, ctx);
                    break;
                case GraphicFrame gf:
                    HandleGraphicFrame(gf, slideInfo, ctx);
                    break;
                case Picture pic:
                    HandlePicture(pic, slideInfo, ctx);
                    break;
                case GroupShape grp:
                    foreach (var child in grp.ChildElements)
                        WalkShapeElement(child, slideInfo, ctx);
                    break;
                default:
                    break;
            }
        }

        static void HandleShape(Shape shp, SlideReport slideInfo, WalkContext ctx)
        {
            var ph = shp.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .GetFirstChild<PlaceholderShape>();
            var text = ExtractText(shp.TextBody);
            var bounds = GetBounds(shp.Transform2D);

            if (ph != null)
            {
                var (idx, type) = PlaceholderType(ph, ctx.LayoutPlaceholderMap);
                slideInfo.Placeholders.Add(new PlaceholderReport
                {
                    Index = idx,
                    PlaceholderType = type,
                    ShapeKind = "Shape",
                    Text = text,
                    Bounds = bounds
                });
            }
            else
            {
                slideInfo.OtherShapes.Add(new OtherShapeReport
                {
                    ShapeKind = "Shape",
                    Text = text,
                    Bounds = bounds
                });
            }
        }

        static void HandleGraphicFrame(GraphicFrame gf, SlideReport slideInfo, WalkContext ctx)
        {
            var ph = gf.NonVisualGraphicFrameProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .GetFirstChild<PlaceholderShape>();
            var bounds = GetBounds(gf.Transform);

            // Table?
            var tbl = gf.Graphic?.GraphicData?.GetFirstChild<A.Table>();
            if (tbl != null)
            {
                var tinfo = new TableInfo
                {
                    Rows = tbl.Descendants<A.TableRow>().Count(),
                    Columns = tbl.TableGrid?.GridColumnList?.Count ?? 0,
                    TextPreview = ExtractTableText(tbl)
                };

                if (ph != null)
                {
                    var (idx, type) = PlaceholderType(ph, ctx.LayoutPlaceholderMap);
                    slideInfo.Placeholders.Add(new PlaceholderReport
                    {
                        Index = idx,
                        PlaceholderType = type,
                        ShapeKind = "Table",
                        Table = tinfo,
                        Bounds = bounds
                    });
                }
                else
                {
                    slideInfo.OtherShapes.Add(new OtherShapeReport
                    {
                        ShapeKind = "Table",
                        Table = tinfo,
                        Bounds = bounds
                    });
                }
                return;
            }

            // Chart?
            var chartRef = gf.Graphic?.GraphicData?.GetFirstChild<C.Chart>();
            if (chartRef != null)
            {
                string? kind = null;
                try
                {
                    var relId = chartRef.Id!;
                    var cpart = (ChartPart)ctx.SlidePart.GetPartById(relId);
                    kind = DetectChartKind(cpart);
                }
                catch { }

                var cinfo = new ChartInfo { ChartType = kind };

                if (ph != null)
                {
                    var (idx, type) = PlaceholderType(ph, ctx.LayoutPlaceholderMap);
                    slideInfo.Placeholders.Add(new PlaceholderReport
                    {
                        Index = idx,
                        PlaceholderType = type,
                        ShapeKind = "Chart",
                        Chart = cinfo,
                        Bounds = bounds
                    });
                }
                else
                {
                    slideInfo.OtherShapes.Add(new OtherShapeReport
                    {
                        ShapeKind = "Chart",
                        Chart = cinfo,
                        Bounds = bounds
                    });
                }
                return;
            }

            // Other
            if (ph != null)
            {
                var (idx2, type2) = PlaceholderType(ph, ctx.LayoutPlaceholderMap);
                slideInfo.Placeholders.Add(new PlaceholderReport
                {
                    Index = idx2,
                    PlaceholderType = type2,
                    ShapeKind = "GraphicFrame",
                    Bounds = bounds
                });
            }
            else
            {
                slideInfo.OtherShapes.Add(new OtherShapeReport
                {
                    ShapeKind = "GraphicFrame",
                    Bounds = bounds
                });
            }
        }

        static void HandlePicture(Picture pic, SlideReport slideInfo, WalkContext ctx)
        {
            var ph = pic.NonVisualPictureProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .GetFirstChild<PlaceholderShape>();
            var blip = pic.BlipFill?.Blip;
            var relId = blip?.Embed?.Value;
            ImageInfo? iinfo = null;
            var bounds = GetBounds(pic.ShapeProperties?.Transform2D);

            if (!string.IsNullOrEmpty(relId))
            {
                try
                {
                    var part = (ImagePart)ctx.SlidePart.GetPartById(relId);
                    using var s = part.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    var bytes = ms.ToArray();

                    var ext = ExtensionFromContentType(part.ContentType);
                    var fileName = $"slide{ctx.SlideIndex}_img_{Guid.NewGuid():N}.{ext}";
                    var outPath = Path.Combine(ctx.MediaDir, fileName);
                    File.WriteAllBytes(outPath, bytes);

                    iinfo = new ImageInfo
                    {
                        ContentType = part.ContentType,
                        Bytes = ctx.EmbedImages ? Convert.ToBase64String(bytes) : null,
                        File = Path.Combine(Path.GetFileName(ctx.MediaDir), fileName).Replace('\\','/'),
                        SizeBytes = bytes.Length
                    };
                }
                catch { /* ignore */ }
            }

            if (ph != null)
            {
                var (idx, type) = PlaceholderType(ph, ctx.LayoutPlaceholderMap);
                slideInfo.Placeholders.Add(new PlaceholderReport
                {
                    Index = idx,
                    PlaceholderType = type,
                    ShapeKind = "Picture",
                    Image = iinfo,
                    Bounds = bounds
                });
            }
            else
            {
                slideInfo.OtherShapes.Add(new OtherShapeReport
                {
                    ShapeKind = "Picture",
                    Image = iinfo,
                    Bounds = bounds
                });
            }
        }

        static string ExtensionFromContentType(string contentType) => contentType switch
        {
            "image/png" => "png",
            "image/jpeg" => "jpg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            "image/tiff" => "tiff",
            "image/x-emf" => "emf",
            "image/x-wmf" => "wmf",
            _ => "bin"
        };

        static (int? idx, string type) PlaceholderType(PlaceholderShape ph, Dictionary<int, string> layoutMap)
        {
            int? idx = ph.Index?.Value;
            string type = ph.Type?.Value.ToString()
                          ?? (idx.HasValue && layoutMap.TryGetValue(idx.Value, out var t) ? t : "Body");
            return (idx, type);
        }

        static Bounds? GetBounds(A.Transform2D? xfrm)
        {
            if (xfrm == null) return null;
            return new Bounds
            {
                X = xfrm.Offset?.X?.Value ?? 0,
                Y = xfrm.Offset?.Y?.Value ?? 0,
                CX = xfrm.Extents?.Cx?.Value ?? 0,
                CY = xfrm.Extents?.Cy?.Value ?? 0
            };
        }

        static string? ExtractText(A.TextBody? tb)
        {
            if (tb == null) return null;
            var paras = tb.Descendants<A.Paragraph>()
                          .Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text)))
                          .ToList();
            var text = string.Join("\n", paras).Trim();
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        static string? ExtractTableText(A.Table tbl)
        {
            var rows = new List<string>();
            foreach (var tr in tbl.Descendants<A.TableRow>())
            {
                var cells = tr.Descendants<A.TableCell>()
                              .Select(tc => ExtractText(tc.TextBody) ?? "")
                              .ToArray();
                rows.Add(string.Join(" | ", cells));
            }
            var joined = string.Join("\n", rows).Trim();
            return string.IsNullOrWhiteSpace(joined) ? null : joined;
        }

        static string? DetectChartKind(ChartPart cp)
        {
            var plot = cp.ChartSpace?.Chart?.PlotArea;
            if (plot == null) return null;
            if (plot.PieChart != null) return "Pie";
            if (plot.BarChart != null && plot.BarChart.BarDirection?.Val?.Value == C.BarDirectionValues.Column) return "Column";
            if (plot.BarChart != null && plot.BarChart.BarDirection?.Val?.Value == C.BarDirectionValues.Bar) return "Bar";
            if (plot.LineChart != null) return "Line";
            if (plot.ScatterChart != null) return "Scatter";
            if (plot.AreaChart != null) return "Area";
            if (plot.DoughnutChart != null) return "Doughnut";
            if (plot.BubbleChart != null) return "Bubble";
            return "Unknown";
        }

        static Dictionary<int, string> BuildLayoutPlaceholderMap(SlideLayout? layout)
        {
            var map = new Dictionary<int, string>();
            if (layout?.CommonSlideData?.ShapeTree == null) return map;

            foreach (var shp in layout.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                var ph = shp.NonVisualShapeProperties?
                           .ApplicationNonVisualDrawingProperties?
                           .GetFirstChild<PlaceholderShape>();
                if (ph?.Index?.Value is int idx)
                {
                    var t = ph.Type?.Value.ToString() ?? "Body";
                    map[idx] = t;
                }
            }
            foreach (var gf in layout.CommonSlideData.ShapeTree.Elements<GraphicFrame>())
            {
                var ph = gf.NonVisualGraphicFrameProperties?
                           .ApplicationNonVisualDrawingProperties?
                           .GetFirstChild<PlaceholderShape>();
                if (ph?.Index?.Value is int idx)
                {
                    var t = ph.Type?.Value.ToString() ?? "Body";
                    map[idx] = t;
                }
            }
            return map;
        }

        static bool EqualsCI(string? a, string? b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    }

    // ---------------- Models for JSON ----------------
    public class DeckReport
    {
        public string File { get; set; } = "";
        public int SlideCount { get; set; }
        public List<SlideReport> Slides { get; set; } = new();
    }

    public class SlideReport
    {
        public int SlideNumber { get; set; }
        public string? Title { get; set; }
        public List<PlaceholderReport> Placeholders { get; set; } = new();
        public List<OtherShapeReport> OtherShapes { get; set; } = new();
    }

    public class PlaceholderReport
    {
        public int? Index { get; set; }
        public string PlaceholderType { get; set; } = "Body"; // Title, CtrTitle, Body, Date, SldNum, Footer, etc.
        public string ShapeKind { get; set; } = "Shape";      // Shape | Table | Chart | Picture | GraphicFrame
        public string? Text { get; set; }
        public TableInfo? Table { get; set; }
        public ChartInfo? Chart { get; set; }
        public ImageInfo? Image { get; set; }
        public Bounds? Bounds { get; set; }
    }

    public class OtherShapeReport
    {
        public string ShapeKind { get; set; } = "Shape";      // Shape | Table | Chart | Picture | GraphicFrame
        public string? Text { get; set; }
        public TableInfo? Table { get; set; }
        public ChartInfo? Chart { get; set; }
        public ImageInfo? Image { get; set; }
        public Bounds? Bounds { get; set; }
    }

    public class TableInfo
    {
        public int Rows { get; set; }
        public int Columns { get; set; }
        public string? TextPreview { get; set; }
    }

    public class ChartInfo
    {
        public string? ChartType { get; set; } // Pie, Column, Bar, Line, Scatter, etc.
    }

    public class ImageInfo
    {
        public string? File { get; set; }        // relative path under mediaDir
        public int? SizeBytes { get; set; }
        public string? ContentType { get; set; }
        public string? Bytes { get; set; }       // base64 if --embed-images true
    }

    public class Bounds
    {
        public long X { get; set; }
        public long Y { get; set; }
        public long CX { get; set; }
        public long CY { get; set; }
    }

    // ---------------- Context ----------------
    internal class WalkContext
    {
        public SlidePart SlidePart { get; set; } = default!;
        public Dictionary<int, string> LayoutPlaceholderMap { get; set; } = default!;
        public string MediaDir { get; set; } = "media";
        public bool EmbedImages { get; set; } = false;
        public int SlideIndex { get; set; }
    }
}
