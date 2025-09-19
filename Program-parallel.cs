// File: Program.cs
// Build: dotnet build
// Run:   dotnet run -- --input deck.pptx --out report.json --media-dir media --embed-images false --max-degree 8
// Requires: dotnet add package DocumentFormat.OpenXml; dotnet add package System.CommandLine
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace PptStructureInspector.Parallelized
{
    public class Program
    {
        public static int Main(string[] args)
        {
            var inputOpt = new Option<FileInfo?>(name: "--input", description: "Path to .pptx", getDefaultValue: () => null);
            var outOpt   = new Option<FileInfo?>(name: "--out", description: "Path to output JSON", getDefaultValue: () => new FileInfo("report.json"));
            var mediaDirOpt = new Option<DirectoryInfo?>(name: "--media-dir", description: "Directory to export images (created if missing)", getDefaultValue: () => new DirectoryInfo("media"));
            var embedImagesOpt = new Option<bool>(name: "--embed-images", description: "Embed images as base64 in JSON (may be large)", getDefaultValue: () => false);
            var maxDegreeOpt = new Option<int>(name: "--max-degree", description: "Max parallelism (default=min(8, CPU count))", getDefaultValue: () => Math.Min(Environment.ProcessorCount, 8));

            var root = new RootCommand("PowerPoint Structure Inspector (parallel, SDK-per-slide)");
            root.AddOption(inputOpt);
            root.AddOption(outOpt);
            root.AddOption(mediaDirOpt);
            root.AddOption(embedImagesOpt);
            root.AddOption(maxDegreeOpt);

            root.SetHandler((FileInfo? input, FileInfo? output, DirectoryInfo? mediaDir, bool embedImages, int maxDegree) =>
            {
                if (input is null || !input.Exists) throw new FileNotFoundException("Input .pptx not found", input?.FullName);
                if (output is null) throw new ArgumentNullException("--out");
                if (mediaDir is null) throw new ArgumentNullException("--media-dir");
                if (maxDegree <= 0) maxDegree = 1;

                var deck = Inspector.Run(input.FullName, mediaDir.FullName, embedImages, maxDegree);
                var json = JsonSerializer.Serialize(deck, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                });
                File.WriteAllText(output.FullName, json);
                Console.WriteLine($"Wrote {output.FullName}");
            },
            inputOpt, outOpt, mediaDirOpt, embedImagesOpt, maxDegreeOpt);

            return root.Invoke(args);
        }
    }

    public static class Inspector
    {
        public static DeckReport Run(string pptxPath, string mediaDir, bool embedImages, int maxDegree)
        {
            Directory.CreateDirectory(mediaDir);

            // Stage 0: enumerate slides and prebuild layout placeholder maps (shared, immutable)
            List<SlideTask> tasks;
            var layoutCache = new LayoutCache();
            using (var doc = PresentationDocument.Open(pptxPath, false))
            {
                var pres = doc.PresentationPart ?? throw new InvalidOperationException("Invalid deck.");
                var slideIds = pres.Presentation!.SlideIdList!.Elements<SlideId>().ToList();
                tasks = new List<SlideTask>(slideIds.Count);

                // Build stable ordered list of slide URIs (e.g., /ppt/slides/slide1.xml)
                for (int i = 0; i < slideIds.Count; i++)
                {
                    var sid = slideIds[i];
                    var sp = (SlidePart)pres.GetPartById(sid.RelationshipId!);
                    tasks.Add(new SlideTask { Index = i + 1, SlideUri = sp.Uri.ToString() });
                }

                // Pre-compute placeholder maps per layout URI to share across workers
                foreach (var sm in pres.SlideMasterParts)
                {
                    foreach (var lp in sm.SlideLayoutParts)
                    {
                        var map = PlaceholderUtils.BuildLayoutPlaceholderMap(lp.SlideLayout);
                        layoutCache.TryAdd(lp.Uri.ToString(), map);
                    }
                }
            }

            // Shared image dedupe: hash -> relative file path
            var imageIndex = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Pre-size results
            var results = new SlideReport[tasks.Count];
            var po = new ParallelOptions { MaxDegreeOfParallelism = maxDegree };

            Parallel.ForEach(tasks, po, task =>
            {
                try
                {
                    using var doc = PresentationDocument.Open(pptxPath, false);
                    var pres = doc.PresentationPart!;
                    var sp = pres.SlideParts.FirstOrDefault(p => p.Uri.ToString().Equals(task.SlideUri, StringComparison.OrdinalIgnoreCase));
                    if (sp == null) throw new InvalidOperationException($"Slide not found by URI: {task.SlideUri}");

                    var report = Reader.ReadOneSlide(sp, layoutCache, mediaDir, embedImages, task.Index, imageIndex);
                    results[task.Index - 1] = report;
                }
                catch (Exception ex)
                {
                    results[task.Index - 1] = new SlideReport
                    {
                        SlideNumber = task.Index,
                        Title = null,
                        Placeholders = new(),
                        OtherShapes = new(),
                        Error = ex.Message
                    };
                }
            });

            return new DeckReport
            {
                File = Path.GetFileName(pptxPath),
                SlideCount = results.Length,
                Slides = new List<SlideReport>(results)
            };
        }
    }

    internal class SlideTask
    {
        public int Index { get; set; }
        public string SlideUri { get; set; } = "";
    }

    // ---------------- Reader logic per slide ----------------
    public static class Reader
    {
        public static SlideReport ReadOneSlide(SlidePart sp, LayoutCache layoutCache, string mediaDir, bool embedImages, int slideIndex, ConcurrentDictionary<string, string> imageIndex)
        {
            var layoutUri = sp.SlideLayoutPart?.Uri.ToString();
            var layoutMap = (layoutUri != null && layoutCache.TryGet(layoutUri, out var m)) ? m : new Dictionary<int, string>();

            var slideInfo = new SlideReport { SlideNumber = slideIndex, Placeholders = new(), OtherShapes = new() };
            var tree = sp.Slide.CommonSlideData!.ShapeTree!;

            foreach (var child in tree.ChildElements)
                WalkShapeElement(child, slideInfo, sp, layoutMap, mediaDir, embedImages, slideIndex, imageIndex);

            slideInfo.Title = slideInfo.Placeholders
                .Where(p => EqualsCI(p.PlaceholderType, "Title") || EqualsCI(p.PlaceholderType, "CtrTitle"))
                .Select(p => p.Text).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));

            return slideInfo;
        }

        static void WalkShapeElement(DocumentFormat.OpenXml.OpenXmlElement elem, SlideReport slideInfo, SlidePart sp,
                                     Dictionary<int, string> layoutMap, string mediaDir, bool embedImages, int slideIndex,
                                     ConcurrentDictionary<string, string> imageIndex)
        {
            switch (elem)
            {
                case Shape shp:
                    HandleShape(shp, slideInfo, layoutMap);
                    break;
                case GraphicFrame gf:
                    HandleGraphicFrame(gf, slideInfo, sp, layoutMap);
                    break;
                case Picture pic:
                    HandlePicture(pic, slideInfo, sp, mediaDir, embedImages, slideIndex, imageIndex);
                    break;
                case GroupShape grp:
                    foreach (var child in grp.ChildElements)
                        WalkShapeElement(child, slideInfo, sp, layoutMap, mediaDir, embedImages, slideIndex, imageIndex);
                    break;
                default:
                    break;
            }
        }

        static void HandleShape(Shape shp, SlideReport slideInfo, Dictionary<int, string> layoutMap)
        {
            var ph = shp.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .GetFirstChild<PlaceholderShape>();
            var text = ExtractText(shp.TextBody);
            var bounds = GetBounds(shp.Transform2D);

            if (ph != null)
            {
                var (idx, type) = PlaceholderType(ph, layoutMap);
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

        static void HandleGraphicFrame(GraphicFrame gf, SlideReport slideInfo, SlidePart sp, Dictionary<int, string> layoutMap)
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
                    var (idx, type) = PlaceholderType(ph, layoutMap);
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
                    var cpart = (ChartPart)sp.GetPartById(relId);
                    kind = DetectChartKind(cpart);
                }
                catch { }

                var cinfo = new ChartInfo { ChartType = kind };

                if (ph != null)
                {
                    var (idx, type) = PlaceholderType(ph, layoutMap);
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
                var (idx2, type2) = PlaceholderType(ph, layoutMap);
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

        static void HandlePicture(Picture pic, SlideReport slideInfo, SlidePart sp, string mediaDir, bool embedImages, int slideIndex,
                                  ConcurrentDictionary<string, string> imageIndex)
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
                    var part = (ImagePart)sp.GetPartById(relId);
                    using var s = part.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    var bytes = ms.ToArray();

                    // Hash for dedupe
                    string hashHex;
                    using (var sha1 = SHA1.Create())
                    {
                        var hash = sha1.ComputeHash(bytes);
                        hashHex = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                    }

                    var ext = ExtensionFromContentType(part.ContentType);
                    var relPath = $"media/{hashHex}.{ext}";
                    var outPath = Path.Combine(mediaDir, $"{hashHex}.{ext}");

                    // Deduped write
                    if (imageIndex.TryAdd(hashHex, relPath))
                    {
                        Directory.CreateDirectory(mediaDir);
                        try { File.WriteAllBytes(outPath, bytes); } catch { /* ignore write races */ }
                    }

                    iinfo = new ImageInfo
                    {
                        ContentType = part.ContentType,
                        Bytes = embedImages ? Convert.ToBase64String(bytes) : null,
                        File = relPath.Replace('\\','/'),
                        SizeBytes = bytes.Length
                    };
                }
                catch { /* ignore */ }
            }

            if (ph != null)
            {
                var (idx, type) = PlaceholderType(ph, new Dictionary<int, string>());
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

        static bool EqualsCI(string? a, string? b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    }

    // ---------------- LayoutCache (shared) ----------------
    public class LayoutCache
    {
        private readonly ConcurrentDictionary<string, Dictionary<int, string>> _map = new(StringComparer.OrdinalIgnoreCase);

        public bool TryAdd(string layoutUri, Dictionary<int, string> placeholderMap) => _map.TryAdd(layoutUri, placeholderMap);
        public bool TryGet(string layoutUri, out Dictionary<int, string> placeholderMap) => _map.TryGetValue(layoutUri, out placeholderMap!);
    }

    public static class PlaceholderUtils
    {
        public static Dictionary<int, string> BuildLayoutPlaceholderMap(SlideLayout? layout)
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
        public string? Error { get; set; }
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
}
