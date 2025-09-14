// File: Program.cs
// Requires: dotnet add package DocumentFormat.OpenXml
using System;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OoxmlChartCopyDemo
{
    class Program
    {
        // 1 inch = 914400 EMUs
        const long EMU = 914400;

        static int Main(string[] args)
        {
            try
            {
                string ppt1 = Path.GetFullPath("PPT1_WithChart.pptx");
                string ppt2 = Path.GetFullPath("PPT2_Target.pptx");

                Console.WriteLine("Step 1: Create ChartML (Pie, cache-only, no workbook)...");
                string chartXml = SampleChartML.PieCacheOnly;

                Console.WriteLine("Step 2: Create PPT #1 and insert the chart...");
                CreateNewPresentation(ppt1);
                using (var doc = PresentationDocument.Open(ppt1, true))
                {
                    var presPart = doc.PresentationPart!;
                    var slide = GetOrCreateFirstSlide(presPart);

                    // Add chart to the first slide (position ~1” margin; size ~7.5” x 4.5”)
                    AddChartToSlide(slide, chartXml, 1 * EMU, 1 * EMU, 7_500_000, 4_500_000);
                }

                Console.WriteLine("Step 3: Create PPT #2 (empty)...");
                CreateNewPresentation(ppt2);

                Console.WriteLine("Step 4: Copy the chart from PPT #1 to PPT #2 (pure Open XML)...");
                CopyFirstChartFromSourceToDest(ppt1, ppt2, 1 * EMU, 1 * EMU, 7_500_000, 4_500_000);

                Console.WriteLine("Done.");
                Console.WriteLine($"Source: {ppt1}");
                Console.WriteLine($"Target: {ppt2}");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.ToString());
                return 1;
            }
        }

        // ---------------------------
        // PRESENTATION BASICS
        // ---------------------------

        /// <summary>Create a minimal .pptx with a SlideMaster + one blank Slide.</summary>
        static void CreateNewPresentation(string path)
        {
            if (File.Exists(path)) File.Delete(path);

            using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
            var presPart = doc.AddPresentationPart();
            presPart.Presentation = new Presentation();

            // Master + Layout
            var master = presPart.AddNewPart<SlideMasterPart>();
            master.SlideMaster = new SlideMaster(new CommonSlideData(new ShapeTree()), new SlideLayoutIdList());

            var layout = master.AddNewPart<SlideLayoutPart>();
            layout.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree()));

            // Hook layout into master
            var layoutId = new SlideLayoutId { Id = 1U, RelationshipId = master.GetIdOfPart(layout) };
            master.SlideMaster.SlideLayoutIdList!.Append(layoutId);

            // Register master in presentation
            presPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                new SlideMasterId { Id = 1U, RelationshipId = presPart.GetIdOfPart(master) });

            // Create first slide using that layout
            var slide = presPart.AddNewPart<SlidePart>();
            slide.AddPart(layout);
            slide.Slide = new Slide(new CommonSlideData(new ShapeTree()));

            // Add slide to the deck
            var slideIdList = presPart.Presentation.AppendChild(new SlideIdList());
            slideIdList.AppendChild(new SlideId { Id = 256U, RelationshipId = presPart.GetIdOfPart(slide) });

            slide.Slide.Save();
            presPart.Presentation.Save();
        }

        static SlidePart GetOrCreateFirstSlide(PresentationPart presPart)
        {
            var firstId = presPart.Presentation!.SlideIdList!.Elements<SlideId>().FirstOrDefault();
            if (firstId != null) return (SlidePart)presPart.GetPartById(firstId.RelationshipId!);

            // If strangely absent, create one
            var master = presPart.SlideMasterParts.First();
            var layout = master.SlideLayoutParts.First();
            var slide = presPart.AddNewPart<SlidePart>();
            slide.AddPart(layout);
            slide.Slide = new Slide(new CommonSlideData(new ShapeTree()));

            var newId = new SlideId { Id = 256U, RelationshipId = presPart.GetIdOfPart(slide) };
            if (presPart.Presentation.SlideIdList == null)
                presPart.Presentation.SlideIdList = new SlideIdList();
            presPart.Presentation.SlideIdList.Append(newId);

            slide.Slide.Save();
            presPart.Presentation.Save();
            return slide;
        }

        static SlidePart AddNewSlide(PresentationPart presPart)
        {
            var master = presPart.SlideMasterParts.First();
            var layout = master.SlideLayoutParts.First();
            var slide = presPart.AddNewPart<SlidePart>();
            slide.AddPart(layout);
            slide.Slide = new Slide(new CommonSlideData(new ShapeTree()));

            // Create a unique SlideId
            uint nextId = 256U;
            var list = presPart.Presentation!.SlideIdList ?? presPart.Presentation.AppendChild(new SlideIdList());
            var existing = list.Elements<SlideId>().Select(s => s.Id!.Value).DefaultIfEmpty(256U).Max();
            nextId = existing + 1;

            list.AppendChild(new SlideId { Id = nextId, RelationshipId = presPart.GetIdOfPart(slide) });
            slide.Slide.Save();
            presPart.Presentation.Save();
            return slide;
        }

        // ---------------------------
        // CHART INSERTION
        // ---------------------------

        /// <summary>Adds a ChartPart to a slide and places a graphicFrame that references it.</summary>
        static void AddChartToSlide(SlidePart slide, string chartXml, long x, long y, long cx, long cy)
        {
            // 1) Add chart part
            var chartRelId = "rIdChart" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var chartPart = slide.AddNewPart<ChartPart>(chartRelId);
            chartPart.ChartSpace = LoadChartSpace(chartXml);

            // 2) Insert a graphicFrame referencing that part
            var tree = slide.Slide!.CommonSlideData!.ShapeTree!;
            uint shapeId = GetNextShapeId(tree);

            var gf = new GraphicFrame(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = "Chart " + shapeId },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new Transform(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.Graphic(
                    new A.GraphicData(
                        new C.Chart { Id = chartRelId }
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                )
            );

            tree.Append(gf);
            slide.Slide.Save();
        }

        static uint GetNextShapeId(ShapeTree tree)
        {
            var ids = tree.Descendants<NonVisualDrawingProperties>()
                          .Select(p => p.Id?.Value ?? 0U);
            return (ids.Any() ? ids.Max() + 1 : 1U);
        }

        static C.ChartSpace LoadChartSpace(string xml)
        {
            using var sr = new StringReader(xml);
            using var xr = System.Xml.XmlReader.Create(sr);
            return C.ChartSpace.Load(xr);
        }

        // ---------------------------
        // COPY CHART FROM PPT1 → PPT2
        // ---------------------------

        /// <summary>
        /// Finds the first ChartPart in PPT1, reads its ChartML, then creates a new slide in PPT2 and
        /// rehydrates that chart there (cache-only: no workbook required).
        /// </summary>
        static void CopyFirstChartFromSourceToDest(string sourcePptx, string destPptx, long x, long y, long cx, long cy)
        {
            string? chartXml = null;

            using (var src = PresentationDocument.Open(sourcePptx, false))
            {
                var sp = src.PresentationPart!.SlideParts.FirstOrDefault();
                if (sp == null) throw new InvalidOperationException("Source has no slides.");
                var cpart = sp.ChartParts.FirstOrDefault();
                if (cpart == null) throw new InvalidOperationException("No charts found in source.");

                chartXml = cpart.ChartSpace.OuterXml;
            }

            using (var dst = PresentationDocument.Open(destPptx, true))
            {
                var presPart = dst.PresentationPart!;
                var newSlide = AddNewSlide(presPart);
                AddChartToSlide(newSlide, chartXml!, x, y, cx, cy);
            }
        }
    }

    // ---------------------------
    // SAMPLE CHARTML (Pie, no workbook)
    // ---------------------------
    static class SampleChartML
    {
        public static readonly string PieCacheOnly =
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""
              xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:varyColors val=""1""/>
        <c:ser>
          <c:idx val=""0""/>
          <c:order val=""0""/>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:ptCount val=""1""/>
                <c:pt idx=""0""><c:v>Market Share</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:dLbls><c:showPercent val=""1""/></c:dLbls>

          <!-- Categories -->
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:ptCount val=""3""/>
                <c:pt idx=""0""><c:v>A</c:v></c:pt>
                <c:pt idx=""1""><c:v>B</c:v></c:pt>
                <c:pt idx=""2""><c:v>C</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>

          <!-- Values -->
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val=""3""/>
                <c:pt idx=""0""><c:v>10</c:v></c:pt>
                <c:pt idx=""1""><c:v>20</c:v></c:pt>
                <c:pt idx=""2""><c:v>30</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:firstSliceAng val=""0""/>
      </c:pieChart>
    </c:plotArea>
    <c:plotVisOnly val=""1""/>
  </c:chart>
</c:chartSpace>";
    }
}
