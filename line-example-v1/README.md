# End-to-End Example: LINE Chart

This folder shows a **complete workflow** for a LINE chart using the canonical contract (Option 1).

Contents:
- `Chart.Line.schema.json` — JSON Schema for the LINE family
- `line_payload.json` — provider payload (time-scale X, min/max & major/minor on axes, markers)
- `LineExample.cs` — small entry point that validates & builds the chart with `PptChartKit.cs`

## Steps
1. Add **PptChartKit.cs** to your VSTO add-in project. (If you don't have it yet, grab it from the main bundle.)
2. Add **NJsonSchema** via NuGet.
3. Place `Chart.Line.schema.json` in your project (or use the unified `Chart.All.schema.json`).
4. Call:
   ```csharp
   LineExample.InsertLineFromFile(activeSlide, "line_payload.json", "Chart.Line.schema.json");
   ```

That’s it. The chart will be rendered at (x=72, y=72, w=560, h=320) with your axis bounds and markers.
