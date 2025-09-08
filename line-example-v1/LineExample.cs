
// LineExample.cs
// Example usage: validate a LINE payload and render it onto the active slide.
// Requires: PptChartKit.cs in your project, NJsonSchema (NuGet), and Office interop references.

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NJsonSchema;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PptChartKit; // namespace from PptChartKit.cs

public static class LineExample
{
    private static JsonSchema? _schema;

    /// <summary>
    /// Inserts a LINE chart defined by a JSON file onto the given slide.
    /// </summary>
    public static void InsertLineFromFile(PowerPoint.Slide slide, string jsonPath, string schemaPath)
    {
        // Lazy load schema (sync for simplicity in VSTO)
        _schema ??= JsonSchema.FromFileAsync(schemaPath).GetAwaiter().GetResult();

        string json = File.ReadAllText(jsonPath);

        // Hook validation into the builder
        string? Validator(string raw)
        {
            var errors = _schema!.Validate(raw);
            if (errors.Count == 0) return null;
            return string.Join("\n", errors.Select(e => $"{e.Path}: {e.Kind}"));
        }

        // Build the chart
        ChartBuilder.BuildFromJson(slide, json, msg => System.Diagnostics.Debug.WriteLine(msg), Validator);
    }
}
