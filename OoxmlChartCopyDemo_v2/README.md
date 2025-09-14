# OOXML Chart Copy Demo (No Workbook)

This console app:
1. Generates **Pie** ChartML (cache-only; no embedded workbook).
2. Creates **PPT1_WithChart.pptx** and inserts that chart.
3. Creates **PPT2_Target.pptx** and copies the chart from PPT1 into PPT2 using pure **Open XML** (no Interop).

## Prereqs
- .NET 8 SDK
- NuGet access to `DocumentFormat.OpenXml`

## Run
```bash
dotnet restore
dotnet run
```
Outputs in the working directory:
- `PPT1_WithChart.pptx`
- `PPT2_Target.pptx`
