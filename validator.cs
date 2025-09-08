using NJsonSchema;

// Load the aggregator (it $ref's the family schemas in the same folder)
var schema = await JsonSchema.FromFileAsync("Chart.All.schema.json");
var json = File.ReadAllText("some-chart.json");
var errors = schema.Validate(json);
if (errors.Count > 0)
    throw new InvalidOperationException(string.Join("\n", errors.Select(e => e.ToString())));
