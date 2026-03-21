using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;

class ReflectionOptimizationBenchmark
{
    static void Main()
    {
        const string outputPathEnabled = "Report_ReflectionOptimized.txt";
        const string outputPathDisabled = "Report_ReflectionNotOptimized.txt";

        // Generate a large JSON object with a "persons" array containing 50,000 items.
        string json = GenerateLargeJsonObject(50000);

        Document template = CreateTemplateDocument();

        ReportingEngine.UseReflectionOptimization = true;
        long timeEnabled = BuildReportAndMeasure(template, json, outputPathEnabled);

        ReportingEngine.UseReflectionOptimization = false;
        long timeDisabled = BuildReportAndMeasure(template, json, outputPathDisabled);

        Console.WriteLine($"Reflection optimization enabled : {timeEnabled} ms");
        Console.WriteLine($"Reflection optimization disabled: {timeDisabled} ms");
    }

    // Generates a JSON string representing an object with a "persons" array.
    private static string GenerateLargeJsonObject(int itemCount)
    {
        var sb = new StringBuilder();
        sb.Append("{\"persons\":[");
        for (int i = 0; i < itemCount; i++)
        {
            sb.Append("{\"name\":\"Item ").Append(i).Append("\"}");
            if (i < itemCount - 1)
                sb.Append(',');
        }
        sb.Append("]}");
        return sb.ToString();
    }

    // Creates a minimal template document that uses Reporting Engine syntax.
    private static Document CreateTemplateDocument()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("<<persons.name>>");
        builder.EndTable();
        return doc;
    }

    // Builds the report using the provided template and JSON data,
    // measures the elapsed time in milliseconds, and saves the result.
    private static long BuildReportAndMeasure(Document template, string jsonData, string outputPath)
    {
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonData));
        var dataSource = new JsonDataSource(jsonStream);
        var doc = (Document)template.Clone();

        var engine = new ReportingEngine();
        var stopwatch = Stopwatch.StartNew();
        engine.BuildReport(doc, dataSource, "persons");
        stopwatch.Stop();

        doc.Save(outputPath);
        return stopwatch.ElapsedMilliseconds;
    }
}

// Simple placeholder for Aspose.Words.Document
class Document
{
    private readonly StringBuilder _content = new StringBuilder();

    public void Append(string text) => _content.Append(text);
    public void AppendLine(string text) => _content.AppendLine(text);
    public override string ToString() => _content.ToString();

    public Document Clone()
    {
        var clone = new Document();
        clone._content.Append(this._content.ToString());
        return clone;
    }

    public void Save(string path)
    {
        File.WriteAllText(path, _content.ToString());
    }
}

// Simple placeholder for Aspose.Words.DocumentBuilder
class DocumentBuilder
{
    private readonly Document _doc;

    public DocumentBuilder(Document doc) => _doc = doc;

    public void StartTable() => _doc.AppendLine("[StartTable]");
    public void InsertCell() => _doc.AppendLine("[InsertCell]");
    public void Write(string text) => _doc.AppendLine(text);
    public void EndTable() => _doc.AppendLine("[EndTable]");
}

// Simple placeholder for Aspose.Words.Reporting.JsonDataSource
class JsonDataSource : IDisposable
{
    private readonly JsonDocument _doc;

    public JsonDataSource(Stream jsonStream)
    {
        _doc = JsonDocument.Parse(jsonStream);
    }

    public JsonElement GetRootArray(string name)
    {
        if (_doc.RootElement.TryGetProperty(name, out JsonElement array) && array.ValueKind == JsonValueKind.Array)
            return array;
        return default;
    }

    public void Dispose() => _doc.Dispose();
}

// Simple placeholder for Aspose.Words.Reporting.ReportingEngine
class ReportingEngine
{
    public static bool UseReflectionOptimization { get; set; }

    public void BuildReport(Document doc, JsonDataSource dataSource, string rootName)
    {
        var array = dataSource.GetRootArray(rootName);
        if (array.ValueKind != JsonValueKind.Array)
            return;

        foreach (var item in array.EnumerateArray())
        {
            // Simulate processing each item; the actual content is not used.
            // The presence of UseReflectionOptimization flag can affect how we access the property.
            string name;
            if (UseReflectionOptimization)
            {
                // Direct property access (simulated fast path)
                name = item.GetProperty("name").GetString();
            }
            else
            {
                // Simulate slower reflection-like access
                name = GetPropertyViaReflection(item, "name");
            }

            // Append the name to the document to mimic report generation.
            doc.AppendLine(name);
        }
    }

    private string GetPropertyViaReflection(JsonElement element, string propertyName)
    {
        // Simulate a slower lookup.
        foreach (var prop in element.EnumerateObject())
        {
            if (prop.NameEquals(propertyName))
                return prop.Value.GetString();
        }
        return string.Empty;
    }
}
