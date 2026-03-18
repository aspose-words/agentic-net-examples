using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReflectionOptimizationBenchmark
{
    static void Main()
    {
        // Path to save the generated documents (optional, can be omitted if not needed)
        const string outputPathEnabled = "Report_ReflectionOptimized.docx";
        const string outputPathDisabled = "Report_ReflectionNotOptimized.docx";

        // Generate a large JSON array with 50,000 items.
        string json = GenerateLargeJsonArray(50000);

        // Create a simple template document that repeats a field for each item.
        Document template = CreateTemplateDocument();

        // Benchmark with reflection optimization enabled (default = true).
        ReportingEngine.UseReflectionOptimization = true;
        long timeEnabled = BuildReportAndMeasure(template, json, outputPathEnabled);

        // Benchmark with reflection optimization disabled.
        ReportingEngine.UseReflectionOptimization = false;
        long timeDisabled = BuildReportAndMeasure(template, json, outputPathDisabled);

        // Output the measured times.
        Console.WriteLine($"Reflection optimization enabled : {timeEnabled} ms");
        Console.WriteLine($"Reflection optimization disabled: {timeDisabled} ms");
    }

    // Generates a JSON string representing an array of objects with a single "name" property.
    private static string GenerateLargeJsonArray(int itemCount)
    {
        var sb = new StringBuilder();
        sb.Append('[');
        for (int i = 0; i < itemCount; i++)
        {
            sb.Append("{\"name\":\"Item ").Append(i).Append("\"}");
            if (i < itemCount - 1)
                sb.Append(',');
        }
        sb.Append(']');
        return sb.ToString();
    }

    // Creates a minimal template document that uses Reporting Engine syntax.
    // The placeholder "<<persons.name>>" will be repeated for each array element.
    private static Document CreateTemplateDocument()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        // Insert a table with a single cell containing the placeholder.
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
        // Load JSON data from a memory stream to avoid file I/O overhead.
        using (var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonData)))
        {
            // Use default JsonDataLoadOptions.
            var dataSource = new JsonDataSource(jsonStream);

            // Create a fresh copy of the template for each run.
            var doc = (Document)template.Clone();

            // Build the report.
            var engine = new ReportingEngine();
            var stopwatch = Stopwatch.StartNew();
            engine.BuildReport(doc, dataSource, "persons");
            stopwatch.Stop();

            // Save the generated document (optional, can be omitted for pure benchmarking).
            doc.Save(outputPath);

            return stopwatch.ElapsedMilliseconds;
        }
    }
}
