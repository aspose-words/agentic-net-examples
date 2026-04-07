using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json.Linq;

public class ReportModel
{
    // Empty model; all data accessed via external types.
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a tag that accesses a core .NET type (System.Math).
        builder.Writeln("Value of PI from System.Math: <<[Math.PI]>>");

        // Insert a tag that accesses a third‑party type (Newtonsoft.Json.Linq.JObject).
        // The expression parses a JSON string and extracts the \"value\" property.
        builder.Writeln(
            "Value from Newtonsoft.Json JObject: " +
            "<<[JObject.Parse(\"{\\\"value\\\":123}\")[\"value\"]]>>");

        // Create a simple root data model (required by BuildReport overload with name).
        ReportModel model = new ReportModel();

        // Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Register external types so that the template can use their static members.
        engine.KnownTypes.Add(typeof(Math));               // Core .NET type.
        engine.KnownTypes.Add(typeof(JObject));            // Third‑party type from Newtonsoft.Json.

        // Build the report. The template does not reference the model directly,
        // but we still provide it to satisfy the overload that includes a data source name.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated at: {outputPath}");
    }
}
