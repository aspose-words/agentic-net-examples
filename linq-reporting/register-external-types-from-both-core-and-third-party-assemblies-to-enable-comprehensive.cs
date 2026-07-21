using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Model used for JSON deserialization.
    public class Msg
    {
        public string Message { get; set; } = string.Empty;
    }

    // Helper class that encapsulates JSON deserialization logic.
    // This avoids calling generic methods directly from the template.
    public static class JsonHelper
    {
        public static string GetMessage()
        {
            // Sample JSON payload.
            const string json = "{\"Message\":\"Hello from JSON!\"}";
            // Deserialize using Newtonsoft.Json and return the Message property.
            Msg? obj = JsonConvert.DeserializeObject<Msg>(json);
            return obj?.Message ?? string.Empty;
        }
    }

    public static void Main()
    {
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Tag that uses a core .NET type (System.Math).
        builder.Writeln("Value of Math.PI: <<[Math.PI]>>");

        // Tag that uses a third‑party helper to obtain a JSON message.
        builder.Writeln("Deserialized JSON message: <<[JsonHelper.GetMessage()]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document for reporting.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Register core and third‑party types so their static members can be accessed.
        engine.KnownTypes.Add(typeof(Math));          // Core .NET type.
        engine.KnownTypes.Add(typeof(JsonHelper));   // Wrapper for third‑party JSON logic.

        // No data source is required because we only use static members.
        object dummyDataSource = new object();

        // Build the report.
        engine.BuildReport(doc, dummyDataSource, "");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}
