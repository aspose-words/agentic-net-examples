using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "people.json");
        string jsonContent = @"[
  { ""Name"": ""Alice"", ""Age"": 30 },
  { ""Name"": ""Bob"",   ""Age"": 25 },
  { ""Name"": ""Charlie"", ""Age"": 28 }
]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("People List:");
        // Use the foreach tag with the syntax <<foreach [in jsonData]>>.
        builder.Writeln("<<foreach [in jsonData]>>");
        // Inside the loop, reference fields of each JSON object.
        builder.Writeln("<<[Name]>> - <<[Age]>>");
        builder.Writeln("<</foreach>>");

        // Load the JSON data as a data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source name "jsonData".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "jsonData");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }
}
