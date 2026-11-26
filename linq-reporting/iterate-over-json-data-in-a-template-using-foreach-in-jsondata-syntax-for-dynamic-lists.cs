using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Create sample JSON data ----------
        var people = new[]
        {
            new { Name = "Alice", Age = 28 },
            new { Name = "Bob", Age = 35 },
            new { Name = "Charlie", Age = 22 }
        };
        string json = JsonConvert.SerializeObject(people, Formatting.Indented);
        // Save JSON to a temporary file (optional, can also use a stream directly)
        string jsonPath = "people.json";
        File.WriteAllText(jsonPath, json, Encoding.UTF8);

        // ---------- Create the template document ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("==============");
        // LINQ Reporting foreach tag iterating over the JSON root collection named "jsonData"
        builder.Writeln("<<foreach [person in jsonData]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // ---------- Load the template ----------
        var doc = new Document(templatePath);

        // ---------- Load JSON data source ----------
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        var jsonDataSource = new JsonDataSource(jsonStream);

        // ---------- Build the report ----------
        var engine = new ReportingEngine();
        // The root name used in the template tags is "jsonData"
        engine.BuildReport(doc, jsonDataSource, "jsonData");

        // ---------- Save the generated report ----------
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Optionally, indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
