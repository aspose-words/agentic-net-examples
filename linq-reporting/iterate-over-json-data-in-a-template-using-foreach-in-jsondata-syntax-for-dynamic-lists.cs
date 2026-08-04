using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample JSON data file.
        string jsonFile = Path.Combine(Directory.GetCurrentDirectory(), "people.json");
        File.WriteAllText(jsonFile,
            @"[
  { ""Name"": ""Alice"",   ""Age"": 30 },
  { ""Name"": ""Bob"",     ""Age"": 25 },
  { ""Name"": ""Charlie"", ""Age"": 28 }
]");

        // Build a template document containing LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("People List:");
        builder.Writeln("<<foreach [in jsonData]>>");
        builder.Writeln("Name: <<[Name]>>, Age: <<[Age]>>");
        builder.Writeln("<</foreach>>");

        // Load the JSON data source.
        var jsonDataSource = new JsonDataSource(jsonFile);

        // Generate the report by merging the template with the JSON data.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "jsonData");

        // Save the resulting document.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputFile);
    }
}
