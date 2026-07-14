using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample JSON string with a root array of person objects.
        string json = @"[
            { ""Name"": ""John Doe"", ""Age"": 30 },
            { ""Name"": ""Jane Smith"", ""Age"": 25 }
        ]";

        // Write the JSON to a memory stream and reset its position.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        jsonStream.Position = 0;

        // Create a JsonDataSource from the stream.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

        // Create a new blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln("Persons Report:");
        builder.Writeln();

        // Insert a foreach loop that iterates over the root array (named "persons").
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
