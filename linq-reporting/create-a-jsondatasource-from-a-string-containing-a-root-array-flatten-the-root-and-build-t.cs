using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample JSON containing a root array of person objects.
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 28 }
        ]";

        // Convert the JSON string to a memory stream.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        jsonStream.Position = 0; // Ensure the stream is at the beginning.

        // Create a JsonDataSource from the stream.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Build a simple template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tags: iterate over the root array (named "persons").
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Populate the template with data.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, jsonDataSource, "persons");

        // Optionally, you could check the success flag if InlineErrorMessages were enabled.
        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)} (Success: {success})");
    }
}
