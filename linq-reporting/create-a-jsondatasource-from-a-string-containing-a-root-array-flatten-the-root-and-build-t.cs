using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data (root is an array of objects).
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Carol"", ""Age"": 28 }
        ]";

        // Write JSON to a memory stream and reset its position.
        using var jsonStream = new MemoryStream();
        using (var writer = new StreamWriter(jsonStream, leaveOpen: true))
        {
            writer.Write(json);
            writer.Flush();
            jsonStream.Position = 0;
        }

        // Create a JsonDataSource from the stream.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // -----------------------------------------------------------------
        // Step 1: Create the template document programmatically.
        // -----------------------------------------------------------------
        var templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the JSON data source. The root name is "persons".
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the generated report.
        var outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
