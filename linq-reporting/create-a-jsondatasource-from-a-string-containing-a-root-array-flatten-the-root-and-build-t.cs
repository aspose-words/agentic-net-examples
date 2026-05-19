using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample JSON string containing a root array of person objects.
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 40 }
        ]";

        // Convert the JSON string to a memory stream.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));

        // Create a JsonDataSource from the stream.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // -----------------------------------------------------------------
        // Step 1: Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Persons Report");
        builder.Writeln();

        // Insert LINQ Reporting tags.
        // The data source will be referenced by the name "persons".
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

        // Build the report using the JSON data source.
        // The third argument ("persons") is the name used in the template tags.
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
