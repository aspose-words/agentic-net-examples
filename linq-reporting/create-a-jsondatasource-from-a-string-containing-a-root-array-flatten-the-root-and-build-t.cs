using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample JSON string with a root array.
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Carol"", ""Age"": 28 }
        ]";

        // Convert JSON string to a memory stream.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));

        // Create a JsonDataSource from the stream.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // -----------------------------------------------------------------
        // Create a template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template back for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using the JSON data source.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the generated report.
        const string outputPath = "ReportOutput.docx";
        reportDoc.Save(outputPath);
    }
}
