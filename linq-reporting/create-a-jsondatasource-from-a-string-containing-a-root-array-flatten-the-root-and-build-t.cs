using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code pages provider for .NET Core.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // JSON string containing a root array of objects.
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"", ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 35 }
        ]";

        // Write the JSON string to a temporary file because JsonDataSource expects a file path.
        string jsonPath = "PeopleData.json";
        File.WriteAllText(jsonPath, json, Encoding.UTF8);

        // Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("People Report");
        builder.Writeln();

        // LINQ Reporting tags: iterate over the root collection and output fields.
        builder.Writeln("<<foreach [person in root]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template document (required before building the report).
        Document doc = new Document(templatePath);

        // Build the report using the JSON data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "root");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
