using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for older .NET versions.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample JSON data.
        // -----------------------------------------------------------------
        string jsonPath = "sampleData.json";
        string jsonContent = @"[
  { ""Name"": ""Alice"", ""Age"": 30 },
  { ""Name"": ""Bob"",   ""Age"": 25 },
  { ""Name"": ""Carol"", ""Age"": 28 }
]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin the foreach loop – it will repeat the whole table for each item.
        builder.Writeln("<<foreach [item in data]>>");

        // Start the table.
        Table table = builder.StartTable();

        // Header row – column names.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row – values will be filled from JSON.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Age]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Create a JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "data");

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
