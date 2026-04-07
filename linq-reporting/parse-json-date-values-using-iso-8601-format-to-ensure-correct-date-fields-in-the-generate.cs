using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for extended encodings if needed.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string jsonPath = Path.Combine(outputDir, "order.json");
        string resultPath = Path.Combine(outputDir, "Report.docx");

        // Create sample JSON with ISO 8601 date values.
        string jsonContent = @"{
  ""OrderDate"": ""2023-08-15T14:30:00Z"",
  ""CustomerName"": ""John Doe"",
  ""Items"": [
    { ""Name"": ""Apple"", ""Quantity"": 3 },
    { ""Name"": ""Banana"", ""Quantity"": 5 }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent);

        // Build the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Order Report");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Table row for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the saved template.
        var reportDoc = new Document(templatePath);

        // Configure JSON data source with ISO 8601 parsing formats.
        var jsonOptions = new JsonDataLoadOptions
        {
            ExactDateTimeParseFormats = new List<string>
            {
                "yyyy-MM-ddTHH:mm:ssZ",
                "yyyy-MM-ddTHH:mm:ss"
            }
        };
        var jsonDataSource = new JsonDataSource(jsonPath, jsonOptions);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "order");

        // Save the generated report.
        reportDoc.Save(resultPath);
    }
}
