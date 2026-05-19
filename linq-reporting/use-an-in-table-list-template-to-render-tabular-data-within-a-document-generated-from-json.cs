using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template, JSON source and the generated report.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "Data.json");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create sample JSON data.
        // -----------------------------------------------------------------
        string jsonContent = @"{
  ""Items"": [
    { ""Index"": 1, ""Name"": ""Apple"" },
    { ""Index"": 2, ""Name"": ""Banana"" },
    { ""Index"": 3, ""Name"": ""Cherry"" }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Open the foreach block that iterates over the collection Items.
        builder.Writeln("<<foreach [item in model.Items]>>");

        // Create a table inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row – each iteration will fill these cells.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // Close the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using JSON data.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure JSON data source options to ensure the root object is generated.
        JsonDataLoadOptions jsonOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, jsonOptions);

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Use property assignment as required.

        // Build the report. The root object name used in the template is "model".
        engine.BuildReport(reportDoc, jsonDataSource, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
