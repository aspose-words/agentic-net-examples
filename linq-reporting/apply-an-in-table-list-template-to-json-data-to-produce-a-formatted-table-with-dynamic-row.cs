using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for Table type

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample JSON data.
        string jsonPath = Path.Combine(outputDir, "data.json");
        string jsonContent = @"[
  { ""Index"": 1, ""Name"": ""Apple"",  ""Quantity"": 10 },
  { ""Index"": 2, ""Name"": ""Banana"", ""Quantity"": 20 },
  { ""Index"": 3, ""Name"": ""Cherry"", ""Quantity"": 15 }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Product List");
        builder.Writeln();

        // Open foreach block.
        builder.Writeln("<<foreach [item in data]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data row – each cell contains a tag that references the current item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();

        // Close table and foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // 4. Create a JSON data source.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;   // Set options explicitly.
        engine.BuildReport(reportDoc, jsonData, "data");

        // 6. Save the final report.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(resultPath);
    }
}
