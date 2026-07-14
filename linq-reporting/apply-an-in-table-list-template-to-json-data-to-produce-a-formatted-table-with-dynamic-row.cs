using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the template, JSON data and the final report.
        string templatePath = "template.docx";
        string jsonPath = "data.json";
        string outputPath = "report.docx";

        // 1. Create sample JSON data (an array of objects).
        string jsonContent = @"[
            { ""Index"": 1, ""Name"": ""Apple"",  ""Quantity"": 10 },
            { ""Index"": 2, ""Name"": ""Banana"", ""Quantity"": 20 },
            { ""Index"": 3, ""Name"": ""Cherry"", ""Quantity"": 15 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product List");
        builder.Writeln("==============");
        builder.Writeln("<<foreach [item in items]>>");

        // Start a table with a header row.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data row – placeholders will be replaced by JSON values.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(reportDoc, jsonDataSource, "items");

        // 6. Save the generated report.
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
