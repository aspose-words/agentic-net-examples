using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;          // ReportingEngine, JsonDataSource, ReportBuildOptions
using Aspose.Words.Tables;             // Table class for building tables

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample JSON data (array of objects) and write it to a file.
        // -----------------------------------------------------------------
        string jsonPath = "data.json";
        string jsonContent = @"[
            { ""Id"": 1, ""Name"": ""Apple"",  ""Quantity"": 10 },
            { ""Id"": 2, ""Name"": ""Banana"", ""Quantity"": 20 },
            { ""Id"": 3, ""Name"": ""Cherry"", ""Quantity"": 15 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // -----------------------------------------------------------------
        // 2. Create a Word template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the root collection named "items".
        builder.Writeln("<<foreach [item in items]>>");

        // Build a table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Table row bound to the current item.
        builder.InsertCell();
        builder.Writeln("<<[item.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();

        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using the JSON data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Create a JsonDataSource from the JSON file.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report. The root name is "items".
        bool success = engine.BuildReport(reportDoc, jsonDataSource, "items");

        // Optional: check the success flag (relevant when InlineErrorMessages is set).
        if (!success)
        {
            Console.WriteLine("Report generation encountered errors.");
        }

        // Save the generated report.
        reportDoc.Save("output.docx");
    }
}
