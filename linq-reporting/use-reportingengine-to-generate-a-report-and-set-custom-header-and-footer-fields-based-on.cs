using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportModel
{
    public string Header { get; set; } = "";
    public string Footer { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public string Value { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonContent = @"{
            ""Header"": ""My Report Header"",
            ""Footer"": ""Page Footer - Confidential"",
            ""Items"": [
                { ""Name"": ""Item1"", ""Value"": ""100"" },
                { ""Name"": ""Item2"", ""Value"": ""200"" },
                { ""Name"": ""Item3"", ""Value"": ""300"" }
            ]
        }";

        string jsonPath = "data.json";
        File.WriteAllText(jsonPath, jsonContent);

        // Deserialize JSON into the model.
        string jsonData = File.ReadAllText(jsonPath);
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(jsonData)!;

        // Create the template document with header, footer, and body tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header with custom field.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("<<[model.Header]>>");

        // Footer with custom field.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("<<[model.Footer]>>");

        // Body content.
        builder.MoveToDocumentEnd();
        builder.Writeln("Report Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
