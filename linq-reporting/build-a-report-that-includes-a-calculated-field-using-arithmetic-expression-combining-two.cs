using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "ReportTemplate.docx");
        string jsonPath = Path.Combine(workDir, "Data.json");
        string outputPath = Path.Combine(workDir, "ReportResult.docx");

        // -----------------------------------------------------------------
        // 1. Create a JSON file that will serve as the data source.
        // -----------------------------------------------------------------
        string jsonContent = @"[
  { ""Name"": ""Item A"", ""Value1"": 10, ""Value2"": 5 },
  { ""Name"": ""Item B"", ""Value1"": 7,  ""Value2"": 3 },
  { ""Name"": ""Item C"", ""Value1"": 12, ""Value2"": 8 }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        //    The template contains LINQ Reporting tags, including a calculated
        //    field that adds Value1 and Value2.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report of Items");
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Value1: <<[item.Value1]>>");
        builder.Writeln("Value2: <<[item.Value2]>>");
        // Calculated field: sum of the two numeric properties.
        builder.Writeln("Sum (Value1 + Value2): <<[item.Value1 + item.Value2]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the JSON data source.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var jsonDataSource = new JsonDataSource(jsonPath);

        // -----------------------------------------------------------------
        // 4. Build the final report.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        // The root object name used in the template is "items".
        engine.BuildReport(loadedTemplate, jsonDataSource, "items");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath);
    }
}
