using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for environments that require it.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple JSON file representing an array of product objects.
        string jsonContent = @"[
            { ""Name"": ""Apple"",  ""Quantity"": 5 },
            { ""Name"": ""Banana"", ""Quantity"": 3 },
            { ""Name"": ""Cherry"", ""Quantity"": 12 }
        ]";
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, jsonContent);

        // Build a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Product List:");
        builder.Writeln("<<foreach [p in items]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Load the JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, jsonDataSource, "items");

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
