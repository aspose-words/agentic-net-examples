using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        const string jsonPath = "data.json";
        string jsonContent = @"{
            ""Header"": ""Custom Header Text"",
            ""Footer"": ""Custom Footer Text"",
            ""Title"": ""Sample Report"",
            ""Items"": [
                { ""Name"": ""Item 1"" },
                { ""Name"": ""Item 2"" },
                { ""Name"": ""Item 3"" }
            ]
        }";
        File.WriteAllText(jsonPath, jsonContent);

        // Create a template document programmatically.
        const string templatePath = "template.docx";
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert header tag.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("<<[model.Header]>>");

        // Insert footer tag.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("<<[model.Footer]>>");

        // Return to the main body.
        builder.MoveToDocumentEnd();

        // Insert title and a simple items list using a foreach block.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Load JSON data source.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonData, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
