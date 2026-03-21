using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

class DynamicPageBreakReport
{
    static void Main()
    {
        // Create a temporary JSON file with sample data.
        string jsonPath = Path.Combine(Path.GetTempPath(), "items.json");
        File.WriteAllText(jsonPath, @"[
            { ""Name"": ""Item 1"" },
            { ""Name"": ""Item 2"" },
            { ""Name"": ""Item 3"" }
        ]");

        // Create a simple template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<foreach [in items]>>");
        builder.Writeln("Item: <<[Name]>>");
        builder.Writeln("<</foreach>>");

        // Load JSON data as a data source for the reporting engine.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report by populating the template with the JSON data.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, jsonDataSource, "items");

        // Insert a page break after each item paragraph except the last one.
        DocumentBuilder reportBuilder = new DocumentBuilder(template);
        int paragraphCount = template.Sections[0].Body.Paragraphs.Count;
        for (int i = 0; i < paragraphCount - 1; i++)
        {
            reportBuilder.MoveToParagraph(i, 0);
            reportBuilder.InsertBreak(BreakType.PageBreak);
        }

        // Save the final document.
        template.Save("ReportWithDynamicPageBreaks.docx");

        // Clean up the temporary JSON file.
        File.Delete(jsonPath);
    }
}
