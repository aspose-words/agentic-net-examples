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
        // Path to the JSON file containing an array of items.
        // Example JSON: [{ "Name": "Item 1" }, { "Name": "Item 2" }, { "Name": "Item 3" }]
        string jsonPath = "items.json";

        // Create a simple template document.
        // The template uses a foreach loop to repeat the item name.
        // After each item we will later insert a page break programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<foreach [in items]>>");
        builder.Writeln("Item: <<[Name]>>");
        builder.Writeln("<</foreach>>");
        // Save the template for debugging (optional).
        template.Save("Template.docx");

        // Load JSON data as a data source for the reporting engine.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report by populating the template with the JSON data.
        ReportingEngine engine = new ReportingEngine();
        // The data source name "items" matches the name used in the foreach tag.
        engine.BuildReport(template, jsonDataSource, "items");

        // After the report is built, insert a page break after each item paragraph
        // (except after the last item).
        // The report currently contains a series of paragraphs: "Item: <name>"
        // We will walk through the paragraphs and insert a page break after each.
        for (int i = 0; i < template.Sections[0].Body.Paragraphs.Count - 1; i++)
        {
            // Move the builder to the end of the current paragraph.
            builder.MoveToParagraph(i, 0);
            // Insert a page break.
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the final document.
        template.Save("ReportWithDynamicPageBreaks.docx");
    }
}
