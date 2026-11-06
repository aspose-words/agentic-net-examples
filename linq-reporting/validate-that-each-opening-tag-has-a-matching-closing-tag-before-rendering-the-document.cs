using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Load the template for validation and reporting.
        Document templateDoc = new Document(templatePath);

        // Validate that each opening tag has a matching closing tag.
        if (!ValidateReportingTags(templateDoc))
        {
            Console.WriteLine("Tag validation failed. Report generation aborted.");
            return;
        }

        // Prepare sample data.
        ReportModel model = new()
        {
            Title = "Sample Report",
            Items = new()
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        templateDoc.Save(reportPath);
        Console.WriteLine($"Report generated successfully at: {reportPath}");
    }

    // Creates a Word document containing LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Title
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header (once per document, but placed inside loop for simplicity).
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Table row for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Bookmark example with matching closing tag.
        builder.Writeln("<<bookmark [\"MyBookmark\"]>>");
        builder.Writeln("This is bookmarked text.");
        builder.Writeln("<</bookmark>>");

        doc.Save(path);
    }

    // Validates that opening tags have matching closing tags in the document.
    private static bool ValidateReportingTags(Document doc)
    {
        string docText = doc.GetText();

        // Tags that require a closing counterpart.
        string[] pairedTags = { "foreach", "if", "bookmark", "textColor", "backColor" };

        bool isValid = true;

        foreach (string tag in pairedTags)
        {
            // Opening tags: <<tag ...>>
            string openingPattern = $@"<<\s*{tag}\b[^>]*>>";
            // Closing tags: <</tag>>
            string closingPattern = $@"<</\s*{tag}\s*>>";

            int openingCount = Regex.Matches(docText, openingPattern).Count;
            int closingCount = Regex.Matches(docText, closingPattern).Count;

            if (openingCount != closingCount)
            {
                Console.WriteLine($"Tag mismatch for '{tag}': {openingCount} opening tag(s) vs {closingCount} closing tag(s).");
                isValid = false;
            }
        }

        return isValid;
    }
}

// Sample data model for the report.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Item used inside the foreach loop.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
