using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var items = new List<Item>
        {
            new Item { Title = "First Item", BookmarkName = "FirstBookmark" },
            new Item { Title = "Second Item", BookmarkName = "" }, // Empty name – bookmark will be skipped.
            new Item { Title = "Third Item", BookmarkName = "ThirdBookmark" }
        };

        // Create the template document programmatically.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting.
        var engine = new ReportingEngine
        {
            // Remove paragraphs that become empty after the conditional bookmark is omitted.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Wrap the data source in a public class (required by the engine).
        var model = new ReportModel { Items = items };

        // Build the report. No data source name is needed because the template references members directly.
        engine.BuildReport(doc, model);

        // Save the generated report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }

    // Creates a Word template containing a foreach loop with a conditional bookmark.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Write the title for each item.
        builder.Writeln("Title: <<[item.Title]>>");

        // Conditional bookmark – created only when BookmarkName is not empty.
        builder.Writeln("<<if [item.BookmarkName != \"\"]>>");
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("Bookmark content for <<[item.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Simple data model for each item.
    public class Item
    {
        public string Title { get; set; } = string.Empty;
        public string BookmarkName { get; set; } = string.Empty;
    }

    // Wrapper class required by ReportingEngine (non‑anonymous).
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }
}
