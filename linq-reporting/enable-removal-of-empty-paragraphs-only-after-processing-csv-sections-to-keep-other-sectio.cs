using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV reading.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV file.
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath,
            @"Name,Value
Alice,123
Bob,
,456
Charlie,789
,");

        // Create a Word template programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header section (static).
        builder.Writeln("=== Report Header ===");
        builder.Writeln();

        // Bookmark that will surround the CSV‑generated content.
        builder.Writeln("<<bookmark [\"CsvSection\"]>>");
        // CSV data section.
        builder.Writeln("<<foreach [row in CsvRows]>>");
        // Each row will be placed in its own paragraph.
        builder.Writeln("<<[row.Name]>> <<[row.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</bookmark>>");

        builder.Writeln();
        // Footer section (static).
        builder.Writeln("=== Report Footer ===");

        // Save the template (optional, for inspection).
        doc.Save("template.docx");

        // Load CSV data source.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
            // Default separator is ',', default quote is '"', default comment is '#'.
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "CsvRows");

        // Remove empty paragraphs that belong to the CSV section only.
        RemoveEmptyParagraphsInBookmark(doc, "CsvSection");

        // Save the final report.
        string outputPath = "output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated and saved to '{outputPath}'.");
    }

    private static void RemoveEmptyParagraphsInBookmark(Document doc, string bookmarkName)
    {
        // Verify the bookmark exists.
        if (!doc.Range.Bookmarks.Any(b => b.Name == bookmarkName))
            return;

        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
        Paragraph startParagraph = bookmark.BookmarkStart.GetAncestor(NodeType.Paragraph) as Paragraph;
        Paragraph endParagraph = bookmark.BookmarkEnd.GetAncestor(NodeType.Paragraph) as Paragraph;

        if (startParagraph == null || endParagraph == null)
            return;

        // Collect paragraphs between start and end (inclusive).
        List<Paragraph> paragraphsInRange = new();
        bool inside = false;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para == startParagraph)
                inside = true;

            if (inside)
                paragraphsInRange.Add(para);

            if (para == endParagraph)
                break;
        }

        // Remove empty paragraphs.
        foreach (Paragraph para in paragraphsInRange)
        {
            // GetText() returns the paragraph text plus a paragraph break.
            string text = para.GetText().Trim();
            if (string.IsNullOrEmpty(text))
                para.Remove();
        }
    }
}
