using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string BookmarkName { get; set; } = "SampleBookmark";
    public string Title { get; set; } = "Hello World";
}

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // 1. Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a bookmark tag that will be filled from the model.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template for report generation.
        Document doc = new Document(templatePath);

        // 3. Prepare the data model.
        Model model = new Model
        {
            BookmarkName = "MyBookmark",
            Title = "Generated Title"
        };

        // 4. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated document.
        doc.Save(outputPath);

        // 6. Load the generated document and extract bookmark names.
        Document resultDoc = new Document(outputPath);
        List<string> extractedBookmarks = new();
        foreach (Bookmark bm in resultDoc.Range.Bookmarks)
        {
            extractedBookmarks.Add(bm.Name);
        }

        // 7. Verify the extracted bookmark names against expected values.
        List<string> expectedBookmarks = new() { "MyBookmark" };
        bool match = extractedBookmarks.SequenceEqual(expectedBookmarks);

        // 8. Output verification result.
        Console.WriteLine(match
            ? "All bookmarks match expected values."
            : $"Bookmark mismatch. Expected: [{string.Join(", ", expectedBookmarks)}], Found: [{string.Join(", ", extractedBookmarks)}]");
    }
}
