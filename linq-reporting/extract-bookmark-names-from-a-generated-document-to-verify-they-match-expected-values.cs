using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string BookmarkName { get; set; } = "";
    public string Title { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // 1. Create a template document with a bookmark tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        Document doc = new Document(templatePath);

        // 3. Prepare the data model.
        ReportModel model = new ReportModel
        {
            BookmarkName = "SampleBookmark",
            Title = "Hello World"
        };

        // 4. Build the report using LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated document.
        doc.Save(outputPath);

        // 6. Extract bookmark names from the generated document.
        List<string> actualBookmarks = doc.Range.Bookmarks
            .Select(b => b.Name)
            .ToList();

        // 7. Define expected bookmark names.
        List<string> expectedBookmarks = new() { "SampleBookmark" };

        // 8. Verify that the actual bookmarks match the expected ones.
        bool match = actualBookmarks.SequenceEqual(expectedBookmarks);
        Console.WriteLine($"Bookmarks match: {match}");

        // Optional: print the extracted bookmark names.
        Console.WriteLine("Extracted bookmarks:");
        foreach (string name in actualBookmarks)
        {
            Console.WriteLine($"- {name}");
        }
    }
}
