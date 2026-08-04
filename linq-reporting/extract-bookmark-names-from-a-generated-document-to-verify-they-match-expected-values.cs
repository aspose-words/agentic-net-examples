using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model with expected bookmark names.
        var model = new ReportModel
        {
            Bookmarks = new List<string> { "FirstBookmark", "SecondBookmark", "ThirdBookmark" }
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // LINQ Reporting tags: iterate over the Bookmarks collection and create a bookmark for each name.
        builder.Writeln("<<foreach [b in Bookmarks]>>");
        builder.Writeln("<<bookmark [b]>>");
        builder.Writeln("<<[b]>>"); // Content inside the bookmark.
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model);

        // Save the generated document.
        const string outputPath = "GeneratedReport.docx";
        reportDoc.Save(outputPath);

        // Extract bookmark names from the generated document.
        List<string> actualBookmarkNames = reportDoc.Range.Bookmarks
            .Select(b => b.Name)
            .ToList();

        // Verify that the extracted bookmark names match the expected ones.
        bool areEqual = actualBookmarkNames.SequenceEqual(model.Bookmarks);
        Console.WriteLine(areEqual
            ? "Success: Bookmark names match expected values."
            : "Failure: Bookmark names do not match expected values.");

        // Optionally, list the extracted bookmark names.
        Console.WriteLine("Extracted bookmark names:");
        foreach (var name in actualBookmarkNames)
        {
            Console.WriteLine($"- {name}");
        }
    }
}

// Public data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Collection of bookmark names to be generated.
    public List<string> Bookmarks { get; set; } = new();
}
