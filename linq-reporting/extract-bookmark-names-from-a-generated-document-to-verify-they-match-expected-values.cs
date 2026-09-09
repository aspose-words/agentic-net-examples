using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define file paths in the working directory.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document that contains LINQ Reporting bookmark tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // First bookmark tag.
        builder.Writeln("<<bookmark [model.Bookmark1]>>");
        builder.Writeln("Content of first bookmark.");
        builder.Writeln("<</bookmark>>");

        // Second bookmark tag.
        builder.Writeln("<<bookmark [model.Bookmark2]>>");
        builder.Writeln("Content of second bookmark.");
        builder.Writeln("<</bookmark>>");

        // Save the template to disk.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model that supplies bookmark names.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Bookmark1 = "BM_First",
            Bookmark2 = "BM_Second"
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using ReportingEngine.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        report.Save(reportPath);

        // -----------------------------------------------------------------
        // 4. Extract bookmark names from the generated document.
        // -----------------------------------------------------------------
        List<string> actualBookmarkNames = report.Range.Bookmarks
                                                   .Select(b => b.Name)
                                                   .ToList();

        // Expected bookmark names based on the model.
        List<string> expectedBookmarkNames = new List<string> { model.Bookmark1, model.Bookmark2 };

        // -----------------------------------------------------------------
        // 5. Verify that the extracted names match the expected ones.
        // -----------------------------------------------------------------
        bool match = actualBookmarkNames.SequenceEqual(expectedBookmarkNames);

        Console.WriteLine($"Bookmark verification result: {(match ? "Success" : "Failure")}");
        Console.WriteLine("Expected bookmarks: " + string.Join(", ", expectedBookmarkNames));
        Console.WriteLine("Actual bookmarks:   " + string.Join(", ", actualBookmarkNames));
    }
}

// Data model used by the LINQ Reporting engine.
// All properties are initialized to avoid nullable warnings.
public class ReportModel
{
    public string Bookmark1 { get; set; } = string.Empty;
    public string Bookmark2 { get; set; } = string.Empty;
}
