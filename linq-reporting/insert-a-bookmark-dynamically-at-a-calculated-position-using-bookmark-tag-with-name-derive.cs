using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Title = "Quarterly Sales",
            // Bookmark name is derived from the title.
            BookmarkName = "BM_" + "Quarterly_Sales"
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a simple title placeholder.
        builder.Writeln("Report Title: <<[model.Title]>>");

        // Insert a bookmark tag whose name comes from the data model.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("This paragraph is inside the dynamically named bookmark.");
        builder.Writeln("<</bookmark>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        report.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Public data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Title displayed in the report.
    public string Title { get; set; } = string.Empty;

    // Name of the bookmark; must be non‑empty.
    public string BookmarkName { get; set; } = string.Empty;
}
