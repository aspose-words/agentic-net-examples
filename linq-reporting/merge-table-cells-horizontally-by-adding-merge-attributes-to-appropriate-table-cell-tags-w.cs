using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Build a simple 2‑column table.
        // The first row will have horizontally merged cells using the <<cellMerge>> tag.
        builder.StartTable();

        // First cell – contains the merge tag and the text that will be shared.
        builder.InsertCell();
        builder.Write("<<cellMerge>>Group A");

        // Second cell – same merge tag and identical text.
        builder.InsertCell();
        builder.Write("<<cellMerge>>Group A");

        // End the first row.
        builder.EndRow();

        // Add a normal row to demonstrate that only the first row is merged.
        builder.InsertCell();
        builder.Write("Item 1");
        builder.InsertCell();
        builder.Write("Item 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the template to disk.
        template.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report.
        // -------------------------------------------------
        Document report = new Document(templatePath);

        // The template does not reference any data, so an empty object is sufficient.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, new object());

        // -------------------------------------------------
        // 3. Save the generated report.
        // -------------------------------------------------
        report.Save(outputPath);

        // Inform the user (optional, no interactive wait).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
