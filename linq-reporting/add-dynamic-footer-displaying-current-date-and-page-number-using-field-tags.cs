using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ReportModel
{
    // Current date formatted as a short date string.
    public string CurrentDate { get; set; } = DateTime.Now.ToString("d");
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a simple body paragraph (optional, just to have content).
        builder.Writeln("This is a sample report generated with Aspose.Words LINQ Reporting.");

        // Move the cursor to the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert the current date using a LINQ Reporting expression tag.
        builder.Write("Date: ");
        builder.Writeln("<<[model.CurrentDate]>>");

        // Insert the page number using a Word field.
        builder.Write("Page ");
        builder.InsertField("PAGE \\* MERGEFORMAT");
        builder.Writeln();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options

        // Build the report using the model instance.
        engine.BuildReport(reportDoc, new ReportModel(), "model");

        // Save the final report.
        reportDoc.Save(reportPath);
    }
}
