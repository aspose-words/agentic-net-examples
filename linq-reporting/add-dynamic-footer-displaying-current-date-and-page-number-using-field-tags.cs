using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model – no fields are required for this example.
    public class ReportModel { }

    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "ReportWithFooter.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Ensure the document has at least one section.
        builder.MoveToSection(0);

        // Create a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a page number field (current page).
        builder.InsertField("PAGE");

        // Add static text separator.
        builder.Write(" of ");

        // Insert a total pages field.
        builder.InsertField("NUMPAGES");

        // Add a separator before the date.
        builder.Write(" - ");

        // Insert the current date field with a custom format.
        builder.InsertField(@"DATE \@ ""MMMM d, yyyy""");

        // Save the template to disk.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Configure the reporting engine to update Word fields after the report is built.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.UpdateFieldsSyntaxAware
        };

        // Build the report using an empty data source (the model has no members).
        engine.BuildReport(doc, new ReportModel(), "model");

        // -----------------------------------------------------------------
        // 3. Save the final document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
