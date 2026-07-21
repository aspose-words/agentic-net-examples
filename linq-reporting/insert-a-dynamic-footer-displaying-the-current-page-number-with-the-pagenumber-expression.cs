using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Folder for generated files
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a template document with a footer that shows the
        //         current page number (e.g., "1 of 3").
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Move the cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Center the page number text.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert Word fields for the current page number and total pages.
        builder.InsertField(FieldType.FieldPage, true);          // PAGE field
        builder.Write(" of ");
        builder.InsertField(FieldType.FieldNumPages, true);     // NUMPAGES field

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // No data source is required for this simple example; an empty object is sufficient.
        engine.BuildReport(report, new object());

        // -----------------------------------------------------------------
        // Step 3: Save the final document.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "ReportWithPageNumbers.docx");
        report.Save(resultPath);
    }
}
