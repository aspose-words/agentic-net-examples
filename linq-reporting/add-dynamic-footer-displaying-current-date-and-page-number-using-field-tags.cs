using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // No data needed for this example.
    public ReportModel() { }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a footer with dynamic date and page number using Word field codes.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Generated on: ");
        builder.InsertField(@"DATE \@ ""MMMM d, yyyy"""); // Current date.
        builder.Write("   Page: ");
        builder.InsertField("PAGE");          // Current page number.
        builder.Write(" of ");
        builder.InsertField("NUMPAGES");      // Total pages.

        // Return to the main body and add sample content to generate multiple pages.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is a sample report generated with Aspose.Words LINQ Reporting.");
        builder.Writeln("The footer above displays the current date and page number dynamically.");
        // Insert enough text to create a second page.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content. The footer should continue to show the correct page number.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        ReportModel model = new ReportModel();

        ReportingEngine engine = new ReportingEngine();
        // No special options needed for this simple example.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
