using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

public class ReportModel
{
    // No data needed for the footer.
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a new document and add a primary footer with DATE and PAGE fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Move to the primary footer of the first (and only) section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Generated on ");
        builder.InsertField(FieldType.FieldDate, true);
        builder.Write(" - Page ");
        builder.InsertField(FieldType.FieldPage, true);

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine (no data needed).
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, new ReportModel(), "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
