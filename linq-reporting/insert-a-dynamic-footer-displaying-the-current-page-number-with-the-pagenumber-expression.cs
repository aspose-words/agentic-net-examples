using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // No properties needed for this simple example.
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a footer with a dynamic page number field.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Page ");
        builder.InsertField("{=PageNumber}");
        builder.Write(" of ");
        builder.InsertField("{=NUMPAGES}");

        // Add sample content to generate multiple pages.
        builder.MoveToDocumentStart();
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"This is sample paragraph {i + 1}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        }

        // Save the template to disk.
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare a dummy root model.
        ReportModel model = new ReportModel();

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);
    }
}
