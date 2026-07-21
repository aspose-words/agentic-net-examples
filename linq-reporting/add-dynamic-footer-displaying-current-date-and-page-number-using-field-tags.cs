using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report
        string templatePath = "template.docx";
        string reportPath = "report.docx";

        // -----------------------------------------------------------------
        // Create a template document with a footer that contains the date,
        // current page number and total page count.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Move to the primary footer
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert the date field
        builder.Write("Date: ");
        builder.InsertField("DATE \\@ \"MMMM d, yyyy\"");

        // Insert page number fields
        builder.Write("   Page ");
        builder.InsertField("PAGE");
        builder.Write(" of ");
        builder.InsertField("NUMPAGES");

        // Save the template
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the final report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Dummy model – not used in this example but required by the engine
        var model = new ReportModel();

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report (no data source needed for the footer fields)
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        reportDoc.Save(reportPath);
    }
}

// Dummy model class required by the ReportingEngine (no properties needed for this scenario)
public class ReportModel
{
    // Parameterless constructor to avoid nullable warnings
    public ReportModel() { }
}
