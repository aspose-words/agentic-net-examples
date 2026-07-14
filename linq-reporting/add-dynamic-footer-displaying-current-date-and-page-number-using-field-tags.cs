using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary footer (appears on all pages except first/odd/even if not set).
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a DATE field that will display the current date.
        builder.InsertField("DATE");

        // Add a separator and a PAGE field that will display the current page number.
        builder.Write(" - Page ");
        builder.InsertField("PAGE");

        // The report does not need any data, but the LINQ Reporting engine still requires a data source.
        ReportModel model = new ReportModel();

        // Build the report – this will simply copy the template as‑is because there are no tags to process.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("DynamicFooterReport.docx");
    }

    // Empty model class required by the ReportingEngine.
    public class ReportModel
    {
    }
}
