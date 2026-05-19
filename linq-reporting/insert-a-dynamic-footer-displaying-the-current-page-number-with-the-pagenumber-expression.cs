using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

public class ReportModel
{
    // No data members are required for this example.
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Ensure the document has a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a PAGE field that automatically shows the current page number.
        builder.InsertField(FieldType.FieldPage, true);

        // Add a couple of pages so the footer can be seen.
        builder.MoveToSection(0);
        builder.Writeln("Content page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content page 2");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // Create a ReportingEngine instance.
        ReportingEngine engine = new ReportingEngine();

        // Build the report using an empty data source (the footer does not depend on data).
        engine.BuildReport(doc, new ReportModel(), "model");

        // -------------------------------------------------
        // 3. Save the generated report.
        // -------------------------------------------------
        doc.Save(reportPath);
    }
}
