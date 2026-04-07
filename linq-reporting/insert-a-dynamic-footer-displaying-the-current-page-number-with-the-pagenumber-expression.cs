using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Reporting;

public class ReportModel
{
    // No members are required for this example.
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Move the cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert static text and a PAGE field that displays the current page number.
        builder.Write("Page ");
        builder.InsertField(FieldType.FieldPage, true); // Inserts a PAGE field.
        builder.Writeln(); // End the paragraph.

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. No data members are used, but we still provide a root object and name.
        engine.BuildReport(loadedTemplate, new ReportModel(), "model");

        // -------------------------------------------------
        // 3. Save the final document.
        // -------------------------------------------------
        loadedTemplate.Save(outputPath);
    }
}
