using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Build a table where the first two cells of the first row are merged horizontally.
        // The <<cellMerge>> tag tells the LINQ Reporting engine to merge cells that have
        // identical textual content (ignoring surrounding whitespace).
        Table table = builder.StartTable();

        // First cell – contains the merge tag and the text "Group A".
        builder.InsertCell();
        builder.Write("<<cellMerge>>Group A");

        // Second cell – same merge tag and identical text.
        builder.InsertCell();
        builder.Write("<<cellMerge>>Group A");

        // End the first row.
        builder.EndRow();

        // Add a normal (unmerged) row for demonstration.
        builder.InsertCell();
        builder.Write("Item 1");
        builder.InsertCell();
        builder.Write("Description 1");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // The model can be empty because the template does not reference any data fields.
        var model = new ReportModel();

        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(reportPath);
    }

    // Simple wrapper class used as the root data source.
    public class ReportModel
    {
        // Add properties here if the template needs to reference data.
    }
}
