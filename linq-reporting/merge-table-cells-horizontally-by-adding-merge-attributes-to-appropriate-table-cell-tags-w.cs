using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

namespace AsposeWordsLinqReportingMergeCells
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Title displayed in the report.
        public string Title { get; set; } = "Sample Report";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string reportPath   = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a title that will be filled from the data model.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // Create a table where the first two cells of the first row will be merged horizontally.
            builder.Writeln("Table with horizontally merged cells:");
            Table table = builder.StartTable();

            // First row – cells to be merged.
            builder.InsertCell();
            // The <<cellMerge>> tag tells the LINQ Reporting engine to merge this cell horizontally.
            builder.Writeln("<<cellMerge>>Group");

            builder.InsertCell();
            builder.Writeln("<<cellMerge>>Group"); // Same content and tag as the previous cell.

            builder.EndRow();

            // Second row – regular cells (no merging).
            builder.InsertCell();
            builder.Writeln("Cell 1");

            builder.InsertCell();
            builder.Writeln("Cell 2");

            builder.EndRow();
            builder.EndTable();

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the model as the root data source.
            engine.BuildReport(reportDoc, new ReportModel(), "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
