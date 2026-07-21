using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags that reference the DataSet root object named "ds".
        // The root object will be a DataRow, so the fields can be accessed directly.
        builder.Writeln("Report Title: <<[ds.Title]>>");
        builder.Writeln("Report Author: <<[ds.Author]>>");
        builder.Writeln("Report Date: <<[ds.ReportDate]>>");

        // Save the template to a temporary file.
        const string templatePath = "ReportTemplate.docx";
        template.Save(templatePath);

        // Prepare a DataSet with a single DataTable containing the data.
        DataSet ds = new DataSet();
        DataTable table = new DataTable("ReportData");
        table.Columns.Add("Title", typeof(string));
        table.Columns.Add("Author", typeof(string));
        table.Columns.Add("ReportDate", typeof(DateTime));

        // Add one row of sample data.
        table.Rows.Add("Quarterly Sales Summary", "Jane Doe", DateTime.Today);
        ds.Tables.Add(table);

        // Load the template document (demonstrates load step).
        Document doc = new Document(templatePath);

        // Use the first row of the DataTable as the data source for the report.
        DataRow row = ds.Tables["ReportData"].Rows[0];

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, row, "ds");

        // After the report is generated, set custom document properties based on the DataSet values.
        doc.CustomDocumentProperties.Add("ReportTitle", row["Title"].ToString());
        doc.CustomDocumentProperties.Add("ReportAuthor", row["Author"].ToString());
        doc.CustomDocumentProperties.Add(
            "ReportGeneratedOn",
            ((DateTime)row["ReportDate"]).ToString("yyyy-MM-dd"));

        // Save the final report.
        const string outputPath = "GeneratedReport.docx";
        doc.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated and saved to '{outputPath}'.");
    }
}
