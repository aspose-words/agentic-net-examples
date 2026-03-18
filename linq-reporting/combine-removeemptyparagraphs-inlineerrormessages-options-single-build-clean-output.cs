using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportBuilderExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some report tags that may produce empty paragraphs or syntax errors.
        // Example: a tag that will be removed because the data source does not contain the field.
        builder.Writeln("<<[Customer.Name]>>");
        builder.Writeln("<<[MissingField]>>"); // This will cause a syntax error.

        // Prepare a simple data source.
        DataSet data = new DataSet();
        DataTable table = new DataTable("Customer");
        table.Columns.Add("Name");
        table.Rows.Add("John Doe");
        data.Tables.Add(table);

        // Configure the ReportingEngine to remove empty paragraphs and inline error messages.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages,
            MissingMemberMessage = "Field not found"
        };

        // Build the report using the document template and the data set.
        engine.BuildReport(doc, data, string.Empty);

        // Save the resulting document.
        doc.Save("CombinedReport.docx");
    }
}
