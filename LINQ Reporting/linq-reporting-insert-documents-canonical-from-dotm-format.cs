using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTM template that contains reporting tags.
        Document template = new Document("Template.dotm");

        // Create a simple data source (DataTable) for the report.
        DataTable data = new DataTable("Employees");
        data.Columns.Add("Name");
        data.Columns.Add("Position");
        data.Rows.Add("John Doe", "Manager");
        data.Rows.Add("Jane Smith", "Developer");

        // Populate the template with data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data, "ds");

        // Load an additional document that we want to insert into the report.
        Document appendix = new Document("Appendix.docx");

        // Move the cursor to the end of the populated report.
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.MoveToDocumentEnd();

        // Insert the appendix document, preserving its original formatting.
        builder.InsertDocument(appendix, Aspose.Words.ImportFormatMode.KeepSourceFormatting);

        // Save the final merged document.
        template.Save("Result.docx", SaveFormat.Docx);
    }
}
