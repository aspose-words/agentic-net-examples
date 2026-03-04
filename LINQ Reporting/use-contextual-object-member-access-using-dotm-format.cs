using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class DotmMemberAccessExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DOTM (mail‑merge) template that accesses nested members.
        // The syntax <<[Customer.Name]>> will be replaced by the value of the Name column
        // of the Customer table. If the column is missing, the MissingMemberMessage will be used.
        builder.Writeln("Customer name: <<[Customer.Name]>>");
        builder.Writeln("Customer age: <<[Customer.Age]>>");
        builder.Writeln("Missing field test: <<[Customer.NonExistent]>>");

        // Prepare a DataSet with a DataTable named "Customer".
        DataSet data = new DataSet();
        DataTable table = new DataTable("Customer");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        // Populate the table with a single row.
        table.Rows.Add("John Doe", 42);
        data.Tables.Add(table);

        // Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            // Allow the engine to continue when a member is missing.
            Options = ReportBuildOptions.AllowMissingMembers,
            // Text to display for missing members.
            MissingMemberMessage = "[Missing]"
        };

        // Build the report using the template and the data source.
        engine.BuildReport(doc, data, string.Empty);

        // Save the resulting document.
        doc.Save("DotmMemberAccessResult.docx");
    }
}
