using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportingEngineFallbackExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a template that references a missing member (Age) of the data object.
        // The expression syntax "<<[object.member]>>" is used by ReportingEngine.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age : <<[person.Age]>>"); // Age does not exist in the data source.

        // Prepare a data source that contains only the "Name" field.
        DataSet data = new DataSet();
        DataTable table = new DataTable("person");
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("John Doe");
        data.Tables.Add(table);

        // Configure the ReportingEngine to allow missing members and provide a fallback text.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "N/A" // This text will be printed for any missing member.
        };

        // Build the report – the missing Age member will be replaced with "N/A".
        engine.BuildReport(builder.Document, data, string.Empty);

        // Save the resulting document.
        doc.Save("ReportingEngineFallback.docx");
    }
}
