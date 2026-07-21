using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a data band that iterates over the rows of a DataTable named "Data".
        builder.Writeln("<<foreach [row in Data]>>");
        // Output each column value for the current row.
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data in a DataTable.
        DataTable table = new DataTable("Data");
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));

        table.Rows.Add(1, "Alice");
        table.Rows.Add(2, "Bob");
        table.Rows.Add(3, "Charlie");

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, table, "Data");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
