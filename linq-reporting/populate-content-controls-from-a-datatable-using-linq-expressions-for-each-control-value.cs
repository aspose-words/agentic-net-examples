using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data source.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));

        table.Rows.Add(1, "Alice", 30);
        table.Rows.Add(2, "Bob", 25);
        table.Rows.Add(3, "Charlie", 35);

        // Create the template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Heading.
        builder.Writeln("Employee Report");
        builder.Writeln();

        // Begin foreach loop over the DataTable rows.
        builder.Writeln("<<foreach [row in Data]>>");

        // Table that will be repeated for each row.
        Table tbl = builder.StartTable();

        // Id column.
        builder.InsertCell();
        builder.Write("<<[row.Id]>>");

        // Name column.
        builder.InsertCell();
        builder.Write("<<[row.Name]>>");

        // Age column.
        builder.InsertCell();
        builder.Write("<<[row.Age]>>");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the DataTable as the root data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, table, "Data");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}
