using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a sample DataTable with some columns and rows.
        var dataTable = new DataTable("Data");
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Price", typeof(double));

        dataTable.Rows.Add(1, "Apple", 0.5);
        dataTable.Rows.Add(2, "Banana", 0.3);
        dataTable.Rows.Add(3, "Cherry", 0.8);

        // Create a Word document template programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln("Product List:");

        // Define a data band that iterates over the rows of the DataTable.
        builder.Writeln("<<foreach [row in Data]>>");
        // Output each column value for the current row.
        builder.Writeln("Id: <<[row.Id]>>\tName: <<[row.Name]>>\tPrice: <<[row.Price]>>");
        // Close the data band.
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, dataTable, "Data");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
