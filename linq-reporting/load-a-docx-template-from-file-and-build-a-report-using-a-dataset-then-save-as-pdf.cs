using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final PDF.
        const string templatePath = "Template.docx";
        const string outputPdfPath = "Report.pdf";

        // -----------------------------------------------------------------
        // 1. Create a DOCX template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Begin the data band.
        builder.Writeln("<<foreach [row in Orders]>>");

        // Start the table that will be repeated for each row.
        Table table = builder.StartTable();

        // Header row (once, outside the foreach loop).
        builder.InsertCell();
        builder.Writeln("Customer");
        builder.InsertCell();
        builder.Writeln("Date");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.EndRow();

        // Data row (inside the foreach loop).
        builder.InsertCell();
        builder.Writeln("<<[row.CustomerName]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.OrderDate]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Amount]>>");
        builder.EndRow();

        // Close the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare a DataSet with sample data.
        // -----------------------------------------------------------------
        DataSet dataSet = new DataSet();

        DataTable ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("CustomerName", typeof(string));
        ordersTable.Columns.Add("OrderDate", typeof(DateTime));
        ordersTable.Columns.Add("Amount", typeof(decimal));

        ordersTable.Rows.Add("Alice Johnson", new DateTime(2023, 5, 12), 250.75m);
        ordersTable.Rows.Add("Bob Smith", new DateTime(2023, 5, 13), 120.00m);
        ordersTable.Rows.Add("Carol White", new DateTime(2023, 5, 14), 560.40m);

        dataSet.Tables.Add(ordersTable);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.BuildReport(reportDoc, dataSet, "");

        // -----------------------------------------------------------------
        // 4. Save the generated report as PDF.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
