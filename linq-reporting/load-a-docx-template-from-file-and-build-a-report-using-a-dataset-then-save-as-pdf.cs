using System;
using System.Data;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample DataSet with an Orders table
        var dataSet = new DataSet();
        var ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("OrderID", typeof(int));
        ordersTable.Columns.Add("CustomerName", typeof(string));
        ordersTable.Columns.Add("Amount", typeof(decimal));
        ordersTable.Rows.Add(1, "Alice", 120.50m);
        ordersTable.Rows.Add(2, "Bob", 75.00m);
        ordersTable.Rows.Add(3, "Charlie", 200.00m);
        dataSet.Tables.Add(ordersTable);

        // Create a DOCX template with LINQ Reporting tags
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln(); // blank line

        // Begin foreach loop over Orders table (use root name "data")
        builder.Writeln("<<foreach [order in data.Orders]>>");

        // Start table inside the foreach block
        Table table = builder.StartTable();

        // Header row (only once, before data rows)
        builder.InsertCell();
        builder.Writeln("Order ID");
        builder.InsertCell();
        builder.Writeln("Customer");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.EndRow();

        // Data row (repeated for each order)
        builder.InsertCell();
        builder.Writeln("<<[order.OrderID]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.CustomerName]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Amount]>>");
        builder.EndRow();

        // End table and foreach block
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // Load the template for report generation
        var reportDoc = new Document(templatePath);

        // Build the report using the DataSet as the root object named "data"
        var engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, dataSet, "data");

        // Save the generated report as PDF if the build succeeded
        if (success)
        {
            reportDoc.Save("Report.pdf");
        }
    }
}
