using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingExample
{
    public static void Main()
    {
        // Create a simple DOCX template with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount: $<<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Prepare a DataSet with sample data.
        DataSet dataSet = new DataSet();

        DataTable ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("OrderId", typeof(int));
        ordersTable.Columns.Add("CustomerName", typeof(string));
        ordersTable.Columns.Add("Amount", typeof(decimal));

        ordersTable.Rows.Add(1001, "Alice Johnson", 250.75m);
        ordersTable.Rows.Add(1002, "Bob Smith", 120.00m);
        ordersTable.Rows.Add(1003, "Carol Davis", 560.40m);

        dataSet.Tables.Add(ordersTable);

        // Load the template document from file.
        Document doc = new Document(templatePath);

        // Build the report using the DataSet as the root data source.
        ReportingEngine engine = new ReportingEngine();
        // Root name "ds" matches the name used in the template tags (Orders table is accessed directly).
        engine.BuildReport(doc, dataSet, "ds");

        // Save the generated report as PDF.
        const string outputPath = "Report.pdf";
        doc.Save(outputPath);
    }
}
