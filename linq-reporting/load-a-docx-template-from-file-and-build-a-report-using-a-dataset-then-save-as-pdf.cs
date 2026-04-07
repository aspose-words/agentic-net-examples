using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class LinqReportingExample
{
    public static void Main()
    {
        // Paths for the template and the final PDF.
        const string templatePath = "Template.docx";
        const string pdfPath = "Report.pdf";

        // -----------------------------------------------------------------
        // 1. Create a DOCX template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Orders Report ===");
        // Iterate over the DataTable named 'Orders' inside the DataSet (root name: ds).
        builder.Writeln("<<foreach [order in ds.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
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
        ordersTable.Rows.Add("Carol Davis", new DateTime(2023, 5, 14), 560.40m);

        dataSet.Tables.Add(ordersTable);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using the DataSet.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        // The root name used in the template is "ds".
        engine.BuildReport(reportDoc, dataSet, "ds");

        // -----------------------------------------------------------------
        // 4. Save the populated document as PDF.
        // -----------------------------------------------------------------
        reportDoc.Save(pdfPath, SaveFormat.Pdf);
    }
}
