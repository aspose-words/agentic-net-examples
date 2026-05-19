using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final PDF.
        const string templatePath = "template.docx";
        const string outputPdfPath = "Report.pdf";

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Sample Report ===");
        // Begin a foreach loop over the DataTable named "Orders" inside the DataSet "ds".
        builder.Writeln("<<foreach [order in ds.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Product : <<[order.Product]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Build a DataSet that matches the template structure.
        // -----------------------------------------------------------------
        DataSet ds = new DataSet();

        DataTable ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("CustomerName", typeof(string));
        ordersTable.Columns.Add("Product", typeof(string));
        ordersTable.Columns.Add("Quantity", typeof(int));

        ordersTable.Rows.Add("Alice Johnson", "Laptop", 2);
        ordersTable.Rows.Add("Bob Smith", "Smartphone", 5);
        ordersTable.Rows.Add("Carol White", "Tablet", 3);

        ds.Tables.Add(ordersTable);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The root name "ds" must match the name used in the template tags.
        engine.BuildReport(reportDoc, ds, "ds");

        // -----------------------------------------------------------------
        // 4. Save the generated report as PDF.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
