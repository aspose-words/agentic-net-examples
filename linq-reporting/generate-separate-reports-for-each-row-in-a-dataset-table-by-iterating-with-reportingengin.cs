using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a DataSet with a single DataTable containing sample rows.
        DataSet dataSet = new DataSet();
        DataTable ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("OrderId", typeof(int));
        ordersTable.Columns.Add("CustomerName", typeof(string));
        ordersTable.Columns.Add("Total", typeof(decimal));
        ordersTable.Rows.Add(1, "Alice", 123.45m);
        ordersTable.Rows.Add(2, "Bob", 678.90m);
        ordersTable.Rows.Add(3, "Charlie", 250.00m);
        dataSet.Tables.Add(ordersTable);

        // Path for the template document.
        const string templatePath = "Template.docx";

        // Create the template with LINQ Reporting tags.
        CreateTemplate(templatePath);

        // Load the template once.
        Document template = new Document(templatePath);

        // Ensure the output directory exists.
        const string outputDir = "Reports";
        Directory.CreateDirectory(outputDir);

        // Iterate over each DataRow and generate a separate report.
        foreach (DataRow row in ordersTable.Rows)
        {
            // Clone the template for the current row.
            Document report = (Document)template.Clone(true);

            // Build the report using the DataRow as the root object named "row".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, row, "row");

            // Save the generated report with a distinct file name.
            string fileName = $"Report_Row_{row["OrderId"]}.docx";
            string outputPath = Path.Combine(outputDir, fileName);
            report.Save(outputPath);
        }
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("Order ID: <<[row.OrderId]>>");
        builder.Writeln("Customer: <<[row.CustomerName]>>");
        builder.Writeln("Total Amount: $<<[row.Total]>>");

        doc.Save(path);
    }
}
