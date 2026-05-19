using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingExample
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a simple LINQ Reporting template.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The tags reference column names of the DataRow that will be supplied later.
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[CustomerName]>>");
        builder.Writeln("Address: <<[Address]>>");
        builder.Writeln("Phone: <<[Phone]>>");

        // Save the template so it can be loaded for each iteration.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Build a DataSet with sample data.
        // -----------------------------------------------------------------
        DataSet dataSet = new DataSet();
        DataTable table = new DataTable("Customers");
        table.Columns.Add("CustomerName", typeof(string));
        table.Columns.Add("Address", typeof(string));
        table.Columns.Add("Phone", typeof(string));

        table.Rows.Add("Alice Johnson", "123 Maple St., Springfield", "555-0101");
        table.Rows.Add("Bob Smith", "456 Oak Ave., Shelbyville", "555-0202");
        table.Rows.Add("Carol White", "789 Pine Rd., Capital City", "555-0303");

        dataSet.Tables.Add(table);

        // -----------------------------------------------------------------
        // 3. Iterate over each DataRow and generate a separate report.
        // -----------------------------------------------------------------
        int index = 1;
        foreach (DataRow row in table.Rows)
        {
            // Load a fresh copy of the template for each report.
            Document reportDoc = new Document(templatePath);

            // Use the ReportingEngine to populate the template with the current row.
            ReportingEngine engine = new ReportingEngine();
            // BuildReport overload that accepts a DataRow does not require a data source name.
            engine.BuildReport(reportDoc, row);

            // Save the generated report.
            string outputPath = Path.Combine(outputDir, $"Report_{index}.docx");
            reportDoc.Save(outputPath);
            index++;
        }
    }
}
