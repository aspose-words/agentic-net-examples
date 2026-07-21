using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingExample
{
    public static void Main()
    {
        // Prepare sample data in a DataSet with one DataTable.
        DataSet dataSet = new DataSet();
        DataTable table = new DataTable("Customers");
        table.Columns.Add("CustomerName", typeof(string));
        table.Columns.Add("Address", typeof(string));
        table.Rows.Add("Alice Johnson", "123 Maple Street");
        table.Rows.Add("Bob Smith", "456 Oak Avenue");
        table.Rows.Add("Carol White", "789 Pine Road");
        dataSet.Tables.Add(table);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[CustomerName]>>");
        builder.Writeln("Address: <<[Address]>>");
        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Iterate over each DataRow and generate a separate report.
        int index = 0;
        foreach (DataRow row in table.Rows)
        {
            // Load the template for each iteration.
            Document report = new Document(templatePath);
            // Build the report using the current DataRow as the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, row);
            // Save the generated report.
            string outputPath = $"Report_{index}.docx";
            report.Save(outputPath);
            index++;
        }
    }
}
