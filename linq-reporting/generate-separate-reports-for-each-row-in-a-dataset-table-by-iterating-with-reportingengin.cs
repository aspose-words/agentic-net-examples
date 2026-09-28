using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample DataSet with one DataTable.
        DataSet dataSet = new DataSet();
        DataTable table = new DataTable("SampleTable");
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));

        table.Rows.Add(1, "Alice", 123.45m);
        table.Rows.Add(2, "Bob", 678.90m);
        table.Rows.Add(3, "Charlie", 250.00m);

        dataSet.Tables.Add(table);

        // Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report for a single row");
        builder.Writeln("Id: <<[row.Id]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Amount: <<[row.Amount]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Prepare the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Iterate over each DataRow and generate a separate report.
        foreach (DataRow dataRow in table.Rows)
        {
            // Load a fresh copy of the template for each iteration.
            Document reportDoc = new Document(templatePath);

            // Build the report using the current DataRow as the root object.
            engine.BuildReport(reportDoc, dataRow, "row");

            // Save the generated report with a unique file name.
            string outputPath = $"Report_{dataRow["Id"]}.docx";
            reportDoc.Save(outputPath);
        }
    }
}
